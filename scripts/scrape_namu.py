"""
나무위키 여론조사 페이지 스크레이퍼 (Playwright 기반)
====================================================
매일 GitHub Actions 에서 실행되어 새 여론조사를 감지하면
polls.json 을 갱신하고, 워크플로우가 PR 을 자동 생성한다.

전략:
  1) Playwright 로 진짜 브라우저 띄워 Cloudflare 우회
  2) 표를 파싱해 (date, pollster, n, dem%, pp%) 추출
  3) 기존 polls.json 과 키 (city, date, pollster) 로 비교 → diff
  4) 새 항목 있으면 polls.json 갱신, GITHUB_OUTPUT 으로 has_new=true
  5) 파싱 실패하면 페이지 hash 만 비교해 변경 감지 신호

실패해도 워크플로우가 죽지 않도록 try/except 로 감싼다.
"""

import hashlib
import json
import os
import re
import sys
from datetime import datetime, timezone, timedelta
from pathlib import Path

KST = timezone(timedelta(hours=9))
ROOT = Path(__file__).resolve().parent.parent
POLLS_PATH = ROOT / "polls.json"
HASH_PATH = ROOT / ".namu_hash.json"

# 도시별 나무위키 URL + 후보 이름 (파서가 표 식별에 사용)
TARGETS = [
    {
        "city": "서울시장",
        "url": "https://namu.wiki/w/제9회 전국동시지방선거/여론조사/서울특별시",
        "cand_dem": "정원오",
        "cand_pp": "오세훈",
    },
    {
        "city": "부산시장",
        "url": "https://namu.wiki/w/제9회 전국동시지방선거/여론조사/부산광역시",
        "cand_dem": "전재수",
        "cand_pp": "박형준",
    },
    # 성남은 별도 페이지가 분명치 않아 일단 제외
    # 필요 시 추가: "https://namu.wiki/w/제9회 전국동시지방선거/여론조사/경기도"
]


def fetch_html(url, timeout=90000):
    """Playwright + stealth 로 Cloudflare 봇 검증 우회 시도."""
    from playwright.sync_api import sync_playwright
    try:
        from tf_playwright_stealth import stealth_sync
        has_stealth = True
    except ImportError:
        has_stealth = False
        print(f"    [fetch] (stealth 미설치 — 일반 모드)")

    print(f"    [fetch] launching chromium...")
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--disable-features=IsolateOrigins,site-per-process",
                "--no-sandbox",
            ],
        )
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/131.0.0.0 Safari/537.36",
            locale="ko-KR",
            timezone_id="Asia/Seoul",
            viewport={"width": 1366, "height": 800},
            extra_http_headers={
                "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8",
            },
        )
        page = context.new_page()
        if has_stealth:
            stealth_sync(page)

        print(f"    [fetch] goto {url[:80]}...")
        resp = page.goto(url, wait_until="domcontentloaded", timeout=timeout)
        status = resp.status if resp else 0
        print(f"    [fetch] HTTP {status}, Cloudflare 챌린지 대기 중...")

        # Cloudflare 챌린지: 실제 콘텐츠 셀렉터가 나타날 때까지 최대 30초 대기
        cf_passed = False
        for i in range(30):
            page.wait_for_timeout(1000)
            title = page.title()
            if "잠시만" not in title and "Just a moment" not in title and "Cloudflare" not in title:
                cf_passed = True
                print(f"    [fetch] Cloudflare 통과 ({i+1}초). title={title!r}")
                break
        if not cf_passed:
            print(f"    [fetch] ⚠ Cloudflare 챌린지 미통과 (30초 timeout). title={page.title()!r}")

        # 컨텐츠 안정화
        try:
            page.wait_for_load_state("networkidle", timeout=10000)
        except Exception:
            pass

        html = page.content()
        browser.close()
        print(f"    [fetch] done. html_len={len(html):,}")
        return html


def save_debug(city, html):
    """디버그 위해 HTML 일부 + 테이블 개수 저장 (workflow artifact 로 업로드)"""
    from bs4 import BeautifulSoup
    debug_dir = ROOT / "debug"
    debug_dir.mkdir(exist_ok=True)

    soup = BeautifulSoup(html, "lxml")
    n_tables = len(soup.find_all("table"))

    safe_name = city.replace("/", "_")
    (debug_dir / f"{safe_name}.html").write_text(html, encoding="utf-8")
    summary = (
        f"city: {city}\n"
        f"html_len: {len(html):,}\n"
        f"tables_found: {n_tables}\n"
        f"title: {soup.title.string if soup.title else '(no title)'}\n"
        f"body_text_first_500:\n{soup.get_text(' ', strip=True)[:500]}\n"
    )
    (debug_dir / f"{safe_name}.txt").write_text(summary, encoding="utf-8")
    print(f"    [debug] tables={n_tables}, html→debug/{safe_name}.html")
    return n_tables


def parse_polls(html, target):
    """표를 찾아 폴 행 추출. BeautifulSoup 사용, 매우 보수적으로 파싱."""
    from bs4 import BeautifulSoup

    soup = BeautifulSoup(html, "lxml")
    out = []
    cd, cp = target["cand_dem"], target["cand_pp"]

    for table in soup.find_all("table"):
        text = table.get_text(" ", strip=True)
        # 두 후보 이름이 모두 있어야 후보 비교 표일 가능성 ↑
        if cd not in text or cp not in text:
            continue

        rows = table.find_all("tr")
        for tr in rows:
            cells = [td.get_text(" ", strip=True) for td in tr.find_all(["td", "th"])]
            joined = " | ".join(cells)
            # 날짜 패턴 (YYYY-MM-DD 또는 YYYY.MM.DD 또는 MM월 DD일)
            m_date = re.search(r"(20\d{2})[.\-]\s*(\d{1,2})[.\-]\s*(\d{1,2})", joined)
            if not m_date:
                continue
            year, mo, dy = m_date.groups()
            date = f"{year}-{int(mo):02d}-{int(dy):02d}"

            # 표본수 (1,000명 같은 패턴)
            m_n = re.search(r"([0-9,]{3,5})\s*명", joined) or re.search(r"\bn\s*=\s*([0-9,]+)", joined)
            n = int(m_n.group(1).replace(",", "")) if m_n else None

            # 퍼센트 두 개 이상 추출
            pcts = [float(x) for x in re.findall(r"(\d{1,2}\.?\d?)\s*%", joined)]
            if len(pcts) < 2:
                continue

            # 첫 두 % 가 보통 양당 후보 — 단 휴리스틱이라 100% 신뢰 X
            # 실제 운영시엔 PR 검토 단계에서 사람이 검증
            p_dem, p_pp = pcts[0] / 100, pcts[1] / 100

            # 의뢰처/조사기관은 첫 셀들에서 추정
            client = cells[0] if cells else "?"
            pollster = cells[1] if len(cells) > 1 else "?"

            out.append({
                "city": target["city"],
                "date": date,
                "pollster": pollster[:30],
                "client": client[:30],
                "n": n if n else 1000,
                "candDem": cd, "pDem": round(p_dem, 3),
                "candPP": cp, "pPP": round(p_pp, 3),
                "_raw": joined[:200],  # 디버그용 (PR 본문에 노출)
            })

    return out


def emit_output(key, value):
    """GitHub Actions 출력 변수 설정"""
    out_file = os.environ.get("GITHUB_OUTPUT")
    if out_file:
        with open(out_file, "a", encoding="utf-8") as f:
            f.write(f"{key}={value}\n")
    print(f"  output: {key}={value}")


def main():
    print("=" * 60)
    print(f"  나무위키 여론조사 스크레이퍼 — {datetime.now(KST).isoformat(timespec='seconds')}")
    print("=" * 60)

    # 기존 polls.json 로드
    with open(POLLS_PATH, encoding="utf-8") as f:
        data = json.load(f)
    existing_keys = {(p["city"], p["date"], p["pollster"]) for p in data["polls"]}
    print(f"  기존: {len(data['polls'])} 건")

    new_polls = []
    parse_failures = []
    page_hash_changed = False

    # 페이지 hash 기록 로드
    hashes = {}
    if HASH_PATH.exists():
        try:
            hashes = json.loads(HASH_PATH.read_text(encoding="utf-8"))
        except Exception:
            hashes = {}

    for tgt in TARGETS:
        print(f"\n--- {tgt['city']} ---")
        try:
            html = fetch_html(tgt["url"])
            new_hash = hashlib.sha256(html.encode("utf-8")).hexdigest()[:16]
            old_hash = hashes.get(tgt["city"])
            if old_hash and old_hash != new_hash:
                print(f"  페이지 변경 감지 ({old_hash} → {new_hash})")
                page_hash_changed = True
            hashes[tgt["city"]] = new_hash

            # 디버그 — html 저장 + 테이블 개수 확인
            n_tables = save_debug(tgt["city"], html)

            polls = parse_polls(html, tgt)
            print(f"  파싱: {len(polls)} 건 발견 (전체 테이블 {n_tables}개 중)")
            for p in polls:
                key = (p["city"], p["date"], p["pollster"])
                if key not in existing_keys:
                    new_polls.append(p)
                    existing_keys.add(key)
        except Exception as e:
            print(f"  ❌ 실패: {e}")
            parse_failures.append({"city": tgt["city"], "error": str(e)})

    # hash 저장
    HASH_PATH.write_text(json.dumps(hashes, ensure_ascii=False, indent=2), encoding="utf-8")

    # 결과
    print(f"\n--- 결과 ---")
    print(f"  새 폴: {len(new_polls)} 건")
    print(f"  파싱 실패: {len(parse_failures)} 도시")
    print(f"  페이지 변경: {'예' if page_hash_changed else '아니오'}")

    if new_polls:
        # polls.json 갱신
        # _raw 필드는 저장 안 함 (디버그용)
        clean_new = [{k: v for k, v in p.items() if not k.startswith("_")} for p in new_polls]
        data["polls"].extend(clean_new)
        data["lastUpdated"] = datetime.now(KST).isoformat(timespec="seconds")
        with open(POLLS_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        emit_output("has_new", "true")
        emit_output("new_count", str(len(new_polls)))
    else:
        emit_output("has_new", "false")
        emit_output("new_count", "0")

    if page_hash_changed and not new_polls:
        emit_output("hash_changed", "true")
    else:
        emit_output("hash_changed", "false")

    if parse_failures:
        emit_output("parse_failed", "true")
    else:
        emit_output("parse_failed", "false")

    # PR 본문용 마크다운
    summary_path = os.environ.get("GITHUB_STEP_SUMMARY")
    if summary_path:
        with open(summary_path, "a", encoding="utf-8") as f:
            f.write(f"## 스크레이핑 결과\n\n")
            f.write(f"- 새 폴: **{len(new_polls)}건**\n")
            f.write(f"- 페이지 변경: {'예' if page_hash_changed else '아니오'}\n")
            f.write(f"- 파싱 실패 도시: {len(parse_failures)}\n\n")
            if new_polls:
                f.write("### 새로 추가된 폴\n\n")
                for p in new_polls:
                    f.write(f"- `{p['city']}` `{p['date']}` "
                            f"{p['client']} / {p['pollster']} (n={p['n']}) "
                            f"→ {p['candDem']} {p['pDem']*100:.1f}% vs "
                            f"{p['candPP']} {p['pPP']*100:.1f}%\n")
                    f.write(f"  - raw: `{p.get('_raw', '')[:120]}`\n")
            if parse_failures:
                f.write("\n### 파싱 실패\n\n")
                for fail in parse_failures:
                    f.write(f"- {fail['city']}: {fail['error']}\n")


if __name__ == "__main__":
    main()
