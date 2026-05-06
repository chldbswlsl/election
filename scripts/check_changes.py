"""
간단한 변경 감지기 — Wikipedia API 기반
==========================================
Cloudflare 차단된 나무위키 대신, 한국 위키백과 API 로
관련 페이지의 최신 revision 을 체크한다.

전략:
  1) 위키백과 API 호출 (rate limit 너그러움, Cloudflare 없음)
  2) 마지막으로 본 revision id 와 비교 (.wiki_state.json 에 저장)
  3) revision 이 바뀌면 워크플로우가 Issue 자동 생성

이걸로 새 여론조사가 "있는지"는 모르지만, 페이지가 편집되었다는 신호는 받음.
사용자가 알림 받고 페이지 가서 확인 → 새 폴이면 polls.csv 에 수동 추가.
"""

import json
import os
import sys
from datetime import datetime, timezone, timedelta
from pathlib import Path
from urllib.parse import quote
from urllib.request import Request, urlopen

KST = timezone(timedelta(hours=9))
ROOT = Path(__file__).resolve().parent.parent
STATE_PATH = ROOT / ".wiki_state.json"

# 모니터링할 위키백과 페이지들
PAGES = [
    "제9회_전국동시지방선거",
    "서울특별시장_선거",
    "부산광역시장_선거",
]

UA = "election-dashboard-bot/1.0 (https://github.com/chldbswlsl/election; bydgpt@gmail.com)"


def fetch_latest_revision(title):
    """Wikipedia API로 최신 revision 정보 조회"""
    api = (
        f"https://ko.wikipedia.org/w/api.php?"
        f"action=query&format=json&titles={quote(title)}"
        f"&prop=revisions&rvprop=ids|timestamp|user|comment|size&rvlimit=1"
    )
    req = Request(api, headers={"User-Agent": UA, "Accept": "application/json"})
    with urlopen(req, timeout=20) as resp:
        data = json.loads(resp.read().decode("utf-8"))

    pages = data.get("query", {}).get("pages", {})
    if not pages:
        return None
    page = next(iter(pages.values()))
    if page.get("missing") is not None:
        return {"missing": True}
    revs = page.get("revisions", [])
    if not revs:
        return None
    rev = revs[0]
    return {
        "revid": rev["revid"],
        "timestamp": rev["timestamp"],
        "user": rev.get("user", "?"),
        "comment": rev.get("comment", "")[:200],
        "size": rev.get("size", 0),
    }


def emit_output(key, value):
    out_file = os.environ.get("GITHUB_OUTPUT")
    if out_file:
        with open(out_file, "a", encoding="utf-8") as f:
            f.write(f"{key}={value}\n")
    print(f"  output: {key}={value}")


def main():
    print("=" * 60)
    print(f"  변경 감지기 — {datetime.now(KST).isoformat(timespec='seconds')}")
    print("=" * 60)

    # 이전 상태 로드
    state = {}
    if STATE_PATH.exists():
        try:
            state = json.loads(STATE_PATH.read_text(encoding="utf-8"))
        except Exception:
            state = {}

    changes = []
    failures = []

    for title in PAGES:
        print(f"\n--- {title} ---")
        try:
            cur = fetch_latest_revision(title)
            if cur is None or cur.get("missing"):
                print(f"  페이지 없음 또는 빈 응답")
                continue

            prev = state.get(title)
            if prev is None:
                print(f"  최초 기록 (revid={cur['revid']}, {cur['timestamp']})")
            elif prev["revid"] != cur["revid"]:
                print(f"  변경 감지: {prev['revid']} → {cur['revid']}")
                print(f"    by {cur['user']} at {cur['timestamp']}")
                print(f"    comment: {cur['comment']}")
                changes.append({
                    "title": title,
                    "old_revid": prev["revid"],
                    "new_revid": cur["revid"],
                    "user": cur["user"],
                    "timestamp": cur["timestamp"],
                    "comment": cur["comment"],
                    "size_delta": cur["size"] - prev.get("size", 0),
                })
            else:
                print(f"  변경 없음 (revid={cur['revid']})")

            state[title] = cur
        except Exception as e:
            print(f"  ❌ 실패: {e}")
            failures.append({"title": title, "error": str(e)})

    # 상태 저장
    STATE_PATH.write_text(
        json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    print(f"\n--- 결과 ---")
    print(f"  변경: {len(changes)} 페이지")
    print(f"  실패: {len(failures)} 페이지")

    emit_output("has_changes", "true" if changes else "false")
    emit_output("change_count", str(len(changes)))

    # PR 본문용 마크다운
    summary_path = os.environ.get("GITHUB_STEP_SUMMARY")
    body_lines = []
    if changes:
        body_lines.append("## 위키백과 변경 감지\n")
        for c in changes:
            url = f"https://ko.wikipedia.org/wiki/{quote(c['title'])}"
            diff_url = f"https://ko.wikipedia.org/w/index.php?title={quote(c['title'])}&diff={c['new_revid']}&oldid={c['old_revid']}"
            body_lines.append(
                f"### [{c['title'].replace('_', ' ')}]({url})\n"
                f"- **편집자**: {c['user']}\n"
                f"- **시각**: {c['timestamp']}\n"
                f"- **요약**: {c['comment'] or '(편집 요약 없음)'}\n"
                f"- **크기 변화**: {c['size_delta']:+d} bytes\n"
                f"- [diff 보기]({diff_url})\n"
            )
        body_lines.append("\n→ 페이지 가서 새 여론조사 정보가 추가됐는지 확인하세요.")
        body_lines.append("\n→ 새 폴이면 `polls.csv` 한 줄 추가 후 `python csv_to_json.py` → `git push`")

    body = "\n".join(body_lines)
    if summary_path and body:
        with open(summary_path, "a", encoding="utf-8") as f:
            f.write(body)

    # Issue 본문은 별도 파일로 출력 (workflow 가 읽음)
    if changes:
        issue_body_path = ROOT / ".issue_body.md"
        issue_body_path.write_text(body, encoding="utf-8")


if __name__ == "__main__":
    main()
