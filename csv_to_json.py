"""
polls.csv → polls.json + dashboard.html(EMBEDDED) 동시 변환기

새 여론조사를 polls.csv 에 한 줄 추가하고 이 스크립트 실행하면:
  1) polls.json 갱신 (대시보드가 fetch 하는 LIVE 데이터)
  2) dashboard.html 의 EMBEDDED 블록 갱신 (file:// 직접 열기 fallback)

→ 두 파일이 자동으로 동기화 되어 이중 관리 불필요.

실행: python csv_to_json.py
"""
import csv
import json
import re
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).parent
CSV_PATH = ROOT / "polls.csv"
JSON_PATH = ROOT / "polls.json"
HTML_PATH = ROOT / "dashboard.html"

EMBEDDED_START = "// EMBEDDED-START"
EMBEDDED_END = "// EMBEDDED-END"


def load_polls():
    polls = []
    with open(CSV_PATH, encoding="utf-8") as f:
        for row in csv.DictReader(f):
            p_dem = float(row["p_dem"])
            p_pp = float(row["p_pp"])
            p_others = float(row.get("p_others") or 0.0)
            # 부동층 = 1 - 양당 - 기타 (응답 거부·미정 포함)
            p_und = max(0.0, round(1.0 - p_dem - p_pp - p_others, 4))
            polls.append({
                "city": row["city"],
                "date": row["date"],
                "pollster": row["pollster"],
                "client": row["client"],
                "n": int(row["n"]),
                "candDem": row["cand_dem"],
                "pDem": p_dem,
                "candPP": row["cand_pp"],
                "pPP": p_pp,
                "pOthers": p_others,
                "pUnd": p_und,
                "method": row.get("method") or "?",
                "responseRate": float(row["response_rate"]) if row.get("response_rate") else None,
                "sourceUrl": row.get("source_url") or "",
            })
    return polls


def write_json(polls, last_updated):
    payload = {
        "lastUpdated": last_updated,
        "electionDate": "2026-06-03",
        "source": "한국갤럽·코리아리서치·입소스·한길리서치·KSOI·바로미터연구소·여론조사공정",
        "polls": polls,
    }
    with open(JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    return payload


def js_str(s):
    """JS 문자열 리터럴 — 백슬래시·따옴표 이스케이프"""
    return '"' + s.replace("\\", "\\\\").replace('"', '\\"') + '"'


def js_num(x):
    """소수점 trailing zero 제거 — 0.450 → 0.45, 0.0 → 0"""
    if x == int(x):
        return str(int(x))
    return repr(x).rstrip("0").rstrip(".")


def render_poll_line(p):
    """polls[i] → JS 객체 리터럴 한 줄 (기존 EMBEDDED 형식과 호환)"""
    rr = "null" if p["responseRate"] is None else js_num(p["responseRate"])
    return (
        f'    {{city:{js_str(p["city"])}, '
        f'date:{js_str(p["date"])}, '
        f'pollster:{js_str(p["pollster"])}, '
        f'client:{js_str(p["client"])}, '
        f'n:{p["n"]}, '
        f'candDem:{js_str(p["candDem"])}, pDem:{js_num(p["pDem"])}, '
        f'candPP:{js_str(p["candPP"])}, pPP:{js_num(p["pPP"])}, '
        f'pOthers:{js_num(p["pOthers"])}, pUnd:{js_num(p["pUnd"])}, '
        f'method:{js_str(p["method"])}, '
        f'responseRate:{rr}, '
        f'sourceUrl:{js_str(p["sourceUrl"])}}},'
    )


def update_embedded_in_html(polls, last_updated):
    if not HTML_PATH.exists():
        print(f"  ⚠ {HTML_PATH.name} 없음 — EMBEDDED 갱신 스킵")
        return False

    src = HTML_PATH.read_text(encoding="utf-8")
    start_idx = src.find(EMBEDDED_START)
    end_idx = src.find(EMBEDDED_END)
    if start_idx == -1 or end_idx == -1:
        print(f"  ⚠ {HTML_PATH.name} 에 {EMBEDDED_START}/{EMBEDDED_END} 마커 없음 — 스킵")
        return False
    if end_idx < start_idx:
        print(f"  ⚠ 마커 순서 오류 (END < START) — 스킵")
        return False

    lines = [render_poll_line(p) for p in polls]
    new_block = (
        f"{EMBEDDED_START} — csv_to_json.py 가 자동 갱신 (수동 편집 금지)\n"
        f"const EMBEDDED = {{\n"
        f'  lastUpdated: "{last_updated}",\n'
        f"  polls: [\n"
        + "\n".join(lines) + "\n"
        f"  ]\n"
        f"}};\n"
        f"{EMBEDDED_END}"
    )

    new_src = src[:start_idx] + new_block + src[end_idx + len(EMBEDDED_END):]
    if new_src == src:
        print(f"  · {HTML_PATH.name} EMBEDDED 변경 없음")
        return False
    HTML_PATH.write_text(new_src, encoding="utf-8")
    print(f"  ✓ {HTML_PATH.name} EMBEDDED 자동 갱신됨 ({len(polls)}건)")
    return True


def main():
    polls = load_polls()
    last_updated = datetime.now().astimezone().isoformat(timespec="seconds")

    write_json(polls, last_updated)
    update_embedded_in_html(polls, last_updated)

    cities = sorted(set(p["city"] for p in polls))
    print(f"  ✓ {len(polls)} 건의 여론조사를 {len(cities)} 개 도시로 정리")
    for c in cities:
        n = sum(1 for p in polls if p["city"] == c)
        print(f"     - {c}: {n} 건")
    print(f"  ✓ {JSON_PATH.name} 갱신됨 (lastUpdated = {last_updated})")


if __name__ == "__main__":
    main()
