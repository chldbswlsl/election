"""
polls.csv → polls.json 변환기
새 여론조사를 polls.csv 에 한 줄 추가하고 이 스크립트 실행하면 polls.json 갱신됨.
실행: python csv_to_json.py
"""
import csv
import json
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).parent
CSV_PATH = ROOT / "polls.csv"
JSON_PATH = ROOT / "polls.json"


def main():
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
            })

    payload = {
        "lastUpdated": datetime.now().astimezone().isoformat(timespec="seconds"),
        "electionDate": "2026-06-03",
        "source": "한국갤럽·코리아리서치·입소스·한길리서치·KSOI·바로미터연구소·여론조사공정",
        "polls": polls,
    }

    with open(JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    cities = sorted(set(p["city"] for p in polls))
    print(f"  ✓ {len(polls)} 건의 여론조사를 {len(cities)} 개 도시로 정리")
    for c in cities:
        n = sum(1 for p in polls if p["city"] == c)
        print(f"     - {c}: {n} 건")
    print(f"  ✓ {JSON_PATH.name} 갱신됨 (lastUpdated = {payload['lastUpdated']})")


if __name__ == "__main__":
    main()
