"""
2026년 6월 3일 지방선거 - 당선 예측 콘솔 분석 (서울/성남/부산)
================================================================
PDF '머신러닝을 위한 통계 1' 의 핵심 개념을 그대로 적용:
  - 모비율 추정(p44):     p̂ ± 1.96 * sqrt(p̂q̂/n)         (95% 신뢰구간)
  - 오차범위 내 접전(p45): 두 후보의 신뢰구간이 겹치면 박빙
  - 중심극한정리(p28):     여러 표본의 평균은 정규분포에 수렴
  - 가설 검정(p59):        p-value = H1을 채택할 근거

실행:  python analyze.py            → 3개 도시 모두 분석
       python analyze.py 서울시장   → 특정 도시만
"""

import csv
import math
import sys
from pathlib import Path

Z_95 = 1.96   # 95% 신뢰수준 (PDF p39)
Z_99 = 2.58   # 99% 신뢰수준 (PDF p29, p47)


def load_polls(path):
    rows = []
    with open(path, encoding="utf-8") as f:
        for r in csv.DictReader(f):
            r["n"] = int(r["n"])
            r["p_dem"] = float(r["p_dem"])
            r["p_pp"] = float(r["p_pp"])
            rows.append(r)
    return rows


def ci_proportion(p_hat, n, z=Z_95):
    se = math.sqrt(p_hat * (1 - p_hat) / n)
    return p_hat - z * se, p_hat + z * se, z * se


def diff_se(p_a, p_b, n):
    """다항분포 공분산 반영: Var(p_A - p_B) = (pA + pB - (pA-pB)^2) / n"""
    return math.sqrt((p_a + p_b - (p_a - p_b) ** 2) / n)


def Phi(z):
    return 0.5 * (1 + math.erf(z / math.sqrt(2)))


def fmt_pct(x):
    return f"{x*100:5.1f}%"


def fmt_band(lo, hi):
    return f"[{fmt_pct(lo)}, {fmt_pct(hi)}]"


def analyse_one(poll):
    n = poll["n"]
    pd, pp = poll["p_dem"], poll["p_pp"]
    ld, hd, _ = ci_proportion(pd, n)
    lp, hp, _ = ci_proportion(pp, n)
    se = diff_se(pd, pp, n)
    diff = pd - pp
    z = diff / se if se > 0 else 0
    # 단측 H1: 민주 후보(p_dem) > 국힘 후보(p_pp). Z 부호와 무관하게 1-Φ(z) 사용.
    # 음의 z 가 나오면 p-value > 0.5 가 되어 "H1 증거 없음" 으로 올바르게 판정됨.
    p_value = 1 - Phi(z)
    overlap = not (ld > hp or lp > hd)
    decisive = abs(diff) > Z_95 * se
    return dict(poll=poll, ld=ld, hd=hd, lp=lp, hp=hp,
                se=se, diff=diff, z=z, p_value=p_value,
                overlap=overlap, decisive=decisive)


def poll_of_polls(polls):
    """표본수 가중 평균 (중심극한정리)"""
    W = sum(p["n"] for p in polls)
    pd = sum(p["p_dem"] * p["n"] for p in polls) / W
    pp = sum(p["p_pp"] * p["n"] for p in polls) / W
    return pd, pp, W


def analyse_city(city, polls):
    polls = [p for p in polls if p["city"] == city]
    if not polls:
        print(f"  ⚠ {city} 데이터 없음")
        return
    cand_dem = polls[0]["cand_dem"]
    cand_pp = polls[0]["cand_pp"]

    print("=" * 80)
    print(f"  {city}  ({cand_dem} 민주 vs {cand_pp} 국힘)")
    print("=" * 80)

    # [1] 조사별 신뢰구간
    print("\n[1] 조사별 95% 신뢰구간")
    print("-" * 80)
    print(f"{'조사일':<11}{'의뢰처':<10}{'n':>5}  "
          f"{cand_dem+' 95% CI':<22}{cand_pp+' 95% CI':<22}{'판정':<10}")
    print("-" * 80)
    results = []
    for poll in polls:
        r = analyse_one(poll)
        results.append(r)
        verdict = ("민주 우세" if r["decisive"] and r["diff"] > 0 else
                   "국힘 우세" if r["decisive"] and r["diff"] < 0 else
                   "오차범위 내")
        print(f"{poll['date']:<11}{poll['client']:<10}{poll['n']:>5}  "
              f"{fmt_band(r['ld'], r['hd']):<22}"
              f"{fmt_band(r['lp'], r['hp']):<22}{verdict}")

    # [2] 차이 신뢰구간 + 가설검정
    print(f"\n[2] 두 후보 차이 ({cand_dem} - {cand_pp}) 의 95% 신뢰구간 / p-value")
    print(f"    H0: 두 후보 동률   vs   H1: {cand_dem} > {cand_pp}")
    print("-" * 80)
    print(f"{'조사일':<11}{'의뢰처':<10}{'차이':>7}  "
          f"{'95% CI of 차이':<22}{'Z':>7}  {'p-value':>10}")
    print("-" * 80)
    for r in results:
        d = r["diff"]
        lo_d, hi_d = d - Z_95 * r["se"], d + Z_95 * r["se"]
        print(f"{r['poll']['date']:<11}{r['poll']['client']:<10}"
              f"{d*100:+6.1f}%p  {fmt_band(lo_d, hi_d):<22}"
              f"{r['z']:+7.2f}  {r['p_value']:10.5f}")

    # [3] Poll-of-polls
    pd, pp, n_eff = poll_of_polls(polls)
    ld, hd, _ = ci_proportion(pd, n_eff)
    lp, hp, _ = ci_proportion(pp, n_eff)
    se = diff_se(pd, pp, n_eff)
    diff = pd - pp
    z = diff / se if se > 0 else 0
    # 단측 H1: 민주 후보(p_dem) > 국힘 후보(p_pp). Z 부호와 무관하게 1-Φ(z) 사용.
    # 음의 z 가 나오면 p-value > 0.5 가 되어 "H1 증거 없음" 으로 올바르게 판정됨.
    p_value = 1 - Phi(z)
    leader = cand_dem if diff > 0 else cand_pp
    print(f"\n[3] Poll-of-polls (n_eff = {n_eff:,}) — 중심극한정리")
    print("-" * 80)
    print(f"  {cand_dem} 가중평균 : {fmt_pct(pd)}   95% CI = {fmt_band(ld, hd)}")
    print(f"  {cand_pp} 가중평균 : {fmt_pct(pp)}   95% CI = {fmt_band(lp, hp)}")
    print(f"  차이                : {diff*100:+.2f}%p   SE = {se*100:.2f}%p")
    print(f"  Z-statistic         : {z:+.2f}")
    # 매우 작은 p-value 도 정확히 표시 (1e-6 미만이면 지수 표기)
    p_str = f"{p_value:.6f}" if p_value >= 1e-6 else f"{p_value:.3e}"
    print(f"  p-value             : {p_str}")

    if p_value < 0.01:
        confidence = "매우 높음 (99% 신뢰수준)"
    elif p_value < 0.05:
        confidence = "높음 (95% 신뢰수준)"
    elif p_value < 0.10:
        confidence = "중간 (90% 신뢰수준)"
    else:
        confidence = "낮음 (오차범위 내 접전)"

    print(f"\n[4] 최종 예측")
    print("-" * 80)
    print(f"  당선 예측  : {leader}")
    print(f"  통계적 확신: {confidence}")
    print(f"  예상 득표율: {cand_dem} ≈ {fmt_pct(pd)}, {cand_pp} ≈ {fmt_pct(pp)}")
    print()


def main():
    base = Path(__file__).parent
    polls = load_polls(base / "polls.csv")

    cities = sorted(set(p["city"] for p in polls))

    if len(sys.argv) > 1:
        target = sys.argv[1]
        if target not in cities:
            print(f"❌ '{target}' 데이터 없음. 사용 가능: {', '.join(cities)}")
            return
        analyse_city(target, polls)
    else:
        for c in cities:
            analyse_city(c, polls)


if __name__ == "__main__":
    main()
