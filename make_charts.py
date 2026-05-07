"""
서울시장 선거 예측 - 보고서용 그래프 생성
==========================================
analyze.py 와 동일한 통계 공식을 사용해 4종 차트 PNG 출력:
  1) trend.png         시계열 추세 + 95% 신뢰구간 밴드
  2) ci_compare.png    조사별 후보 신뢰구간 비교 (forest plot 스타일)
  3) clt_effect.png    Poll-of-polls 효과 (중심극한정리)
  4) hypothesis.png    가설검정 p-value 시각화 (PDF p59)
"""

import csv
import math
from pathlib import Path

import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
from matplotlib import rcParams

# 한글 폰트 (Windows: Malgun Gothic)
rcParams["font.family"] = "Malgun Gothic"
rcParams["axes.unicode_minus"] = False
rcParams["figure.dpi"] = 110

Z95 = 1.96
COLOR_JUNG = "#2E7DD7"   # 파랑 (민주당)
COLOR_OH = "#E03131"     # 빨강 (국민의힘)


# ---------- 유틸 ----------
def ci(p, n, z=Z95):
    se = math.sqrt(p * (1 - p) / n)
    return p - z * se, p + z * se, z * se


def diff_se(pa, pb, n):
    return math.sqrt((pa + pb - (pa - pb) ** 2) / n)


def Phi(z):
    return 0.5 * (1 + math.erf(z / math.sqrt(2)))


CITY = "서울시장"  # README 의 정적 PNG 차트는 서울시장 분석만 다룸

def load_polls(path):
    """polls.csv 에서 CITY 도시만 골라 옛 코드 호환 키(p_jung/p_oh)로 변환"""
    rows = []
    with open(path, encoding="utf-8") as f:
        for r in csv.DictReader(f):
            if r["city"] != CITY:
                continue
            rows.append({
                "date": r["date"],
                "client": r["client"],
                "pollster": r["pollster"],
                "n": int(r["n"]),
                "p_jung": float(r["p_dem"]),  # 민주당 후보 → 정원오
                "p_oh":   float(r["p_pp"]),   # 국민의힘 후보 → 오세훈
            })
    return rows


# ---------- 차트 1: 시계열 추세 + CI 밴드 ----------
def chart_trend(polls, outpath):
    # 메이저 4건만 사용 (펜앤마이크는 시점이 다른 의뢰처라 제외)
    majors = sorted(polls, key=lambda r: r["date"])  # 5건 모두 사용

    dates = [p["date"] for p in majors]
    pj = np.array([p["p_jung"] for p in majors])
    po = np.array([p["p_oh"] for p in majors])
    ej = np.array([Z95 * math.sqrt(p["p_jung"] * (1 - p["p_jung"]) / p["n"])
                   for p in majors])
    eo = np.array([Z95 * math.sqrt(p["p_oh"] * (1 - p["p_oh"]) / p["n"])
                   for p in majors])

    fig, ax = plt.subplots(figsize=(10, 5.5))

    # 신뢰구간 밴드
    ax.fill_between(dates, (pj - ej) * 100, (pj + ej) * 100,
                    color=COLOR_JUNG, alpha=0.18, label="정원오 95% CI")
    ax.fill_between(dates, (po - eo) * 100, (po + eo) * 100,
                    color=COLOR_OH, alpha=0.18, label="오세훈 95% CI")

    # 점·선
    ax.plot(dates, pj * 100, "o-", color=COLOR_JUNG, lw=2.2, ms=8,
            label="정원오 (민주당)")
    ax.plot(dates, po * 100, "s-", color=COLOR_OH, lw=2.2, ms=8,
            label="오세훈 (국민의힘)")

    # 각 점에 값 라벨
    for d, v in zip(dates, pj * 100):
        ax.annotate(f"{v:.0f}%", (d, v), textcoords="offset points",
                    xytext=(0, 10), ha="center", fontsize=9, color=COLOR_JUNG)
    for d, v in zip(dates, po * 100):
        ax.annotate(f"{v:.0f}%", (d, v), textcoords="offset points",
                    xytext=(0, -15), ha="center", fontsize=9, color=COLOR_OH)

    # 박빙 → 우세 구간 표시
    ax.axvspan(-0.5, 1.5, alpha=0.08, color="gray")
    ax.text(0.5, 50, "오차범위 내 접전\n(Case B 가능)", ha="center",
            fontsize=9, color="gray", style="italic")
    ax.axvspan(1.5, 3.5, alpha=0.08, color=COLOR_JUNG)
    ax.text(2.5, 50, "정원오 우세 확정\n(Case A)", ha="center",
            fontsize=9, color=COLOR_JUNG, style="italic")

    ax.set_xlabel("조사일")
    ax.set_ylabel("지지율 (%)")
    ax.set_title("서울시장 후보 지지율 추세 — 95% 신뢰구간 (메이저 언론사 4건)",
                 fontsize=13, pad=12)
    ax.set_ylim(25, 55)
    ax.grid(True, alpha=0.3)
    ax.legend(loc="upper left", framealpha=0.9)

    plt.tight_layout()
    plt.savefig(outpath, dpi=130)
    plt.close()
    print(f"saved: {outpath}")


# ---------- 차트 2: 조사별 CI 비교 (forest plot) ----------
def chart_ci_compare(polls, outpath):
    fig, ax = plt.subplots(figsize=(11, 6))

    polls_sorted = sorted(polls, key=lambda r: r["date"])
    y_positions = list(range(len(polls_sorted)))

    for i, poll in enumerate(polls_sorted):
        n, pj, po = poll["n"], poll["p_jung"], poll["p_oh"]
        lj, hj, _ = ci(pj, n)
        lo, ho, _ = ci(po, n)
        overlap = lj <= ho

        # 정원오 CI
        ax.errorbar(pj * 100, i + 0.18,
                    xerr=[[(pj - lj) * 100], [(hj - pj) * 100]],
                    fmt="o", color=COLOR_JUNG, ms=9, capsize=6, lw=2.2,
                    label="정원오" if i == 0 else None)
        # 오세훈 CI
        ax.errorbar(po * 100, i - 0.18,
                    xerr=[[(po - lo) * 100], [(ho - po) * 100]],
                    fmt="s", color=COLOR_OH, ms=9, capsize=6, lw=2.2,
                    label="오세훈" if i == 0 else None)

        # 겹침/분리 표기
        verdict = "오차범위 내" if overlap else "오차범위 밖"
        col = "gray" if overlap else "green"
        ax.text(57, i, verdict, va="center", fontsize=9, color=col,
                weight="bold")

    labels = [f"{p['date']}\n({p['client']})" for p in polls_sorted]
    ax.set_yticks(y_positions)
    ax.set_yticklabels(labels)
    ax.set_xlabel("지지율 (%)")
    ax.set_xlim(25, 60)
    ax.set_title("조사별 95% 신뢰구간 비교 (PDF p45 '오차범위 내 접전' 시각화)",
                 fontsize=13, pad=30)
    ax.grid(True, alpha=0.3, axis="x")
    # 범례를 차트 위쪽 바깥으로 — 데이터·라벨 영역 침범 안 함
    ax.legend(loc="lower center", bbox_to_anchor=(0.5, 1.02),
              ncol=2, framealpha=0.9)
    ax.invert_yaxis()  # 최신이 위로

    plt.tight_layout()
    plt.savefig(outpath, dpi=130)
    plt.close()
    print(f"saved: {outpath}")


# ---------- 차트 3: 중심극한정리 - Poll-of-polls 효과 ----------
def chart_clt_effect(polls, outpath):
    majors = polls  # 5건 모두 사용

    # 가중평균 (n 비례)
    W = sum(p["n"] for p in majors)
    pj_agg = sum(p["p_jung"] * p["n"] for p in majors) / W
    po_agg = sum(p["p_oh"] * p["n"] for p in majors) / W
    n_eff = W

    fig, axes = plt.subplots(1, 2, figsize=(13, 5.2), sharey=True)

    # 좌측: 단일 조사 (5/2 SBS)
    poll = next(p for p in majors if p["date"] == "2026-05-02")
    n_single = poll["n"]
    pj_s, po_s = poll["p_jung"], poll["p_oh"]

    x = np.linspace(0.20, 0.55, 800)

    def normpdf(x, m, s):
        return np.exp(-0.5 * ((x - m) / s) ** 2) / (s * math.sqrt(2 * math.pi))

    se_j_s = math.sqrt(pj_s * (1 - pj_s) / n_single)
    se_o_s = math.sqrt(po_s * (1 - po_s) / n_single)

    axes[0].fill_between(x * 100, normpdf(x, pj_s, se_j_s),
                         alpha=0.55, color=COLOR_JUNG, label=f"정원오 ~ N({pj_s*100:.1f}, SE={se_j_s*100:.2f}%p)")
    axes[0].fill_between(x * 100, normpdf(x, po_s, se_o_s),
                         alpha=0.55, color=COLOR_OH, label=f"오세훈 ~ N({po_s*100:.1f}, SE={se_o_s*100:.2f}%p)")
    axes[0].set_title(f"단일 조사 (n={n_single}, SBS 5/2)\nMoE = ±{Z95*se_j_s*100:.2f}%p",
                      fontsize=11)
    axes[0].set_xlabel("지지율 (%)")
    axes[0].set_ylabel("표본평균의 확률밀도")
    axes[0].legend(loc="upper right", fontsize=9)
    axes[0].grid(alpha=0.3)

    # 우측: 합산
    se_j_agg = math.sqrt(pj_agg * (1 - pj_agg) / n_eff)
    se_o_agg = math.sqrt(po_agg * (1 - po_agg) / n_eff)

    axes[1].fill_between(x * 100, normpdf(x, pj_agg, se_j_agg),
                         alpha=0.55, color=COLOR_JUNG, label=f"정원오 ~ N({pj_agg*100:.1f}, SE={se_j_agg*100:.2f}%p)")
    axes[1].fill_between(x * 100, normpdf(x, po_agg, se_o_agg),
                         alpha=0.55, color=COLOR_OH, label=f"오세훈 ~ N({po_agg*100:.1f}, SE={se_o_agg*100:.2f}%p)")
    axes[1].set_title(f"Poll-of-polls (n_eff={n_eff})\nMoE = ±{Z95*se_j_agg*100:.2f}%p",
                      fontsize=11)
    axes[1].set_xlabel("지지율 (%)")
    axes[1].legend(loc="upper right", fontsize=9)
    axes[1].grid(alpha=0.3)

    fig.suptitle(r"중심극한정리: $\sigma/\sqrt{n}$ 효과 — 표본을 합치면 분포 폭이 좁아진다",
                 fontsize=13, y=1.02)
    plt.tight_layout()
    plt.savefig(outpath, dpi=130, bbox_inches="tight")
    plt.close()
    print(f"saved: {outpath}")


# ---------- 차트 4: 가설검정 p-value ----------
def chart_hypothesis(polls, outpath):
    majors = polls  # 5건 모두 사용
    W = sum(p["n"] for p in majors)
    pj = sum(p["p_jung"] * p["n"] for p in majors) / W
    po = sum(p["p_oh"] * p["n"] for p in majors) / W
    n = W
    se = diff_se(pj, po, n)
    diff = pj - po
    z_stat = diff / se
    p_val = 1 - Phi(z_stat)

    fig, ax = plt.subplots(figsize=(11, 5.5))

    # x 범위는 관측 Z 가 보이도록 동적 조정 (Z 가 6 보다 크면 그만큼 늘림)
    x_max = max(6, z_stat + 1.5)
    z = np.linspace(-4, x_max, 1200)
    pdf = np.exp(-0.5 * z * z) / math.sqrt(2 * math.pi)

    # H0 분포
    ax.plot(z, pdf, color="black", lw=1.8,
            label="H0: p(정) = p(오)  →  Z ~ N(0, 1)")

    # 95% 신뢰구간
    ax.fill_between(z, pdf, where=(z >= -Z95) & (z <= Z95),
                    color="lightblue", alpha=0.35, label="95% 신뢰영역 (|Z|<1.96)")

    # p-value 영역 (오른쪽 꼬리, Z >= z_stat)
    ax.fill_between(z, pdf, where=z >= z_stat,
                    color="red", alpha=0.65,
                    label=f"p-value 영역 (Z ≥ {z_stat:.2f})")

    # 검정통계량 위치
    ax.axvline(z_stat, color="red", lw=2, ls="--")
    ax.annotate(f"관측 Z = {z_stat:.2f}\n(차이 = +{diff*100:.2f}%p)",
                xy=(z_stat, 0.05), xytext=(z_stat - 1.3, 0.18),
                fontsize=11, color="red",
                arrowprops=dict(arrowstyle="->", color="red"))

    # 임계값
    ax.axvline(Z95, color="blue", lw=1, ls=":")
    ax.text(Z95 + 0.05, 0.36, "Z=1.96\n(α=0.05)", fontsize=9, color="blue")
    ax.axvline(2.58, color="purple", lw=1, ls=":")
    ax.text(2.58 + 0.05, 0.32, "Z=2.58\n(α=0.01)", fontsize=9, color="purple")

    ax.set_title(f"가설검정 시각화 (PDF p59) — p-value = {p_val:.2e}",
                 fontsize=13, pad=12)
    ax.set_xlabel("표준화 검정통계량 Z")
    ax.set_ylabel("확률밀도")
    ax.set_xlim(-4, x_max)
    ax.set_ylim(0, 0.45)
    ax.grid(alpha=0.3)
    ax.legend(loc="upper left", fontsize=10)

    # 결론 박스
    box_text = (f"H1: p(정) > p(오)\n"
                f"표본: n_eff = {n}\n"
                f"SE(차이) = {se*100:.2f}%p\n"
                f"p-value = {p_val:.2e}\n"
                f"→ 0.01 보다 작아\n   99% 신뢰수준에서 H0 기각")
    ax.text(0.97, 0.97, box_text, transform=ax.transAxes,
            ha="right", va="top", fontsize=10,
            bbox=dict(boxstyle="round,pad=0.5", facecolor="lightyellow",
                      edgecolor="orange", lw=1.5))

    plt.tight_layout()
    plt.savefig(outpath, dpi=130)
    plt.close()
    print(f"saved: {outpath}")


# ============= 메인 =============
def main():
    base = Path(__file__).parent
    polls = load_polls(base / "polls.csv")

    out = base / "charts"
    out.mkdir(exist_ok=True)

    chart_trend(polls, out / "trend.png")
    chart_ci_compare(polls, out / "ci_compare.png")
    chart_clt_effect(polls, out / "clt_effect.png")
    chart_hypothesis(polls, out / "hypothesis.png")

    print(f"\n{len(list(out.glob('*.png')))} 개 차트가 {out} 에 생성되었습니다.")


if __name__ == "__main__":
    main()
