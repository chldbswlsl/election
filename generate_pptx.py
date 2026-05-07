"""
보고서 → 발표용 PPTX 자동 생성
================================
12 슬라이드, 차트 임베드, 발표 노트 포함

실행: python generate_pptx.py
출력: 노희래_통계분석_발표.pptx (바탕화면 복사 포함)
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Cm
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from pathlib import Path
from datetime import datetime

ROOT = Path(__file__).parent
CHARTS = ROOT / "charts"
OUT_LOCAL = ROOT / "노희래_통계분석_발표.pptx"
OUT_DESKTOP = Path("C:/Users/노희래/Desktop/노희래_통계분석_발표.pptx")

KO_FONT = "맑은 고딕"
BLUE = RGBColor(0x1f, 0x3a, 0x6b)
DEM_BLUE = RGBColor(0x2e, 0x7d, 0xd7)
PP_RED = RGBColor(0xe0, 0x31, 0x31)
GRAY = RGBColor(0x55, 0x55, 0x55)


def set_korean(run, size=18, bold=False, color=None, font=KO_FONT):
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    if color: run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    eaFont = rPr.find(qn("a:ea"))
    if eaFont is None:
        eaFont = rPr.makeelement(qn("a:ea"), {"typeface": font})
        rPr.append(eaFont)
    else:
        eaFont.set("typeface", font)


def add_text(slide, text, left_cm, top_cm, width_cm, height_cm,
             size=18, bold=False, color=None, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(Cm(left_cm), Cm(top_cm),
                                   Cm(width_cm), Cm(height_cm))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    set_korean(run, size=size, bold=bold, color=color)
    return box


def add_bullets(slide, bullets, left_cm, top_cm, width_cm, height_cm, size=16):
    box = slide.shapes.add_textbox(Cm(left_cm), Cm(top_cm),
                                   Cm(width_cm), Cm(height_cm))
    tf = box.text_frame
    tf.word_wrap = True
    for i, b in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(8)
        run = p.add_run()
        run.text = "• " + b
        set_korean(run, size=size)


def add_image(slide, image_path, left_cm, top_cm, width_cm):
    return slide.shapes.add_picture(str(image_path),
                                    Cm(left_cm), Cm(top_cm), width=Cm(width_cm))


def add_notes(slide, text):
    """슬라이드 발표 노트 — 발표자가 보는 메모"""
    notes = slide.notes_slide
    tf = notes.notes_text_frame
    tf.text = text


def title_slide(prs, title, subtitle, footer):
    s = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    # 배경 — 위쪽 파란 띠
    from pptx.shapes.autoshape import Shape
    rect = s.shapes.add_shape(1, Cm(0), Cm(0), prs.slide_width, Cm(2.5))
    rect.fill.solid()
    rect.fill.fore_color.rgb = BLUE
    rect.line.fill.background()
    # 메인 타이틀
    add_text(s, title, 1.5, 7, 30, 3, size=40, bold=True, color=BLUE)
    add_text(s, subtitle, 1.5, 11, 30, 1.5, size=20, color=GRAY)
    add_text(s, footer, 1.5, 17, 30, 1, size=14, color=GRAY)
    return s


def section_slide(prs, title, accent="① "):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_text(s, accent + title, 1.5, 0.6, 30, 1.6,
             size=28, bold=True, color=BLUE)
    # 구분선
    from pptx.util import Emu
    line = s.shapes.add_connector(1, Cm(1.5), Cm(2.4), Cm(31.8), Cm(2.4))
    line.line.color.rgb = BLUE
    line.line.width = Pt(2)
    return s


# =================== 슬라이드 생성 ===================

prs = Presentation()
prs.slide_width = Cm(33.867)   # 16:9 widescreen
prs.slide_height = Cm(19.05)

# === 1. 표지 ===
s1 = title_slide(prs,
    "2026 서울시장 선거 당선 예측",
    "여론조사 데이터의 통계적 추론 — 도제 과제 보고서",
    f"노희래 · {datetime.now().strftime('%Y년 %m월')}")
add_notes(s1,
    "안녕하세요. 노희래입니다. 오늘 제가 발표할 주제는 2026년 6월 3일 서울시장 선거 당선 예측입니다. "
    "PDF 강의자료에서 배운 추론 통계 개념을 실제 여론조사 데이터에 적용해, 본인 나름대로 결과를 예측해 봤습니다.")

# === 2. 결론 한 페이지 ===
s2 = section_slide(prs, "한 줄 결론", "")
add_text(s2, "정원오 당선 예측", 1.5, 5.5, 30, 2,
         size=44, bold=True, color=DEM_BLUE, align=PP_ALIGN.CENTER)
add_text(s2, "99% 신뢰수준에서 통계적으로 유의 (p-value = 3.73 × 10⁻¹²)",
         1.5, 9, 30, 1.5, size=22, color=GRAY, align=PP_ALIGN.CENTER)
add_text(s2, "예상 득표율   정원오 ≈ 42.8%   vs   오세훈 ≈ 33.6%",
         1.5, 11.5, 30, 1.5, size=20, color=BLUE, align=PP_ALIGN.CENTER)
add_notes(s2,
    "결론부터 말씀드리면, 분석 결과 정원오 후보의 당선이 99% 신뢰수준에서 통계적으로 유의하게 예측됩니다. "
    "p-value 가 3.73 곱하기 10의 마이너스 12승 으로 사실상 0에 가깝습니다. "
    "이게 어떻게 나온 결과인지 지금부터 설명드리겠습니다.")

# === 3. 과제 목표 ===
s3 = section_slide(prs, "과제 목표 — 두 가지 핵심 질문", "1. ")
add_bullets(s3, [
    "단일 여론조사로 '몇 % 지지'라고 단정할 수 있는가?",
    "  → 못한다. 표본 비율 p̂ 주변에 '신뢰구간' 만 존재 (PDF p44)",
    "",
    "여러 조사를 합치면 더 좁은 추정이 가능한가?",
    "  → 가능하다. 중심극한정리에 의해 σ → σ/√n (PDF p28)",
    "",
    "이 두 가지 원리를 코드(analyze.py)로 옮겨, 서울시장 5건 여론조사를 분석",
], 2, 4, 30, 12, size=18)
add_notes(s3,
    "이 과제에서 풀어야 할 핵심 질문은 두 가지였습니다. "
    "첫째, 여론조사 한 건만으로 모집단 전체의 진짜 지지율을 단정할 수 있는가. 정답은 못 합니다. "
    "둘째, 여러 조사를 합치면 더 정확해지는가. 정답은 가능합니다. 중심극한정리 덕분이죠. "
    "이 두 원리를 그대로 파이썬 코드로 구현해서 서울시장 여론조사 5건을 분석했습니다.")

# === 4. 통계 개념 — 신뢰구간 ===
s4 = section_slide(prs, "통계 개념 ① — 모비율 신뢰구간 (PDF p44)", "2. ")
add_text(s4, "p̂ ± 1.96 · √(p̂(1-p̂)/n)", 1.5, 4, 30, 2,
         size=32, bold=True, color=BLUE, align=PP_ALIGN.CENTER)
add_bullets(s4, [
    "왜 1.96 인가? 표준정규분포 양쪽 꼬리 합 5%를 자르는 임계값 (PDF p15)",
    "PDF p9: '신뢰는 욕먹지 않을 만큼' — 95% = 100번 중 95번 맞을 표기",
    "",
    "5/2 SBS 손계산:",
    "  SE = √(0.41 × 0.59 / 800) = 0.01740",
    "  MoE = 1.96 × 0.01740 = 3.41%p  ← 언론사 발표 ±3.5%p 와 일치 ✓",
    "  CI(정원오) = [37.6%, 44.4%]  ← analyze.py 출력과 동일 ✓",
], 2, 7, 30, 11, size=15)
add_notes(s4,
    "첫 번째 통계 개념은 모비율 신뢰구간입니다. PDF p44에 나온 공식 그대로입니다. "
    "1.96 이라는 숫자는 표준정규분포에서 양쪽 꼬리 5%를 자를 때의 임계값입니다. "
    "PDF p9에서 '신뢰란 욕먹지 않을 만큼'이라고 표현한 게 인상 깊었는데요, "
    "100번 조사하면 95번은 이 구간 안에 모집단의 진짜 비율이 있을 것이라는 의미입니다. "
    "5월 2일 SBS 조사로 손계산해 보니 코드 출력값과 정확히 일치했습니다.")

# === 5. 통계 개념 — 가설검정 ===
s5 = section_slide(prs, "통계 개념 ② — 가설검정과 p-value (PDF p59)", "3. ")
add_text(s5, "Z = (p̂_정 - p̂_오) / SE_Δ      p-value = 1 - Φ(Z)",
         1.5, 4, 30, 1.8, size=22, bold=True, color=BLUE, align=PP_ALIGN.CENTER)
add_bullets(s5, [
    "H₀: 두 후보 동률   vs   H₁: 정원오 > 오세훈 (단측 검정)",
    "두 후보 차이의 분산(다항분포): Var = (p_A + p_B - (p_A - p_B)²) / n",
    "",
    "p-value 의 의미 (PDF p60):",
    "  • 1종 오류가 일어날 확률 — 차이 없는데 있다고 잘못 발표할 위험",
    "  • 작을수록 H₀ 기각 근거 강함 (보통 0.05 미만이면 유의)",
    "",
    "본 분석 결과: p-value ≈ 3.73 × 10⁻¹² → 사실상 0",
], 2, 7, 30, 11, size=15)
add_notes(s5,
    "두 번째는 가설검정입니다. 두 후보 지지율이 통계적으로 정말 다른지 검증합니다. "
    "귀무가설은 두 후보가 동률이라는 것이고, 대립가설은 정원오가 우세하다는 것입니다. "
    "단측 검정으로 진행했고, 검정통계량 Z 를 계산해 정규분포 위에서 p-value 를 구합니다. "
    "본 분석에서 나온 p-value 가 사실상 0 이라는 건 '동률이라는 가정 하에 이런 차이가 우연히 나올 확률이 거의 0' 이라는 의미입니다.")

# === 6. 데이터 ===
s6 = section_slide(prs, "분석 데이터 — 서울시장 5건 여론조사", "4. ")
table_data = [
    ["조사일", "의뢰처", "조사기관", "n", "정원오", "오세훈"],
    ["2025-12-15", "MBC", "코리아리서치", "800", "34.0%", "36.0%"],
    ["2026-02-10", "MBC", "코리아리서치", "800", "40.0%", "36.0%"],
    ["2026-04-20", "펜앤마이크", "여론조사공정", "1000", "49.5%", "30.9%"],
    ["2026-04-28", "MBC", "코리아리서치", "800", "48.0%", "32.0%"],
    ["2026-05-02", "SBS", "입소스", "800", "41.0%", "34.0%"],
    ["통합", "—", "—", "4,200", "42.83%", "33.64%"],
]
table = s6.shapes.add_table(rows=len(table_data), cols=6,
    left=Cm(2), top=Cm(4), width=Cm(29.8), height=Cm(11)).table
for ci, w in enumerate([4, 4, 5, 4, 6, 6]):
    table.columns[ci].width = Cm(w)
for ri, row in enumerate(table_data):
    for ci, val in enumerate(row):
        cell = table.cell(ri, ci)
        cell.text = val
        for p in cell.text_frame.paragraphs:
            p.alignment = PP_ALIGN.CENTER
            for run in p.runs:
                set_korean(run, size=14, bold=(ri==0 or ri==len(table_data)-1))
                if ri == 0:
                    run.font.color.rgb = RGBColor(0xff, 0xff, 0xff)
        if ri == 0:
            cell.fill.solid(); cell.fill.fore_color.rgb = BLUE
        elif ri == len(table_data) - 1:
            cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(0xe7, 0xee, 0xf7)
add_text(s6, "출처: 코리아리서치, 입소스, 여론조사공정 등 (중앙선거여론조사심의위 등록)",
         2, 16, 30, 1, size=12, color=GRAY)
add_notes(s6,
    "사용한 데이터입니다. 작년 12월부터 5월까지 5건의 서울시장 여론조사를 모았습니다. "
    "MBC, SBS 같은 메이저 방송사 의뢰 4건과, 보수성향 매체인 펜앤마이크 의뢰 1건 입니다. "
    "통합 표본 수는 4,200명 입니다. 각 조사에서 정원오 후보가 시간이 갈수록 우세를 굳히는 추세를 볼 수 있습니다.")

# === 7. 차트 1: 시계열 추세 ===
s7 = section_slide(prs, "분석 결과 ① — 지지율 추세 + 95% 신뢰구간 밴드", "5. ")
add_image(s7, CHARTS / "trend.png", 2, 3.5, 29)
add_text(s7, "12월 시점 박빙(Case B) → 4월 말부터 정원오 우세 확정(Case A)로 전환",
         2, 16.8, 30, 1, size=14, color=GRAY)
add_notes(s7,
    "첫 번째 차트는 시계열 추세입니다. 파란 선이 정원오, 빨간 선이 오세훈 후보입니다. "
    "주변의 옅은 색 밴드가 95% 신뢰구간이고요. "
    "12월에는 두 후보의 신뢰구간이 겹쳐서 PDF p45에서 본 Case B 박빙 상황이었지만, "
    "4월 말부터는 두 밴드가 완전히 분리되어 Case A 우세 확정으로 전환됩니다.")

# === 8. 차트 2: CI 비교 ===
s8 = section_slide(prs, "분석 결과 ② — 조사별 신뢰구간 비교 (Forest Plot)", "6. ")
add_image(s8, CHARTS / "ci_compare.png", 2, 3.5, 29)
add_text(s8, "겹치면 박빙(Case B), 분리되면 우세 확정(Case A) — PDF p45",
         2, 16.8, 30, 1, size=14, color=GRAY)
add_notes(s8,
    "두 번째 차트입니다. 각 조사를 한 행씩 보여주는 Forest Plot 입니다. "
    "각 후보의 지지율을 점으로, 95% 신뢰구간을 가로 바로 표시했습니다. "
    "두 바가 겹치는지 분리되는지가 핵심인데요, 4월 28일 MBC 조사부터 명확히 분리됩니다. "
    "오른쪽 초록색으로 '오차범위 밖' 이라고 표시된 것이 통계적 우세를 의미합니다.")

# === 9. 차트 3: CLT ===
s9 = section_slide(prs, "분석 결과 ③ — 합치면 좁아진다 (중심극한정리)", "7. ")
add_image(s9, CHARTS / "clt_effect.png", 2, 3.5, 29)
add_text(s9, "단일 조사 MoE ±3.41%p → 5건 합산 ±1.50%p (2.29배 좁아짐 = √5.25)",
         2, 16.8, 30, 1, size=14, color=GRAY)
add_notes(s9,
    "세 번째 차트는 PDF p28의 중심극한정리를 시각적으로 보여줍니다. "
    "왼쪽은 단일 조사 800명 표본일 때, 오른쪽은 5건 합쳐 4,200명 효과 표본일 때입니다. "
    "표본을 5.25배 늘리면 분포 폭이 √5.25 = 2.29배 좁아져, 두 후보 분포의 겹침이 거의 사라집니다. "
    "이게 여러 조사를 합쳐 분석하는 게 의미 있는 이유입니다.")

# === 10. 차트 4: 가설검정 ===
s10 = section_slide(prs, "분석 결과 ④ — 가설검정 시각화", "8. ")
add_image(s10, CHARTS / "hypothesis.png", 2, 3.5, 29)
add_text(s10, "관측 Z=+6.85 → 임계값 1.96, 2.58 모두 한참 초과 → 99% 신뢰수준 H₀ 기각",
         2, 16.8, 30, 1, size=14, color=GRAY)
add_notes(s10,
    "마지막 차트입니다. 표준정규분포 위에서 관측 Z 값의 위치를 보여줍니다. "
    "Z 가 +6.85로, 95% 신뢰수준 임계값 1.96과 99% 신뢰수준 임계값 2.58 모두 한참 넘어섰습니다. "
    "빨간 꼬리 면적이 p-value 인데, 사실상 0에 가깝습니다. "
    "결론적으로 H₀, 즉 '두 후보 동률' 가설을 99% 신뢰수준에서도 강하게 기각할 수 있습니다.")

# === 11. 한계 ===
s11 = section_slide(prs, "한계와 주의사항 — 솔직한 자기 평가", "9. ")
add_bullets(s11, [
    "양자대결 가정 — 후보 등록 후 다자 구도 가능 (조국혁신당, 무소속 등)",
    "비응답 편향 — 응답률 10~12%, 무응답자가 응답자와 다를 가능성 (PDF p4)",
    "시간 가중치 미적용 — 12월 조사와 5월 조사를 동등 가중",
    "House Effect 미보정 — 의뢰처별 편향 무시 (펜앤마이크 등)",
    "",
    "★ 가장 큰 변수: 부동층 결집 (PDF p60 2종 오류)",
    "  부동층 약 22%가 막판에 어느 쪽으로 결집하느냐가 결과를 흔들 수 있음",
    "  비관 시나리오 (부동층 7:3 → 오세훈) 시 격차 9.19%p → 0.4%p 로 박빙 가능",
], 2, 4, 30, 14, size=15)
add_notes(s11,
    "한계도 솔직하게 말씀드리겠습니다. 양자대결 가정과 비응답 편향, 시간 가중치 미적용 같은 이슈가 있습니다. "
    "특히 가장 큰 변수는 부동층 결집입니다. 약 22% 의 부동층이 막판에 어느 쪽으로 흘러가느냐에 따라 결과가 흔들릴 수 있습니다. "
    "비관 시나리오에서는 9%p 격차가 0.4%p 박빙까지 좁혀질 수 있습니다. "
    "p-value 가 작다고 결과가 '확정'된 건 아니라는 점, 그리고 모집단의 마음이 6월 3일까지 안 변한다는 보장이 없다는 점을 인정합니다.")

# === 12. 검증 + 마무리 ===
s12 = section_slide(prs, "6/3 이후 자체 검증 + 정리", "10. ")
add_text(s12, "검증 체크리스트", 2, 4, 28, 1, size=20, bold=True, color=BLUE)
add_bullets(s12, [
    "정원오 실제 득표율이 95% CI [41.3%, 44.3%] 안에 들어왔는가?",
    "오세훈 실제 득표율이 95% CI [32.2%, 35.1%] 안에 들어왔는가?",
    "1위 적중인가? — 가장 큰 검증",
    "빗나갔다면: 부동층 결집? 비응답 편향? 의뢰처 편향?",
], 2.5, 5.5, 28, 7, size=16)
add_text(s12, "라이브 사이트 + 코드 저장소", 2, 12.5, 28, 1, size=18, bold=True, color=BLUE)
add_bullets(s12, [
    "사이트: https://chldbswlsl.github.io/election/",
    "코드:    https://github.com/chldbswlsl/election",
], 2.5, 13.5, 28, 4, size=14)
add_notes(s12,
    "마지막입니다. 6월 3일 개표 결과가 나오면 다음 항목으로 본 모형을 검증할 계획입니다. "
    "정원오와 오세훈 두 후보의 실제 득표율이 각각 95% 신뢰구간 안에 들어왔는지, 그리고 1위가 적중인지가 핵심입니다. "
    "만약 빗나간다면 한계 슬라이드의 어느 항목이 결정적이었는지가 후속 학습 주제입니다. "
    "분석에 사용한 코드와 인터랙티브 사이트도 깃허브에 공개해 두었습니다. 감사합니다. 질문 있으시면 답변드리겠습니다.")

# === 저장 ===
prs.save(OUT_LOCAL)
import shutil
shutil.copy(OUT_LOCAL, OUT_DESKTOP)
size = OUT_LOCAL.stat().st_size
print(f"  ✓ {OUT_LOCAL.name} 생성 완료 ({size:,} bytes)")
print(f"  ✓ 데스크톱에도 복사: {OUT_DESKTOP}")
print(f"  ✓ 슬라이드 {len(prs.slides)}장 / 차트 4개 / 발표 노트 12개 포함")
