# -*- coding: utf-8 -*-
"""27년 해외법인 예산 편성 양식 생성기 (배포용 EN + 검토용 KO).

사용법:  python build_budget.py HUK [출력폴더]
  - 법인 설정은 아래 ENTITIES에 추가한다. 사양/결정사항은 상위 폴더 CLAUDE.md 참고.
  - 원본 데이터: Supabase 테이블 (BP_TABLE, ACT_TABLE) — 실행할 때마다 최신 값을 읽는다.
  - 26년 법인 배포 파일: ../작년 최종 예산/
"""
import os, sys, glob, urllib.parse
import requests
import pandas as pd, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.utils import get_column_letter as CL

TOOL = os.path.dirname(os.path.abspath(__file__))          # .../v0. 요청자료/_작업도구
ROOT = os.path.dirname(TOOL)                                # .../v0. 요청자료
REF_DIR = os.path.join(ROOT, "작년 최종 예산")

# Supabase 원본 — 실적 누계 테이블이 바뀌면(예: 9월 누계) ACT_TABLE과 ACT_M을 함께 바꾼다
SUPABASE_URL = "https://acpfapffygogtdufssko.supabase.co"
BP_TABLE = "26년 예산(BP)"
ACT_TABLE = "26년 실적_8월 누계_260922"
# service role 키: 환경변수 SUPABASE_SERVICE_ROLE_KEY 우선, 없으면 아래 env 파일의 SERVICE_ROLE_KEY
# (두 테이블에 RLS 읽기 정책이 없어 anon 키로는 읽을 수 없음)
ENV_FILE = r"C:\Users\hyju\Desktop\클로드_고정비\Supabase 업로드용\env"
ACT_M = 8                            # actual months (Jan-Aug); 26 forecast = actual / ACT_M * 12
HR_EXTRA = {531030, 531210}          # scholarship / saving schemes : HR per 26 Notes (사용자 확정)
LABOR_MAJORS = {"11", "12", "13"}

# 법인별 설정 — 새 법인은 26년 배포 파일 구조를 확인한 뒤 같은 형식으로 추가
ENTITIES = {
    "HUK": dict(
        cur="GBP",
        ref_glob="3. HUK_*2026Y_Final.xlsx",
        divs=["London", "CS", "H-Direct"],                          # CC별 시트 (HUK만 예외적으로 CC 구분)
        cc_div={213210: "London", 213120: "CS", 260001: "H-Direct"},
        no_cc_div="London",                                         # CC 없는 본사 IT 서비스 비용
        sh_major="0.HUK(GBP)", major_rows=(4, 43),                  # 대계정 영문명
        sh_chart="HUK", chart_rows=(4, 117),                        # 계정 목록 (대계정 소계행 구분)
        div_sheets={"London": "London", "CS": "CS", "H-Direct": "H-Direct"},   # CC별 26BP (U열=26년 합계)
        det_src={"29": dict(sheet="Charges", rows=(9, 35), div=1, item=2, vendor=3, gl=4, total=23),
                 "47": dict(sheet="Promotion&Advertisement", rows=(9, 31), div=1, item=2, vendor=None, gl=3, total=22)},
        det_gl_override={"47": {"H-Direct": 535680}},               # 26BP는 H-Direct 광고비를 535680으로 계상
        langs=["en", "ko"],                                         # HUK만 영문(배포용) + 한글(검토용)
    ),
}
ENT = sys.argv[1] if len(sys.argv) > 1 else "HUK"
LANGS = ENTITIES[ENT].get("langs", ["ko"]) if ENT in ENTITIES else ["ko"]   # 기본은 한글판만
OUT_DIR = sys.argv[2] if len(sys.argv) > 2 else ROOT
E = ENTITIES[ENT]
CUR, DIVS, CC_DIV, NO_CC_DIV = E["cur"], E["divs"], E["cc_div"], E["no_cc_div"]
_refs = [f for f in glob.glob(os.path.join(REF_DIR, E["ref_glob"])) if "~$" not in f]
assert len(_refs) == 1, f"26년 배포 파일을 하나로 특정할 수 없음: {_refs}"
REF = _refs[0]

def _key():
    k = os.environ.get("SUPABASE_SERVICE_ROLE_KEY")
    if not k and os.path.exists(ENV_FILE):
        for line in open(ENV_FILE, encoding="utf-8"):
            if line.strip().startswith("SERVICE_ROLE_KEY="):
                k = line.split("=", 1)[1].strip().strip('"')
    assert k, "Supabase service role 키를 찾을 수 없음 (환경변수 SUPABASE_SERVICE_ROLE_KEY 또는 ENV_FILE 확인)"
    return k

def sb_select(table, cols, **eq):
    """Read rows from a Supabase table via PostgREST (paged, 1,000 rows per request)."""
    key = _key()
    url = f"{SUPABASE_URL}/rest/v1/{urllib.parse.quote(table)}"
    params = {"select": ",".join(cols), **{k: f"eq.{v}" for k, v in eq.items()}}
    rows, step = [], 1000
    while True:
        h = {"apikey": key, "Authorization": f"Bearer {key}", "Range-Unit": "items", "Range": f"{len(rows)}-{len(rows) + step - 1}"}
        r = requests.get(url, params=params, headers=h, timeout=60); r.raise_for_status()
        page = r.json(); rows += page
        if len(page) < step: break
    return pd.DataFrame(rows, columns=cols)

# ======================================================================= data (Supabase)
_BP_COLS = {"hq_corp": "본사/법인", "report_use": "보고용", "report_use_re": "보고용(re)", "gl_account": "G/L Account",
            "sub_account": "소계정", "main_account": "대계정", "category": "구분", "doc_currency": "통화(Doc)", "amount_doc_curr": "Amount(Doc)"}
_ACT_COLS = ["hq_corp", "report_use", "report_use_re", "gl_account", "sub_account", "main_account", "category",
             "local_currency", "amount_local_curr", "amount_krw", "month", "cost_center"]
bp_all = pd.concat([sb_select(BP_TABLE, ["gl_account", "category"])]).rename(columns=_BP_COLS)        # 계정별 구분(인건비) 판정용 전체
bp = sb_select(BP_TABLE, list(_BP_COLS), report_use_re=ENT).rename(columns=_BP_COLS)
for c in ("Amount(Doc)",): bp[c] = pd.to_numeric(bp[c])
bp = bp[(bp["본사/법인"] == "법인") & (~bp["보고용"].astype(str).str.contains("주재원"))]
assert set(bp["통화(Doc)"]) == {CUR}
bp_gl = bp.groupby(pd.to_numeric(bp["G/L Account"]))["Amount(Doc)"].sum()

act_cat = sb_select(ACT_TABLE, ["gl_account", "category"])                                         # 구분 판정용 전체
act = sb_select(ACT_TABLE, _ACT_COLS, report_use_re=ENT)
for c in ("amount_local_curr", "amount_krw"): act[c] = pd.to_numeric(act[c])
act = act[(act.hq_corp == "법인") & (~act.report_use.astype(str).str.contains("주재원"))].copy()
act["gl"] = pd.to_numeric(act.gl_account)
g = act[act.local_currency == CUR]
rate = g.groupby("month").amount_krw.sum() / g.groupby("month").amount_local_curr.sum()     # KRW per GBP, monthly
FX_AVG = round(float(rate.mean()), 2)                                            # Aug YTD avg actual rate
act["amt"] = act.apply(lambda x: x.amount_local_curr if x.local_currency == CUR else x.amount_krw / rate[x.month], axis=1)
act["div"] = pd.to_numeric(act.cost_center, errors="coerce").map(CC_DIV).fillna(NO_CC_DIV)
act_div = {d: act[act["div"] == d].groupby("gl").amt.sum() for d in DIVS}

cat = pd.concat([pd.DataFrame({"gl": pd.to_numeric(act_cat.gl_account), "c": act_cat.category}),
                 pd.DataFrame({"gl": pd.to_numeric(bp_all["G/L Account"], errors="coerce"), "c": bp_all["구분"]})]).dropna()
mode = cat.groupby("gl").c.agg(lambda s: s.value_counts().idxmax())
labor_gl = set(mode[mode == "인건비"].index.astype(int)) | HR_EXTRA

ref = openpyxl.load_workbook(REF, data_only=True)
bp_div = {}
for d in DIVS:
    bp_div[d] = pd.Series({int(r[0]): (r[20] or 0) for r in ref[E["div_sheets"][d]].iter_rows(min_row=4, values_only=True)
                           if isinstance(r[0], (int, float)) and r[0] > 500000})
chk = sum(bp_div[d].reindex(bp_gl.index).fillna(0) for d in DIVS) - bp_gl
assert chk.abs().max() < 0.01, "26BP by division != 26BP file"

major_en = {}
for r in ref[E["sh_major"]].iter_rows(min_row=E["major_rows"][0], max_row=E["major_rows"][1], max_col=2, values_only=True):
    if r[0]: major_en[str(r[0]).split()[0]] = (r[0], r[1])
chart, cur = [], []
for r in ref[E["sh_chart"]].iter_rows(min_row=E["chart_rows"][0], max_row=E["chart_rows"][1], max_col=3, values_only=True):
    if isinstance(r[0], (int, float)): cur.append((int(r[0]), r[1], r[2]))
    elif r[0]: chart.append((str(r[0]).split()[0], cur)); cur = []
names = {}
for _, s in pd.concat([act[["gl", "sub_account", "main_account"]].rename(columns={"sub_account": "kr", "main_account": "mj"}),
                       bp[["G/L Account", "소계정", "대계정"]].rename(columns={"G/L Account": "gl", "소계정": "kr", "대계정": "mj"})]).iterrows():
    names[int(s.gl)] = (s.kr, str(s.mj))
known = {a[0] for _, l in chart for a in l}
missing = sorted((set(bp_gl.index) | set(act.gl.dropna().astype(int))) - known)
for gl in missing:
    kr, mj = names[gl]; code = mj.split()[0]
    for c, l in chart:
        if c == code: l.append((gl, kr, kr)); break
    else: chart.append((code, [(gl, kr, kr)])); major_en.setdefault(code, (mj, mj))
chart.sort(key=lambda x: int(x[0]))
for c, l in chart: l.sort()
DETAIL = {"29": "Charges", "47": "Adv"}             # majors budgeted item-by-item on their own sheet
det_gls = {code: [a for c, l in chart if c == code for a in l] for code in DETAIL}
typ_of = lambda gl, code: "HR" if (gl in labor_gl or code in LABOR_MAJORS) else DETAIL.get(code, "Input")

det_items = {code: [] for code in DETAIL}
for code, sp in E["det_src"].items():           # 26년 상세 항목 (참고용 사전 기입)
    ov = E.get("det_gl_override", {}).get(code, {})
    for r in ref[sp["sheet"]].iter_rows(min_row=sp["rows"][0], max_row=sp["rows"][1], values_only=True):
        if r[sp["div"]] and r[sp["item"]] and r[sp["gl"]]:
            det_items[code].append((r[sp["div"]], " ".join(str(r[sp["item"]]).split()),
                                    r[sp["vendor"]] if sp["vendor"] is not None else None,
                                    ov.get(r[sp["div"]], int(r[sp["gl"]])), r[sp["total"]] or 0))
for code in DETAIL:
    assert abs(sum(i[4] for i in det_items[code]) - sum(bp_gl.get(gl, 0) for gl, _, _ in det_gls[code])) < 0.01

ITEM_KR = {"IT Service Fee": "IT 서비스 비용(본사)", "Sundry": "기타 잡비", "EDI Annual Charge": "EDI 연간 사용료",
           "Waste Management": "폐기물 처리", "Annual Support": "연간 기술지원", "UK registered office": "영국 등기 사무소",
           "Photocopier": "복합기 임차", "Mail collection": "우편 수거", "Payroll costs": "급여 대행 수수료",
           "Subscription": "구독료", "Membership": "멤버십", "Transfer Pricing": "이전가격 자문",
           "EV charging fee": "EV 충전 요금", "Test House": "테스트하우스 비용", "Bank Fees": "은행 수수료",
           "Audit and Tax Fees": "회계감사·세무 수수료", "Employment law support": "노무 자문",
           "EV legal support": "EV 법률 자문", "office and test house cleaning": "사무실·테스트하우스 청소",
           "Call Centre": "콜센터", "Admin": "관리 대행", "IT": "IT 보안 서비스", "Hosting/Mintenance": "호스팅/유지보수",
           "Accounting outsourcing(Jyothsna)": "회계 아웃소싱(Jyothsna)",
           "Online advertising ( Google, Meta, LinkedIn etc)": "온라인 광고 (Google, Meta, LinkedIn 등)",
           "Offline advertising ( Magazines, publications etc)": "오프라인 광고 (잡지, 간행물 등)",
           "Email marketing": "이메일 마케팅",
           "Printing, Display & Merchandise ( Brochures, pens, etc)": "인쇄물·디스플레이·판촉물 (브로셔, 펜 등)",
           "Exhibitions, Awards, Events & Sponsorships": "전시회·시상식·행사·후원",
           "Content creation ( photoshoot, video, design, images, blogs )": "콘텐츠 제작 (촬영, 영상, 디자인, 이미지, 블로그)",
           "PR activity": "PR 활동",
           "Other expenses ( hosting, platform fees, temp staff etc)": "기타 (호스팅, 플랫폼 수수료, 임시 인력 등)",
           "Website": "웹사이트"}

# ======================================================================= labels
L = {
 "en": dict(
  sh_guide="Guide", sh_sum="Summary", sh_tot="HUK Total",
  det={"29": dict(sh="Charges", title="▣ 29 Charges", total="Charges Total", ref="Summary — 29 Charges total (HUK Total)", check="Match with Charges total"),
       "47": dict(sh="Advertisement", title="▣ 47 Advertisement", total="Advertisement Total", ref="Summary — 47 Advertisement total (HUK Total)", check="Match with Advertisement total")},
  title="▣ {e} FY27 Fixed Cost Budget Planning", sub="Currency: {c}   |   Subsidiary costs only (expatriate costs excluded)",
  guide=[("How to fill in", None),
   ("1. London / CS / H-Direct", "Enter the FY27 monthly budget of each cost center in the BLUE cells of its own sheet. Explain the basis in 'Basis / Comment'."),
   ("2. Charges / Advertisement", "Enter every Charges (29) and Advertisement (47) item line by line: Division, item, vendor, G/L description (drop-down), monthly amount. Amounts flow to the cost center sheets automatically (GREEN cells)."),
   ("3. HUK Total", "Read-only. Sum of London + CS + H-Direct."),
   ("4. Summary", "Read-only. Choose the division in the yellow cell and check FY27 BP against the FY26 Forecast."),
   ("", None), ("Colour legend", None),
   ("BLUE", "Input cell (sheets that need input also have a light-blue tab)"), ("GREEN", "Do not enter — auto-filled from what you enter in the Charges / Advertisement sheets"),
   ("GREY", "Do not enter — labour cost (salary, bonus, pension, NI, medical, scholarship etc.). HR team will collect it with a separate form."),
   ("", None), ("Reference columns", None),
   ("FY26 BP", "FY26 Business Plan by cost center (expatriates excluded)"),
   ("FY26 Actual (Jan-Aug)", "Actual Jan–Aug FY26 by cost center. Non-GBP postings (EUR/KRW) converted at monthly actual rate. HQ IT service fee (no cost center) is included in London."),
   ("FY26 Forecast", "FY26 Actual (Jan-Aug) ÷ 8 × 12"),
   ("Ratio (%) / Diff", "Ratio (%) = FY27 BP ÷ FY26 Forecast,  Diff = FY27 BP − FY26 Forecast")],
  legend=["BLUE", "GREEN", "GREY"],
  warn="※ Do not change the format, formulas or row structure. Add Charges / Advertisement lines only inside their tables.",
  b_title="▣ {e} {d}", b_sub="[{c}]  BLUE = input.  GREY = labour cost (HR form, do not enter).  GREEN = from Charges / Advertisement sheets.",
  t_title="▣ {e} Total (London + CS + H-Direct)", t_sub="[{c}]  Sum of the three cost center sheets.",
  b_heads=["Major", "G/L No.", "G/L Description", "계정명", "Type", "FY26 BP", "FY26 Actual\n(Jan-Aug)", "FY26\nForecast"],
  mon=["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
  b_tail=["FY27\nTotal", "Ratio (%)", "Diff", "Basis / Comment"],
  fc_note="※ FY26 Forecast = FY26 Aug YTD actual ÷ 8 × 12   |   Ratio (%) = FY27 BP ÷ FY26 Forecast   |   Diff = FY27 BP − FY26 Forecast",
  ref26="Reference (FY26)", bud27="FY27 Budget", subtotal="Subtotal", total="TOTAL",
  types={"HR": "HR", "Charges": "Charges", "Adv": "Advertisement", "Input": "Input"},
  hr_memo="Labour cost — HR team collects this with a separate form.\nDo NOT enter here.",
  ch_ok="✔ Charges sheet matched", ch_ng="✖ Check Charges sheet (Division / G/L / total)",
  c_sub="[{c}]  One item per line. Choose Division and G/L description from the drop-down lists (G/L No. fills in automatically). Amounts flow to each cost center sheet.",
  c_heads=["No", "Division", "Item (expected event)", "Vendor", "G/L No.\n(auto)", "G/L Description", "FY26 BP\n(ref.)"],
  c_tail=["FY27\nTotal", "Basis / Comment"],
  c_divsub="{d} subtotal", c_ok="✔ Matched", c_ng="✖ Mismatch",
  ck=dict(div="Division missing", gl="G/L missing", bad="Invalid G/L", ok="OK", mis="Mismatch"),
  s_title="▣ {e} FY27 Budget Summary", s_sel="Division ▶", s_note="[{c}]  Read-only. Choose the division in the yellow cell.",
  s_heads=["Account", "Description", "FY26 BP", "FY26 Actual\n(Jan-Aug)", "FY26\nForecast", "FY27 BP", "Ratio (%)", "Diff"],
  s_fx="Business plan FX rate (KRW / GBP) ▶", s_krw="[KRW]  = GBP × business plan FX rate (default: Jan–Aug FY26 average actual rate)",
  s_total="TOTAL", s_chk="Check vs sheet", s_chkline="Charges / Advertisement check: ", s_okw="OK", s_ngw="MISMATCH",
  hr_note="※ Grey (HR) accounts are blank until HQ consolidates the HR form — ratios of 11/12/13/14 are not meaningful until then.",
 ),
 "ko": dict(
  sh_guide="안내", sh_sum="요약", sh_tot="HUK 합계",
  det={"29": dict(sh="지급수수료", title="▣ 29 Charges", total="지급수수료 합계", ref="요약 시트 29 지급수수료 합계 (HUK 합계 기준)", check="지급수수료 합계와 일치 여부"),
       "47": dict(sh="광고선전비", title="▣ 47 Advertisement", total="광고선전비 합계", ref="요약 시트 47 광고선전비 합계 (HUK 합계 기준)", check="광고선전비 합계와 일치 여부")},
  title="▣ {e} 27년 고정비 예산 편성", sub="통화: {c}   |   법인 비용만 포함 (주재원 비용 제외)",
  guide=[("작성 방법", None),
   ("1. London / CS / H-Direct", "각 CC(코스트센터) 시트의 파란 칸에 27년 월별 예산을 입력. '산정근거 / 코멘트' 열에 근거 기재."),
   ("2. 지급수수료 / 광고선전비", "지급수수료(29)·광고선전비(47) 항목을 한 줄씩 입력: Division, 항목, 거래처, 계정명(드롭다운), 월별 금액. 금액은 각 CC 시트(초록 칸)로 자동 반영."),
   ("3. HUK 합계", "읽기 전용. London + CS + H-Direct 합계."),
   ("4. 요약", "읽기 전용. 노란 칸에서 Division 선택. 26년 예상 대비 27년 BP 비율 확인."),
   ("", None), ("색상 안내", None),
   ("파랑", "입력 칸 (입력이 필요한 시트는 시트 탭도 하늘색)"), ("초록", "입력 금지 — 지급수수료·광고선전비 시트에서 작성 시 자동 반영"),
   ("회색", "입력 금지 — 인건비(급여, 상여, 연금, NI, 의료보험, 자녀장학금 등). HR에서 별도 양식으로 취합."),
   ("", None), ("참고 열 설명", None),
   ("26년 BP", "26년 사업계획, CC별 (주재원 제외)"),
   ("26년 실적(1~8월)", "26년 1~8월 실적, CC별. GBP 외(EUR/KRW) 전표는 월별 실적환율로 환산. CC 없는 본사 IT 서비스 비용은 London에 포함."),
   ("26년 예상", "26년 실적(1~8월) ÷ 8 × 12"),
   ("비율(%) / Diff", "비율(%) = 27년 BP ÷ 26년 예상,  Diff = 27년 BP − 26년 예상")],
  legend=["파랑", "초록", "회색"],
  warn="※ 양식, 수식, 행 구조를 변경하지 마세요. 지급수수료·광고선전비 항목은 각 시트의 표 안에서만 추가하세요.",
  b_title="▣ {e} {d}", b_sub="[{c}]  파랑 = 입력.  회색 = 인건비(HR 별도, 입력 금지).  초록 = 지급수수료·광고선전비 시트 연동.",
  t_title="▣ {e} 합계 (London + CS + H-Direct)", t_sub="[{c}]  3개 CC 시트 합계.",
  b_heads=["대계정", "계정코드", "영문 계정명", "계정명", "구분", "26년 BP", "26년 실적\n(1~8월)", "26년\n예상"],
  mon=["1월", "2월", "3월", "4월", "5월", "6월", "7월", "8월", "9월", "10월", "11월", "12월"],
  b_tail=["27년\n합계", "비율(%)", "Diff", "산정근거 / 코멘트"],
  fc_note="※ 26년 예상 = 26년 8월 누계 실적 ÷ 8 × 12   |   비율(%) = 27년 BP ÷ 26년 예상   |   Diff = 27년 BP − 26년 예상",
  ref26="참고 (26년)", bud27="27년 예산", subtotal="소계", total="합계",
  types={"HR": "인건비", "Charges": "지급수수료", "Adv": "광고선전비", "Input": "입력"},
  hr_memo="인건비 — HR에서 별도 양식으로 취합.\n여기에 입력하지 마세요.",
  ch_ok="✔ 지급수수료 시트 일치", ch_ng="✖ 지급수수료 시트 확인 (Division / 계정 / 합계)",
  c_sub="[{c}]  한 줄에 한 항목. Division과 계정명은 드롭다운에서 선택(계정코드는 자동 표시). 금액은 각 CC 시트로 자동 반영.",
  c_heads=["No", "Division", "항목 (예상 내역)", "거래처", "계정코드\n(자동)", "계정명", "26년 BP\n(참고)"],
  c_tail=["27년\n합계", "산정근거 / 코멘트"],
  c_divsub="{d} 소계", c_ok="✔ 일치", c_ng="✖ 불일치",
  ck=dict(div="Division 누락", gl="계정 누락", bad="계정 오류", ok="OK", mis="불일치"),
  s_title="▣ {e} 27년 예산 요약", s_sel="Division ▶", s_note="[{c}]  읽기 전용. 노란 칸에서 Division 선택.",
  s_heads=["대계정", "영문명", "26년 BP", "26년 실적\n(1~8월)", "26년\n예상", "27년 BP", "비율(%)", "Diff"],
  s_fx="사업계획 환율 (KRW / GBP) ▶", s_krw="[KRW]  = GBP × 사업계획 환율 (기본값: 26년 1~8월 실적 평균환율)",
  s_total="합계", s_chk="시트 검증", s_chkline="지급수수료·광고선전비 체크: ", s_okw="OK", s_ngw="불일치",
  hr_note="※ 회색(인건비) 계정은 HR 양식 취합 전까지 공란이라, 11/12/13/14 대계정의 비율은 그 전까지 의미가 없습니다.",
 ),
}

# ======================================================================= styles
NAVY, BLUE_IN, GREY, GREEN, SUBT, TOT, YEL = "1F3864", "DDEBF7", "D9D9D9", "E2EFDA", "F2F2F2", "FFF2CC", "FFFF99"
thin = Side(style="thin", color="BFBFBF"); BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
fill = lambda c: PatternFill("solid", fgColor=c)
HFONT = Font(bold=True, color="FFFFFF", name="Arial", size=9)
BASE = Font(name="Arial", size=9); BOLD = Font(name="Arial", size=9, bold=True)
TITLE = Font(name="Arial", size=12, bold=True, color=NAVY)
NUM, PCT = '#,##0;[Red](#,##0);"-"', '0.0%;[Red]-0.0%;"-"'
RED_FILL, RED_FONT = PatternFill("solid", fgColor="F8CBAD", bgColor="F8CBAD"), Font(color="9C0006", bold=True)
UNLOCK = Protection(locked=False)
q = lambda s: "'" + s + "'"

def hdr(ws, row, c1, c2, color=NAVY):
    for c in range(c1, c2 + 1):
        x = ws.cell(row, c); x.font = HFONT; x.fill = fill(color); x.border = BOX
        x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

def style_row(ws, row, c1, c2, font=BASE, bg=None):
    for c in range(c1, c2 + 1):
        x = ws.cell(row, c); x.font = font; x.border = BOX
        if bg: x.fill = fill(bg)

def build(lang, out):
    T = L[lang]
    wb = openpyxl.Workbook()
    # ------------------------------------------------ Guide
    gd = wb.active; gd.title = T["sh_guide"]; gd.sheet_view.showGridLines = False
    gd.column_dimensions["A"].width = 3; gd.column_dimensions["B"].width = 28; gd.column_dimensions["C"].width = 100
    gd["B2"] = T["title"].format(e=ENT); gd["B2"].font = Font(name="Arial", size=14, bold=True, color=NAVY)
    gd["B3"] = T["sub"].format(c=CUR); gd["B3"].font = Font(name="Arial", size=9, italic=True)
    leg = dict(zip(T["legend"], [BLUE_IN, GREEN, GREY]))
    for i, (k, v) in enumerate(T["guide"]):
        r = 5 + i
        gd.cell(r, 2, k).font = BOLD if v is None else BASE
        if v: gd.cell(r, 3, v).font = BASE; gd.cell(r, 3).alignment = Alignment(wrap_text=True, vertical="top")
        if k in leg: gd.cell(r, 2).fill = fill(leg[k])
    gd.cell(5 + len(T["guide"]) + 1, 2, T["warn"]).font = Font(name="Arial", size=9, color="C00000", bold=True)

    # ------------------------------------------------ List (hidden)
    ls = wb.create_sheet("List")
    ls.cell(1, 5, "Division")
    LCOL = {"29": 1, "47": 9}                         # 29 -> A:C, 47 -> I:K
    for code, c0 in LCOL.items():
        ls.cell(1, c0, f"{code} G/L No."); ls.cell(1, c0 + 1, "Description"); ls.cell(1, c0 + 2, "계정명")
        for i, (gl, en, kr) in enumerate(det_gls[code]):
            ls.cell(i + 2, c0, gl); ls.cell(i + 2, c0 + 1, en); ls.cell(i + 2, c0 + 2, kr)
    for i, d in enumerate([T["sh_tot"]] + DIVS): ls.cell(i + 2, 5, d)
    ls.sheet_state = "hidden"

    # ------------------------------------------------ detail sheets: 29 Charges, 47 Advertisement (built first)
    ck = T["ck"]
    H, CTOT, M1, CBAS = 4, 8, 9, 21
    C_NUM = [7] + list(range(CTOT, M1 + 12))          # numeric cols: 26 ref, 27 total, Jan..Dec
    geo = {}

    def detail_sheet(code):
        D = T["det"][code]
        ch = wb.create_sheet(D["sh"])
        ch.sheet_view.showGridLines = False
        ch["A1"] = D["title"]; ch["A1"].font = TITLE
        ch["A2"] = T["c_sub"].format(c=CUR); ch["A2"].font = BASE
        heads = T["c_heads"] + T["c_tail"][:1] + T["mon"] + T["c_tail"][1:]
        for i, h in enumerate(heads, 1): ch.cell(H, i, h)
        hdr(ch, H, 1, len(heads))
        first, last = H + 1, H + 30
        c0 = LCOL[code]; LA, LN = CL(c0), CL(c0 + (1 if lang == "en" else 2))
        gl_name = {gl: (en if lang == "en" else kr) for gl, en, kr in det_gls[code]}
        items = det_items[code]
        for i in range(last - first + 1):
            r = first + i
            ch.cell(r, 1, i + 1)
            if i < len(items):
                d, it, v, gl, ref26 = items[i]
                ch.cell(r, 2, d); ch.cell(r, 3, ITEM_KR.get(it, it) if lang == "ko" else it); ch.cell(r, 4, v); ch.cell(r, 6, gl_name[gl]); ch.cell(r, 7, round(ref26, 2))
            ch.cell(r, 5, f'=IF(F{r}="","",IFERROR(INDEX(List!${LA}:${LA},MATCH(F{r},List!${LN}:${LN},0)),"{ck["bad"]}"))')
            ch.cell(r, CTOT, f"=SUM({CL(M1)}{r}:{CL(M1 + 11)}{r})")
            style_row(ch, r, 1, len(heads))
            for c in [2, 3, 4, 6] + list(range(M1, M1 + 12)) + [CBAS]:
                ch.cell(r, c).fill = fill(BLUE_IN); ch.cell(r, c).protection = UNLOCK
            for c in C_NUM: ch.cell(r, c).number_format = NUM
        rT = last + 1
        ch.cell(rT, 3, D["total"])
        for c in C_NUM: ch.cell(rT, c, f"=SUM({CL(c)}{first}:{CL(c)}{last})").number_format = NUM
        style_row(ch, rT, 1, len(heads), BOLD, TOT)
        for j, d in enumerate(DIVS):
            r = rT + 1 + j
            ch.cell(r, 2, d); ch.cell(r, 3, T["c_divsub"].format(d=d))
            for c in C_NUM:
                ch.cell(r, c, f"=SUMIFS({CL(c)}${first}:{CL(c)}${last},$B${first}:$B${last},$B{r})").number_format = NUM
            style_row(ch, r, 1, len(heads), BASE, SUBT)
        rC = rT + len(DIVS) + 1                          # check: this sheet's total vs HUK Total major subtotal
        ch.cell(rC, 3, D["check"]); style_row(ch, rC, 1, len(heads), BOLD)
        dv = DataValidation(type="list", formula1=f"List!${LN}$2:${LN}${len(det_gls[code]) + 1}", allow_blank=True); ch.add_data_validation(dv); dv.add(f"F{first}:F{last}")
        dvd = DataValidation(type="list", formula1=f"List!$E$3:$E${len(DIVS) + 2}", allow_blank=True); ch.add_data_validation(dvd); dvd.add(f"B{first}:B{last}")
        for i, w in enumerate([5, 11, 40, 20, 10, 32, 11, 11] + [9.5] * 12 + [40], 1): ch.column_dimensions[CL(i)].width = w
        ch.row_dimensions[H].height = 30; ch.freeze_panes = ch.cell(H + 1, 7)
        geo[code] = dict(ws=ch, CH=q(D["sh"]), first=first, last=last, rT=rT, rC=rC)

    for code in DETAIL: detail_sheet(code)

    # ------------------------------------------------ Budget sheets (Total + each CC), identical row layout
    BH = 4
    bheads = T["b_heads"] + T["b_tail"][:1] + T["mon"] + T["b_tail"][1:]
    B_TOT, B_M1, B_R1, B_R2, B_BAS = 9, 10, 22, 23, 24
    B_LAST = B_M1 + 11
    layout = []
    r = BH + 1
    for code, accts in chart:
        start = r
        for gl, en, kr in accts:
            layout.append(("acct", r, code, gl, en, names[gl][0] if gl in names else kr, typ_of(gl, code))); r += 1
        layout.append(("sub", r, code, start)); r += 1
    D1, DN, BT = BH + 1, r - 1, r
    sub_row = {x[2]: x[1] for x in layout if x[0] == "sub"}

    def budget_sheet(name, div):
        ws = wb.create_sheet(name)
        ws.sheet_view.showGridLines = False
        ws["A1"] = (T["b_title"].format(e=ENT, d=div) if div else T["t_title"].format(e=ENT)); ws["A1"].font = TITLE
        ws["A2"] = (T["b_sub"] if div else T["t_sub"]).format(c=CUR); ws["A2"].font = BASE
        for i, h in enumerate(bheads, 1): ws.cell(BH, i, h)
        hdr(ws, BH, 1, len(bheads))
        for c in range(B_M1, B_M1 + 12): ws.cell(BH, c).fill = fill("2F5597")
        prev_hr = False
        for item in layout:
            if item[0] == "acct":
                _, r, code, gl, en, kr, typ = item
                ws.cell(r, 1, code); ws.cell(r, 2, gl); ws.cell(r, 3, en); ws.cell(r, 4, kr); ws.cell(r, 5, T["types"][typ])
                style_row(ws, r, 1, len(bheads))
                if div:
                    ws.cell(r, 6, round(float(bp_div[div].get(gl, 0)), 2)); ws.cell(r, 7, round(float(act_div[div].get(gl, 0)), 2))
                    for c in range(B_M1, B_M1 + 12):
                        x = ws.cell(r, c)
                        if typ in ("Charges", "Adv"):
                            g = geo[code]; CH, f1, f2 = g["CH"], g["first"], g["last"]; mc = CL(M1 + c - B_M1)
                            x.value = f'=SUMIFS({CH}!${mc}${f1}:${mc}${f2},{CH}!$B${f1}:$B${f2},"{div}",{CH}!$E${f1}:$E${f2},$B{r})'
                            x.fill = fill(GREEN)
                        elif typ == "HR": x.fill = fill(GREY)
                        else: x.fill = fill(BLUE_IN); x.protection = UNLOCK
                    ws.cell(r, B_BAS).fill = fill(BLUE_IN); ws.cell(r, B_BAS).protection = UNLOCK
                else:
                    for c in [6, 7] + list(range(B_M1, B_M1 + 12)):
                        ws.cell(r, c, "=" + "+".join(f"{q(d)}!{CL(c)}{r}" for d in DIVS))
                        if c >= B_M1 and typ != "Input": ws.cell(r, c).fill = fill(GREY if typ == "HR" else GREEN)
                prev_hr = typ == "HR"
                ws.cell(r, 8, f"=G{r}/{ACT_M}*12")
                ws.cell(r, B_TOT, f"=SUM({CL(B_M1)}{r}:{CL(B_M1 + 11)}{r})")
            else:
                _, r, code, start = item
                lbl = major_en.get(code, (code, code))
                ws.cell(r, 3, f"{lbl[1]}  {T['subtotal']}"); ws.cell(r, 4, lbl[0])
                for c in range(6, B_LAST + 1): ws.cell(r, c, f"=SUM({CL(c)}{start}:{CL(c)}{r - 1})")
                style_row(ws, r, 1, len(bheads), BOLD, SUBT)
                ws.row_dimensions.group(start, r - 1, hidden=False, outline_level=1)
                prev_hr = False
            ws.cell(r, B_R1, f'=IF(H{r}=0,"",{CL(B_TOT)}{r}/H{r})'); ws.cell(r, B_R2, f"={CL(B_TOT)}{r}-H{r}")
            for c in range(6, B_LAST + 1): ws.cell(r, c).number_format = NUM
            ws.cell(r, B_R1).number_format = PCT; ws.cell(r, B_R2).number_format = NUM
        ws.cell(BT, 3, T["total"])
        for c in range(6, B_LAST + 1):
            ws.cell(BT, c, f'=SUMIFS({CL(c)}{D1}:{CL(c)}{DN},$B{D1}:$B{DN},"<>")').number_format = NUM
        ws.cell(BT, B_R1, f'=IF(H{BT}=0,"",{CL(B_TOT)}{BT}/H{BT})').number_format = PCT
        ws.cell(BT, B_R2, f"={CL(B_TOT)}{BT}-H{BT}").number_format = NUM
        style_row(ws, BT, 1, len(bheads), BOLD, TOT)
        ws.cell(BT + 2, 1, T["fc_note"]).font = Font(name="Arial", size=8, italic=True, color="808080")
        for i, w in enumerate([6, 8, 34, 24, 9, 11, 11, 11, 11] + [9] * 12 + [9, 12, 45], 1): ws.column_dimensions[CL(i)].width = w
        ws.row_dimensions[BH].height = 30
        ws.column_dimensions.group(CL(B_M1), CL(B_LAST), hidden=False, outline_level=1)
        ws.column_dimensions.group("G", "G", hidden=True, outline_level=1)   # 26 actual: grouped, collapsed
        ws.freeze_panes = "F5"

    budget_sheet(T["sh_tot"], None)
    for d in DIVS: budget_sheet(d, d)
    matches = []
    for code, g in geo.items():                      # detail sheet total vs HUK Total major subtotal
        ref_cell = f"{q(T['sh_tot'])}!{CL(B_TOT)}{sub_row[code]}"
        m = f"ROUND({g['CH']}!{CL(CTOT)}{g['rT']}-{ref_cell},2)=0"; matches.append(m)
        g["ws"].conditional_formatting.add(f"{CL(CTOT)}{g['rC']}", FormulaRule(formula=[f"NOT({m})"], fill=RED_FILL, font=RED_FONT))
        x = g["ws"].cell(g["rC"], CTOT, f'=IF({m},"{T["c_ok"]}","{T["c_ng"]}")'); x.alignment = Alignment(horizontal="center")
    CH_MATCH = "AND(" + ",".join(matches) + ")"

    # ------------------------------------------------ Summary
    sm = wb.create_sheet(T["sh_sum"])
    sm.sheet_view.showGridLines = False
    sm["A1"] = T["s_title"].format(e=ENT); sm["A1"].font = TITLE
    sm["A2"] = T["s_note"].format(c=CUR); sm["A2"].font = BASE
    sm["B3"] = T["s_sel"]; sm["B3"].font = BOLD; sm["B3"].alignment = Alignment(horizontal="right")
    sm["C3"] = T["sh_tot"]; sm["C3"].fill = fill(YEL); sm["C3"].font = BOLD; sm["C3"].border = BOX; sm["C3"].protection = UNLOCK
    dvs = DataValidation(type="list", formula1=f"List!$E$2:$E${len(DIVS) + 2}", allow_blank=False); sm.add_data_validation(dvs); dvs.add("C3")
    ls["G1"] = "sel"; ls["G2"] = f"=MATCH({q(T['sh_sum'])}!$C$3,List!$E$2:$E${len(DIVS) + 2},0)"
    SEL = "List!$G$2"
    chk_all = CH_MATCH
    sm["E3"] = f'="{T["s_chkline"]}"&IF({chk_all},"{T["s_okw"]}","{T["s_ngw"]}")'; sm["E3"].font = BOLD
    sm.conditional_formatting.add("E3", FormulaRule(formula=[f"NOT({chk_all})"], fill=RED_FILL, font=RED_FONT))
    SH = 5; NC = len(T["s_heads"])                     # A..H
    for i, h in enumerate(T["s_heads"], 1): sm.cell(SH, i, h)
    hdr(sm, SH, 1, NC)
    src = {3: "F", 4: "G", 5: "H", 6: CL(B_TOT)}
    order = [T["sh_tot"]] + DIVS

    def pick(col, code):
        parts = [f'SUMIFS({q(n)}!${col}${D1}:${col}${DN},{q(n)}!$A${D1}:$A${DN},"{code}")' for n in order]
        return f"=CHOOSE({SEL},{','.join(parts)})"

    def finish(r, off=0, font=BASE, bg=None):
        n = lambda c: CL(c + off)
        sm.cell(r, 7 + off, f'=IF({n(5)}{r}=0,"",{n(6)}{r}/{n(5)}{r})').number_format = PCT
        sm.cell(r, 8 + off, f"={n(6)}{r}-{n(5)}{r}").number_format = NUM
        style_row(sm, r, (1 if off == 0 else 2) + off, NC + off, font, bg)

    r = SH + 1
    for code, accts in chart:                         # major accounts, ascending
        lbl = major_en.get(code, (code, code))
        sm.cell(r, 1, lbl[0]); sm.cell(r, 2, lbl[1])
        for c, col in src.items(): sm.cell(r, c, pick(col, code)).number_format = NUM
        finish(r)
        if all(typ_of(gl, code) == "HR" for gl, _, _ in accts): sm.cell(r, 1).fill = fill("F7F7F7")   # labour major
        r += 1
    ST = r
    sm.cell(r, 1, T["s_total"])
    for c in range(3, 7): sm.cell(r, c, f"=SUM({CL(c)}{SH + 1}:{CL(c)}{r - 1})").number_format = NUM
    finish(r, 0, BOLD, TOT)
    sm.cell(r + 1, 1, T["s_chk"])
    for c, col in src.items():
        tot_pick = f"CHOOSE({SEL},{','.join(f'{q(n)}!{col}{BT}' for n in order)})"
        sm.cell(r + 1, c, f'=IF(ROUND({CL(c)}{r}-{tot_pick},2)=0,"OK","{ck["mis"]}")').alignment = Alignment(horizontal="center")
    style_row(sm, r + 1, 1, NC, Font(name="Arial", size=8, italic=True, color="808080"))
    sm.cell(r + 3, 1, T["fc_note"]).font = Font(name="Arial", size=8, italic=True, color="808080")
    sm.cell(r + 4, 1, T["hr_note"]).font = Font(name="Arial", size=8, italic=True, color="808080")

    # KRW table (copy of B:H) to the right, driven by the business plan FX cell
    KO = NC + 1                                        # KRW table starts after a 2-column gap
    FXC = f"${CL(2 + KO + 3)}$3"
    sm.cell(3, 2 + KO, T["s_fx"]).font = BOLD
    x = sm[FXC.replace("$", "")]; x.value = FX_AVG; x.number_format = "#,##0.00"; x.fill = fill(YEL); x.font = BOLD; x.border = BOX; x.protection = UNLOCK
    sm.cell(4, 2 + KO, T["s_krw"]).font = Font(name="Arial", size=8, italic=True, color="808080")
    for i, h in enumerate(T["s_heads"][1:], 2): sm.cell(SH, i + KO, h)
    hdr(sm, SH, 2 + KO, NC + KO)
    for r in range(SH + 1, ST + 1):
        sm.cell(r, 2 + KO, f"=B{r}" if r < ST else T["s_total"])
        for c in range(3, 7): sm.cell(r, c + KO, f"={CL(c)}{r}*{FXC}").number_format = NUM
        finish(r, KO, BOLD if r == ST else BASE, TOT if r == ST else None)

    sm.cell(ST + 3, 2 + KO, T["fc_note"]).font = Font(name="Arial", size=8, italic=True, color="808080")
    widths = {1: 24, 2: 32, 3: 13, 4: 13, 5: 13, 6: 13, 7: 12, 8: 13, NC + 1: 3, NC + 2: 3,
              2 + KO: 32, 3 + KO: 16, 4 + KO: 16, 5 + KO: 16, 6 + KO: 16, 7 + KO: 12, 8 + KO: 16}
    for c, w in widths.items(): sm.column_dimensions[CL(c)].width = w
    sm.column_dimensions.group("D", "D", hidden=True, outline_level=1)                    # 26 actual
    sm.column_dimensions.group(CL(4 + KO), CL(4 + KO), hidden=True, outline_level=1)
    sm.row_dimensions[SH].height = 30; sm.freeze_panes = "C6"

    # ------------------------------------------------ order, protection, print
    det_names = [T["det"][code]["sh"] for code in DETAIL]
    order_names = [T["sh_guide"], T["sh_sum"], T["sh_tot"]] + DIVS + det_names + ["List"]
    wb._sheets = [wb[n] for n in order_names]
    for n in DIVS + det_names: wb[n].sheet_properties.tabColor = BLUE_IN      # input sheets
    wb[T["sh_guide"]].sheet_properties.tabColor = "000000"                         # guide = black tab
    hr_rows = [x[1] for x in layout if x[0] == "acct" and x[6] == "HR"]
    LOCK = Protection(locked=True)
    for n in DIVS:                                   # protect only the grey HR month cells (no password)
        ws = wb[n]
        for row in ws.iter_rows(min_row=1, max_row=BT + 20, max_col=30):
            for c in row: c.protection = UNLOCK
        for r in hr_rows:
            for c in range(B_M1, B_M1 + 12): ws.cell(r, c).protection = LOCK
        p = ws.protection
        p.sheet = True
        p.formatCells = p.formatColumns = p.formatRows = False
        p.insertRows = p.insertColumns = p.deleteRows = p.deleteColumns = False
        p.sort = p.autoFilter = False
    for n in order_names[1:-1]:
        ws = wb[n]
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        ws.page_setup.orientation = "landscape"; ws.page_setup.fitToHeight = 0
    wb.active = 0
    wb.save(out)
    return dict(D1=D1, DN=DN, BT=BT, sub29=sub_row["29"], sub47=sub_row["47"], **{f"{c}_{k}": v for c, g in geo.items() for k, v in g.items() if k != "ws"})

if __name__ == "__main__":
    if LANGS == ["en", "ko"]:   # HUK: 영문 배포용 + 한글 검토용
        info = build("en", os.path.join(OUT_DIR, f"{ENT}_2027Y Budget Planning.xlsx"))
        build("ko", os.path.join(OUT_DIR, f"{ENT}_2027Y Budget Planning_한글(검토용).xlsx"))
    else:                       # 기본: 한글판 1개
        info = build("ko", os.path.join(OUT_DIR, f"{ENT}_27년 예산편성.xlsx"))
    print(info)
    for d in DIVS: print(d, "BP", round(bp_div[d].sum(), 1), "ACT", round(act_div[d].sum(), 1))
    print("BP total", round(bp_gl.sum(), 1), "ACT total", round(act.amt.sum(), 1), "missing-from-chart", missing)
