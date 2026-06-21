"""
조달청 특정품목 조달내역 대시보드
app.py

실행: streamlit run app.py
"""

import io
import base64
from datetime import datetime, timedelta, date
from pathlib import Path

import pandas as pd
import streamlit as st
import plotly.express as px

from g2b_api import fetch_all, API_KEY
from report import build_analysis, build_docx, _eok as _r_eok, _won as _r_won
from excel_report import build_excel_report


def _load_logo_b64():
    """로고 파일을 같은 폴더에서 찾아 base64로 반환. 없으면 None."""
    candidates = ["archipace_logo.png", "아키페이스_CI.png", "logo.png"]
    here = Path(__file__).parent
    for name in candidates:
        p = here / name
        if p.exists():
            return base64.b64encode(p.read_bytes()).decode("ascii")
    return None

# ──────────────────────────────────────────────
# 기본 설정
# ──────────────────────────────────────────────
st.set_page_config(
    page_title="ARCHIPACE | 조달 시장 분석",
    page_icon="📊",
    layout="wide",
)

# ──────────────────────────────────────────────
# 디자인 테마 (베이지·브라운 · 오렌지 포인트)
# ──────────────────────────────────────────────
st.markdown("""
<style>
/* 전체 폰트 */
html, body, [class*="css"]  {
    font-family: 'Pretendard', 'Noto Sans KR', -apple-system, BlinkMacSystemFont, sans-serif;
}

/* 메인 배경 */
.stApp {
    background-color: #FAF6F0;
}

/* 사이드바 */
section[data-testid="stSidebar"] {
    background-color: #F5EFE5;
    border-right: 1px solid #E8DFD3;
}
section[data-testid="stSidebar"] > div {
    padding-top: 1.2rem;
}

/* 사이드바 텍스트 */
section[data-testid="stSidebar"] .stMarkdown,
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] p {
    color: #3D2E1F !important;
}

/* 사이드바 제목 */
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 {
    color: #3D2E1F !important;
    font-weight: 600;
}

/* 입력칸 */
.stTextInput input, .stDateInput input, .stSelectbox > div > div {
    background-color: #FFFFFF !important;
    border: 1px solid #E8DFD3 !important;
    border-radius: 4px !important;
    color: #3D2E1F !important;
}
.stTextInput input:focus, .stDateInput input:focus {
    border-color: #DC6400 !important;
    box-shadow: 0 0 0 2px rgba(220, 100, 0, 0.1) !important;
}

/* 기본 버튼 */
.stButton button {
    background-color: #FFFFFF;
    color: #3D2E1F;
    border: 1px solid #D4C5B0;
    border-radius: 4px;
    font-weight: 500;
    transition: all 0.2s ease;
}
.stButton button:hover {
    border-color: #DC6400;
    color: #DC6400;
    background-color: #FFFFFF;
}

/* primary 버튼 (검색하기, 다운로드 등) */
.stButton button[kind="primary"],
.stDownloadButton button[kind="primary"] {
    background-color: #DC6400 !important;
    color: #FFFFFF !important;
    border: 1px solid #DC6400 !important;
    font-weight: 600;
    box-shadow: 0 2px 4px rgba(220, 100, 0, 0.15);
}
.stButton button[kind="primary"]:hover,
.stDownloadButton button[kind="primary"]:hover {
    background-color: #B85400 !important;
    border-color: #B85400 !important;
    color: #FFFFFF !important;
}

/* 메인 영역 제목 */
h1, h2, h3, h4 {
    color: #3D2E1F !important;
    font-weight: 600;
    letter-spacing: -0.3px;
}

/* 본문 텍스트 */
.main p, .main span, .main label, .main div {
    color: #3D2E1F;
}

/* 메트릭 카드 */
[data-testid="stMetric"] {
    background-color: #FFFFFF;
    padding: 16px 20px;
    border-radius: 6px;
    border: 1px solid #E8DFD3;
    border-left: 3px solid #DC6400;
}
[data-testid="stMetricLabel"] {
    color: #9C7C5C !important;
    font-size: 12px !important;
    font-weight: 500 !important;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}
[data-testid="stMetricValue"] {
    color: #3D2E1F !important;
    font-weight: 700;
    font-size: 26px !important;
}

/* 탭 디자인 */
.stTabs [data-baseweb="tab-list"] {
    gap: 0;
    border-bottom: 1px solid #E8DFD3;
}
.stTabs [data-baseweb="tab"] {
    background-color: transparent;
    color: #9C7C5C;
    font-weight: 500;
    padding: 10px 18px;
    border-radius: 0;
    border: none;
}
.stTabs [aria-selected="true"] {
    color: #DC6400 !important;
    border-bottom: 2px solid #DC6400 !important;
    font-weight: 600 !important;
}

/* 라디오 버튼 */
.stRadio label {
    color: #3D2E1F !important;
}

/* 표 */
.stDataFrame {
    border: 1px solid #E8DFD3;
    border-radius: 6px;
    overflow: hidden;
}

/* 정보 박스 */
.stAlert {
    background-color: #FFFFFF;
    border-left: 3px solid #DC6400;
    border-radius: 4px;
}

/* expander */
.streamlit-expanderHeader {
    background-color: #FFFFFF !important;
    border: 1px solid #E8DFD3 !important;
    color: #3D2E1F !important;
    font-weight: 500;
}

/* 구분선 */
hr {
    border-color: #E8DFD3 !important;
}

/* 진행률 막대 */
.stProgress > div > div > div > div {
    background-color: #DC6400;
}

/* 슬라이더 포인트 */
.stSlider [data-baseweb="slider"] [role="slider"] {
    background-color: #DC6400 !important;
}

/* Streamlit 기본 푸터/메뉴 숨기기 (선택) */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

/* 헤더 영역 컨테이너 */
.archi-header {
    background: #FFFFFF;
    padding: 18px 28px;
    border-bottom: 3px solid #DC6400;
    margin: -1rem -1rem 1.5rem -1rem;
    display: flex;
    align-items: center;
    gap: 24px;
}
.archi-header-divider {
    width: 1px;
    height: 36px;
    background: #D4C5B0;
}
.archi-header-title {
    font-size: 17px;
    color: #3D2E1F;
    font-weight: 600;
    letter-spacing: -0.3px;
    margin: 0;
    line-height: 1.3;
}
.archi-header-subtitle {
    font-size: 12px;
    color: #9C7C5C;
    margin: 2px 0 0 0;
}
.archi-header-meta {
    margin-left: auto;
    text-align: right;
    font-size: 11px;
    color: #9C7C5C;
    line-height: 1.6;
}

/* 푸터 */
.archi-footer {
    background: #3D2E1F;
    color: #D4C5B0;
    padding: 14px 28px;
    margin: 3rem -1rem -1rem -1rem;
    font-size: 11px;
    display: flex;
    justify-content: space-between;
    border-radius: 0;
}
.archi-footer .right {
    color: #9C7C5C;
}
</style>
""", unsafe_allow_html=True)

# 응답 필드 → 한글 컬럼명 매핑
FIELD_MAP = {
    "prcrmntDivNm": "조달구분",
    "cntrctDivNm": "계약구분",
    "cntrctDlvrDivNm": "계약납품구분",
    "cntrctDlvrReqDate": "계약일자",
    "cntrctDlvrReqNo": "계약번호",
    "cntrctDlvrReqChgOrd": "변경차수",
    "fnlCntrctDlvrReqChgOrdYn": "최종차수여부",
    "dminsttNm": "수요기관",
    "dmndInsttDivNm": "수요기관구분",
    "dminsttRgnNm": "수요기관지역",
    "dminsttCd": "수요기관코드",
    "prdctClsfcNo": "물품분류번호",
    "prdctClsfcNoNm": "품명",
    "dtilPrdctClsfcNo": "세부품명번호",
    "dtilPrdctClsfcNoNm": "세부품명",
    "prdctIdntNo": "물품식별번호",
    "prdctIdntNoNm": "물품규격명",
    "prdctUprc": "단가",
    "prdctQty": "수량",
    "prdctUnit": "단위",
    "prdctAmt": "공급금액",
    "bizno": "사업자번호",
    "corpNm": "업체명",
    "corpEntrprsDivNmNm": "기업구분",
    "cntrctDlvrReqNm": "계약명",
    "exclcProdctYn": "우수제품여부",
    "cnstwkMtrlDrctPurchsObjYn": "직접구매대상여부",
    "masYn": "MAS여부",
    "cntrctMthdNm": "계약방법",
    "IntlCntrctDlvrReqDate": "최초계약일자",
    "dlvrPlceNm": "납품장소",
    "dlvrTmlmtDate": "납품기한",
    "dlvryCndtnNm": "인도조건",
}

NUMERIC_COLS = ["단가", "수량", "공급금액"]


def to_dataframe(items):
    """API 응답 리스트를 정리된 DataFrame으로 변환"""
    if not items:
        return pd.DataFrame()
    df = pd.DataFrame(items)
    # 한글 컬럼으로 이름 변경
    df = df.rename(columns={k: v for k, v in FIELD_MAP.items() if k in df.columns})
    # 숫자 컬럼 변환
    for c in NUMERIC_COLS:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    # 날짜 컬럼 변환 (YYYYMMDD → datetime)
    if "계약일자" in df.columns:
        df["계약일자_dt"] = pd.to_datetime(df["계약일자"], format="%Y%m%d", errors="coerce")
    return df


def won(x):
    """원화 포맷"""
    try:
        return f"{int(x):,}원"
    except Exception:
        return str(x)


def eok(x):
    """억 단위 간략 표기"""
    try:
        v = float(x)
        if v >= 1e8:
            return f"{v/1e8:.1f}억원"
        if v >= 1e4:
            return f"{v/1e4:.0f}만원"
        return f"{v:.0f}원"
    except Exception:
        return str(x)


# ──────────────────────────────────────────────
# 사이드바: 검색 조건
# ──────────────────────────────────────────────
st.sidebar.title("🔍 검색 조건")

if not API_KEY:
    st.error("API 키가 설정되지 않았어요. .env 파일에 G2B_API_KEY를 넣어주세요.")
    st.stop()

# 품목 검색 방식
search_mode = st.sidebar.radio(
    "품목 검색 방식",
    ["품명으로 검색", "세부품명번호로 검색", "물품규격명으로 검색"],
    help="품명(예:합성목재) / 세부품명번호(예:3010369901) / 규격명 중 선택",
)

prdct_div = {"품명으로 검색": "1", "세부품명번호로 검색": "2", "물품규격명으로 검색": "3"}[search_mode]

# 검색어 입력
kwargs = dict(prdct_clsfc_no=None, prdct_clsfc_no_nm=None,
              dtil_prdct_clsfc_no=None, dtil_prdct_clsfc_no_nm=None,
              prdct_idnt_no=None, prdct_idnt_no_nm=None)

if search_mode == "품명으로 검색":
    term = st.sidebar.text_input("품명 또는 물품분류번호", placeholder="예: 합성목재 또는 30103699")
    if term:
        if term.isdigit():
            kwargs["prdct_clsfc_no"] = term
        else:
            kwargs["prdct_clsfc_no_nm"] = term
elif search_mode == "세부품명번호로 검색":
    term = st.sidebar.text_input("세부품명번호 또는 세부품명", placeholder="예: 3010369901 또는 합성목재")
    if term:
        if term.isdigit():
            kwargs["dtil_prdct_clsfc_no"] = term
        else:
            kwargs["dtil_prdct_clsfc_no_nm"] = term
else:
    term = st.sidebar.text_input("물품식별번호 또는 규격명", placeholder="예: 24267247 또는 데크재")
    if term:
        if term.isdigit():
            kwargs["prdct_idnt_no"] = term
        else:
            kwargs["prdct_idnt_no_nm"] = term

# 기간 선택 (최대 12개월)
st.sidebar.markdown("**조회 기간** (최대 12개월)")
today = date.today()
default_start = today - timedelta(days=90)
col_a, col_b = st.sidebar.columns(2)
bgn = col_a.date_input("시작일", value=default_start, max_value=today)
end = col_b.date_input("종료일", value=today, max_value=today)

# 일자 기준은 "계약(납품요구)일자"로 고정 (조달청 사이트와 동일)
inqry_div = "1"

# 최초계약여부 (조달청 사이트와 동일)
# 작동 원리: API 응답의 "변경차수(cntrctDlvrReqChgOrd)" 컬럼으로 자체 필터링
#  - Y (최초계약만): 변경차수 = 0
#  - N (최초 아님):  변경차수 ≠ 0 (1, 2, 3, ...)
#  - X (전체):       필터 없음
f_chg = st.sidebar.radio(
    "최초계약여부",
    ["전체 (X)", "Y", "N"],
    index=0,
    horizontal=True,
    help=(
        "조달청 사이트의 '최초계약여부'와 동일합니다.\n\n"
        "• 전체(X): 모든 차수 포함\n"
        "• Y: 최초계약만 (변경차수 = 0)\n"
        "• N: 최초계약 아님 (변경차수 ≠ 0)"
    ),
)

# 추가 필터
with st.sidebar.expander("➕ 추가 필터 (선택)"):
    f_exclc = st.selectbox("우수제품여부", ["전체", "우수제품만(Y)", "일반제품만(N)"])
    f_dminst = st.text_input("수요기관명 포함", placeholder="예: 교육청")
    f_corp = st.text_input("업체명 포함", placeholder="예: 더우드")
    f_region = st.text_input("수요기관지역", placeholder="예: 경기도")
    f_prcrmnt = st.selectbox("조달구분", ["전체", "중앙조달(C)", "자체조달(S)"])

extra = {}
# 최초계약여부는 API 파라미터로 전달하지 않고, 받은 데이터를 자체 필터링함 (아래 search 후 처리)
if f_exclc == "우수제품만(Y)":
    extra["exclcProdctYn"] = "Y"
elif f_exclc == "일반제품만(N)":
    extra["exclcProdctYn"] = "N"
if f_dminst:
    extra["dminsttNm"] = f_dminst
if f_corp:
    extra["corpNm"] = f_corp
if f_region:
    extra["dminsttRgnNm"] = f_region
if f_prcrmnt == "중앙조달(C)":
    extra["prcrmntDiv"] = "C"
elif f_prcrmnt == "자체조달(S)":
    extra["prcrmntDiv"] = "S"

search_btn = st.sidebar.button("🔎 검색하기", type="primary", use_container_width=True)

# ──────────────────────────────────────────────
# 메인 영역
# ──────────────────────────────────────────────
_logo_b64 = _load_logo_b64()
_logo_html = (
    f'<img src="data:image/png;base64,{_logo_b64}" '
    f'style="height: 48px; width: auto;" alt="ARCHIPACE" />'
) if _logo_b64 else (
    '<span style="font-size: 22px; font-weight: 700; color: #646464; '
    'letter-spacing: 3px;">ARCHI<span style="color: #DC6400;">|</span>PACE</span>'
)

st.markdown(f"""
<div class="archi-header">
    {_logo_html}
    <div class="archi-header-divider"></div>
    <div>
        <div class="archi-header-title">📊 조달청 특정품목 조달내역 분석</div>
        <div class="archi-header-subtitle">나라장터 조달데이터허브 · 특정품목조달내역 API 기반</div>
    </div>
    <div class="archi-header-meta">
        Procurement Market Intelligence<br/>
        <span style="color: #C8B89C;">v3.0 · 실시간 분석</span>
    </div>
</div>
""", unsafe_allow_html=True)

# 세션 상태로 데이터 보관
if "df" not in st.session_state:
    st.session_state.df = None

if search_btn:
    if not term:
        st.warning("품목 검색어를 입력해주세요.")
        st.stop()
    if (end - bgn).days > 366:
        st.warning("조회 기간은 최대 12개월까지 가능해요.")
        st.stop()
    if end < bgn:
        st.warning("종료일이 시작일보다 빠를 수 없어요.")
        st.stop()

    prog_bar = st.progress(0.0, text="데이터를 가져오는 중...")

    def _progress(page, total_pages, got, total):
        frac = page / max(total_pages, 1)
        prog_bar.progress(min(frac, 1.0), text=f"수집 중... {got:,} / {total:,}건")

    try:
        items, total = fetch_all(
            inqry_div=inqry_div,
            bgn_date=bgn.strftime("%Y%m%d"),
            end_date=end.strftime("%Y%m%d"),
            prdct_div=prdct_div,
            extra=extra,
            progress=_progress,
            **kwargs,
        )
        prog_bar.empty()
        df = to_dataframe(items)
        st.session_state.df = df
        st.session_state.last_term = term
        st.session_state.last_bgn = bgn.strftime("%Y-%m-%d")
        st.session_state.last_end = end.strftime("%Y-%m-%d")
        if df.empty:
            st.info("조건에 맞는 데이터가 없어요. 검색어나 기간을 바꿔보세요.")
    except Exception as e:
        prog_bar.empty()
        st.error(f"데이터를 가져오는 중 오류가 발생했어요:\n\n{e}")

df = st.session_state.df

if df is None:
    st.info("👈 왼쪽에서 품목과 기간을 정하고 **검색하기**를 눌러주세요.")
    st.stop()

if df.empty:
    st.stop()

# ──────────────────────────────────────────────
# 자체 필터링: 최초계약여부 (변경차수 = 0 기준)
# 조달청 사이트의 "최초계약여부 Y/N"와 동일한 결과를 자체적으로 구현
# ──────────────────────────────────────────────
if "변경차수" in df.columns:
    # 변경차수를 숫자로 변환 (문자/숫자/None 혼재 가능성 대응)
    chg_ord = pd.to_numeric(df["변경차수"], errors="coerce").fillna(-1)
    if f_chg == "Y":
        df = df[chg_ord == 0].reset_index(drop=True)
    elif f_chg == "N":
        df = df[chg_ord != 0].reset_index(drop=True)
    # 전체(X)는 필터 적용 안 함

if df.empty:
    st.info(f"'최초계약여부 = {f_chg}' 조건에 맞는 데이터가 없어요. 다른 옵션으로 시도해보세요.")
    st.stop()

# ──────────────────────────────────────────────
# 요약 지표
# ──────────────────────────────────────────────
total_amt = df["공급금액"].sum() if "공급금액" in df else 0
n_contracts = len(df)
n_corps = df["업체명"].nunique() if "업체명" in df else 0
n_insts = df["수요기관"].nunique() if "수요기관" in df else 0

c1, c2, c3, c4 = st.columns(4)
c1.metric("총 공급금액", eok(total_amt))
c2.metric("계약 건수", f"{n_contracts:,}건")
c3.metric("참여 업체 수", f"{n_corps:,}개")
c4.metric("수요기관 수", f"{n_insts:,}개")

st.divider()

# ──────────────────────────────────────────────
# 6개 탭
# ──────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "🏢 업체별 매출", "🏛️ 수요기관별", "📈 기간별 추이",
    "⭐ 우수제품 비중", "🗺️ 지역별 분포", "📋 전체 내역", "📝 보고서",
])

# ── 탭1: 업체별 매출 순위
with tab1:
    if "업체명" in df and "공급금액" in df:
        topn = st.slider("상위 몇 개 업체를 볼까요?", 5, 50, 15, key="t1")
        g = (df.groupby("업체명")["공급금액"].agg(["sum", "count"])
             .sort_values("sum", ascending=False).head(topn).reset_index())
        g.columns = ["업체명", "공급금액", "계약건수"]
        fig = px.bar(g, x="공급금액", y="업체명", orientation="h",
                     text=g["공급금액"].apply(eok),
                     color="공급금액", color_continuous_scale="Blues")
        fig.update_layout(yaxis={"categoryorder": "total ascending"}, height=max(400, topn * 28))
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(
            g.assign(공급금액=g["공급금액"].apply(won)),
            use_container_width=True, hide_index=True,
        )
    else:
        st.info("업체 또는 금액 데이터가 없어요.")

# ── 탭2: 수요기관별 구매액
with tab2:
    if "수요기관" in df and "공급금액" in df:
        topn = st.slider("상위 몇 개 기관을 볼까요?", 5, 50, 15, key="t2")
        g = (df.groupby("수요기관")["공급금액"].agg(["sum", "count"])
             .sort_values("sum", ascending=False).head(topn).reset_index())
        g.columns = ["수요기관", "구매액", "계약건수"]
        fig = px.bar(g, x="구매액", y="수요기관", orientation="h",
                     text=g["구매액"].apply(eok),
                     color="구매액", color_continuous_scale="Greens")
        fig.update_layout(yaxis={"categoryorder": "total ascending"}, height=max(400, topn * 28))
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(
            g.assign(구매액=g["구매액"].apply(won)),
            use_container_width=True, hide_index=True,
        )
    else:
        st.info("수요기관 또는 금액 데이터가 없어요.")

# ── 탭3: 기간별 추이
with tab3:
    if "계약일자_dt" in df and "공급금액" in df:
        freq = st.radio("집계 단위", ["일별", "주별", "월별"], index=2, horizontal=True, key="t3")
        # pandas 버전에 따라 월별 코드가 다름(M→ME). 둘 다 시도.
        rule_map = {"일별": "D", "주별": "W", "월별": "ME"}
        rule = rule_map[freq]
        base_ts = df.dropna(subset=["계약일자_dt"]).set_index("계약일자_dt")["공급금액"]
        try:
            ts = base_ts.resample(rule).agg(["sum", "count"]).reset_index()
        except ValueError:
            rule = {"일별": "D", "주별": "W", "월별": "M"}[freq]
            ts = base_ts.resample(rule).agg(["sum", "count"]).reset_index()
        ts.columns = ["일자", "공급금액", "계약건수"]
        fig = px.line(ts, x="일자", y="공급금액", markers=True)
        fig.update_traces(line_color="#2563eb")
        fig.update_layout(height=420, yaxis_title="공급금액(원)")
        st.plotly_chart(fig, use_container_width=True)
        fig2 = px.bar(ts, x="일자", y="계약건수")
        fig2.update_traces(marker_color="#93c5fd")
        fig2.update_layout(height=280, yaxis_title="계약 건수")
        st.plotly_chart(fig2, use_container_width=True)
    else:
        st.info("날짜 또는 금액 데이터가 없어요.")

# ── 탭4: 우수제품 vs 일반제품
with tab4:
    if "우수제품여부" in df and "공급금액" in df:
        df["_우수"] = df["우수제품여부"].map({"Y": "우수제품", "N": "일반제품"}).fillna("미표기")
        g = df.groupby("_우수")["공급금액"].agg(["sum", "count"]).reset_index()
        g.columns = ["구분", "공급금액", "계약건수"]
        col1, col2 = st.columns(2)
        with col1:
            fig = px.pie(g, names="구분", values="공급금액", hole=0.45,
                         title="공급금액 비중",
                         color_discrete_sequence=["#f59e0b", "#94a3b8", "#cbd5e1"])
            st.plotly_chart(fig, use_container_width=True)
        with col2:
            fig2 = px.pie(g, names="구분", values="계약건수", hole=0.45,
                          title="계약건수 비중",
                          color_discrete_sequence=["#f59e0b", "#94a3b8", "#cbd5e1"])
            st.plotly_chart(fig2, use_container_width=True)
        st.dataframe(
            g.assign(공급금액=g["공급금액"].apply(won)),
            use_container_width=True, hide_index=True,
        )
    else:
        st.info("우수제품 또는 금액 데이터가 없어요.")

# ── 탭5: 지역별 분포
with tab5:
    if "수요기관지역" in df and "공급금액" in df:
        # 시도 단위로 정리 (첫 단어 기준)
        df["_시도"] = df["수요기관지역"].astype(str).str.split().str[0]
        g = (df.groupby("_시도")["공급금액"].agg(["sum", "count"])
             .sort_values("sum", ascending=False).reset_index())
        g.columns = ["지역", "공급금액", "계약건수"]
        fig = px.bar(g, x="지역", y="공급금액", text=g["공급금액"].apply(eok),
                     color="공급금액", color_continuous_scale="Purples")
        fig.update_layout(height=420, yaxis_title="공급금액(원)")
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(
            g.assign(공급금액=g["공급금액"].apply(won)),
            use_container_width=True, hide_index=True,
        )
    else:
        st.info("지역 또는 금액 데이터가 없어요.")

# ── 탭6: 전체 내역 표 + 다운로드
with tab6:
    st.markdown("**전체 계약 내역** (표 안에서 정렬·검색 가능)")

    # 표시할 컬럼 우선순위
    preferred = ["계약일자", "품명", "세부품명", "물품규격명", "업체명", "기업구분",
                 "수요기관", "수요기관지역", "계약명", "단가", "수량", "단위",
                 "공급금액", "우수제품여부", "계약방법", "조달구분", "계약번호"]
    cols = [c for c in preferred if c in df.columns] + \
           [c for c in df.columns if c not in preferred and not c.startswith("_") and c != "계약일자_dt"]

    # 텍스트 필터
    q = st.text_input("표 안에서 검색 (업체/기관/계약명 등 아무 단어)", placeholder="예: 더우드")
    view = df[cols].copy()
    if q:
        mask = view.astype(str).apply(lambda r: r.str.contains(q, case=False, na=False)).any(axis=1)
        view = view[mask]

    st.caption(f"{len(view):,}건 표시 중")
    st.dataframe(view, use_container_width=True, hide_index=True, height=480)

    # 엑셀 다운로드
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_out = df[cols].copy()
        df_out.to_excel(writer, index=False, sheet_name="조달내역")
    buf.seek(0)
    fname = f"조달내역_{term}_{bgn.strftime('%Y%m%d')}_{end.strftime('%Y%m%d')}.xlsx"
    st.download_button("⬇️ 엑셀로 다운로드", data=buf, file_name=fname,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary")

# ── 탭7: 보고서
with tab7:
    st.markdown("### 📝 자동 분석 보고서")
    st.caption("검색된 데이터를 자동 분석해 요약·총평·표로 정리합니다.")

    # 검색 시점 정보를 세션에서 가져오기 (없으면 현재 입력값 사용)
    _term = st.session_state.get("last_term", term)
    _bgn = st.session_state.get("last_bgn", bgn.strftime("%Y-%m-%d"))
    _end = st.session_state.get("last_end", end.strftime("%Y-%m-%d"))

    analysis = build_analysis(df, _term, _bgn, _end)

    # ── 화면 표시 ──
    st.markdown(f"#### 📌 {analysis['검색어']}  ·  {analysis['기간']}")

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("총 공급금액", _r_eok(analysis["총공급금액"]))
    m2.metric("계약 건수", f"{analysis['계약건수']:,}건")
    m3.metric("참여 업체", f"{analysis['업체수']:,}개")
    m4.metric("수요기관", f"{analysis['수요기관수']:,}개")
    m5.metric("평균 계약액", _r_eok(analysis["평균계약금액"]))

    st.divider()

    # 총평 / 인사이트
    st.markdown("#### 🧭 분석 총평")
    for line in analysis.get("인사이트", []):
        st.markdown(f"- {line}")

    st.divider()

    # 표 형태 분석들
    col_left, col_right = st.columns(2)

    with col_left:
        if "상위업체" in analysis:
            st.markdown("#### 🏢 상위 공급업체")
            tb = analysis["상위업체"].copy()
            tb["공급금액"] = tb["sum"].apply(_r_won)
            tb["점유율"] = tb["점유율"].apply(lambda x: f"{x:.1f}%")
            tb["계약건수"] = tb["count"].apply(lambda x: f"{int(x):,}건")
            st.dataframe(tb[["업체명", "공급금액", "계약건수", "점유율"]],
                         use_container_width=True, hide_index=True)

        if "우수제품" in analysis:
            st.markdown("#### ⭐ 우수제품 비중")
            tb = analysis["우수제품"].copy()
            tb["금액"] = tb["금액"].apply(_r_won)
            tb["건수"] = tb["건수"].apply(lambda x: f"{int(x):,}건")
            st.dataframe(tb, use_container_width=True, hide_index=True)

        if "계약방식" in analysis:
            st.markdown("#### 📑 계약 방식별")
            tb = analysis["계약방식"].copy()
            tb["공급금액"] = tb["sum"].apply(_r_won)
            tb["계약건수"] = tb["count"].apply(lambda x: f"{int(x):,}건")
            st.dataframe(tb[["계약방법", "공급금액", "계약건수"]],
                         use_container_width=True, hide_index=True)

    with col_right:
        if "상위기관" in analysis:
            st.markdown("#### 🏛️ 상위 수요기관")
            tb = analysis["상위기관"].copy()
            tb["구매액"] = tb["sum"].apply(_r_won)
            tb["점유율"] = tb["점유율"].apply(lambda x: f"{x:.1f}%")
            tb["계약건수"] = tb["count"].apply(lambda x: f"{int(x):,}건")
            st.dataframe(tb[["수요기관", "구매액", "계약건수", "점유율"]],
                         use_container_width=True, hide_index=True)

        if "지역별" in analysis:
            st.markdown("#### 🗺️ 지역별 분포")
            tb = analysis["지역별"].copy()
            tb["공급금액"] = tb["sum"].apply(_r_won)
            tb["계약건수"] = tb["count"].apply(lambda x: f"{int(x):,}건")
            st.dataframe(tb[["지역", "공급금액", "계약건수"]],
                         use_container_width=True, hide_index=True)

    st.divider()

    # ── Word 다운로드 ──
    st.markdown("#### 📥 보고서 다운로드")
    st.caption("아래 버튼을 누르면 위 분석 내용이 담긴 Word 문서가 생성됩니다.")

    if st.button("📄 Word 보고서 만들기", type="primary"):
        with st.spinner("보고서를 생성하는 중..."):
            # 차트 이미지 생성 시도 (kaleido 있으면 포함, 없으면 표만)
            chart_imgs = []
            try:
                import plotly.io as pio
                if "상위업체" in analysis and len(analysis["상위업체"]) > 0:
                    gg = analysis["상위업체"].head(10)
                    f = px.bar(gg, x="sum", y="업체명", orientation="h",
                               color="sum", color_continuous_scale="Blues",
                               labels={"sum": "공급금액", "업체명": ""})
                    f.update_layout(yaxis={"categoryorder": "total ascending"},
                                    height=400, showlegend=False)
                    chart_imgs.append(("상위 공급업체", pio.to_image(f, format="png", scale=2)))
                if "상위기관" in analysis and len(analysis["상위기관"]) > 0:
                    gg = analysis["상위기관"].head(10)
                    f = px.bar(gg, x="sum", y="수요기관", orientation="h",
                               color="sum", color_continuous_scale="Greens",
                               labels={"sum": "구매액", "수요기관": ""})
                    f.update_layout(yaxis={"categoryorder": "total ascending"},
                                    height=400, showlegend=False)
                    chart_imgs.append(("상위 수요기관", pio.to_image(f, format="png", scale=2)))
            except Exception:
                chart_imgs = []  # kaleido 미설치 시 차트 없이 진행

            try:
                docx_buf = build_docx(analysis, chart_images=chart_imgs)
                rname = f"조달분석보고서_{analysis['검색어']}_{_bgn}_{_end}.docx"
                st.success("보고서가 생성되었어요! 아래 버튼으로 다운로드하세요.")
                if not chart_imgs:
                    st.info("ℹ️ 차트 이미지는 표 형태로 대체되었어요. "
                            "차트도 넣으려면 `pip install kaleido`를 설치하세요.")
                st.download_button("⬇️ Word 보고서 다운로드", data=docx_buf,
                                   file_name=rname,
                                   mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            except ModuleNotFoundError:
                st.error("python-docx가 설치되지 않았어요. "
                         "터미널에서 `pip install python-docx` 를 실행해주세요.")
            except Exception as e:
                st.error(f"보고서 생성 중 오류: {e}")

    # ── 엑셀 양식 보고서 (요청 양식 기반) ──
    st.divider()
    st.markdown("#### 📊 엑셀 양식 보고서 (월별 발주리스트)")
    st.caption("요청하신 엑셀 양식에 데이터를 자동으로 채워서 다운로드합니다. "
               "주차별 구분은 양식 그대로 유지되고, 데이터는 월별로 시간 순으로 채워집니다.")

    with st.expander("ℹ️ 양식의 월별 수용 가능 행 수 안내"):
        st.markdown(
            "- **1월**: 41행  |  **2월**: 70행  |  **3월**: 163행\n"
            "- **4월**: 216행  |  **5월**: 175행  |  **6월**: 62행\n"
            "- 이를 초과하는 데이터는 **'초과데이터' 시트**에 별도로 저장됩니다.\n"
            "- 1~6월 외(예: 7월) 데이터도 '초과데이터' 시트에 저장됩니다."
        )

    if st.button("📊 엑셀 양식 보고서 만들기", type="secondary"):
        with st.spinner("엑셀 양식에 데이터를 채우는 중..."):
            try:
                xlsx_buf = build_excel_report(df, _term, _bgn, _end)
                xname = f"발주리스트_{_term}_{_bgn}_{_end}.xlsx"
                st.success("✅ 엑셀 양식 보고서가 생성되었어요!")
                st.download_button(
                    "⬇️ 엑셀 양식 보고서 다운로드",
                    data=xlsx_buf,
                    file_name=xname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )
            except FileNotFoundError as e:
                st.error(
                    f"❌ {e}\n\n"
                    "**해결 방법**: `요청_양식_샘플_.xlsx` 파일을 `app.py`와 같은 폴더에 넣어주세요."
                )
            except Exception as e:
                st.error(f"엑셀 보고서 생성 중 오류: {e}")


# ──────────────────────────────────────────────
# 푸터
# ──────────────────────────────────────────────
st.markdown("""
<div class="archi-footer">
    <div>© 2026 ARCHIPACE · 조달 시장 분석 플랫폼</div>
    <div class="right">Powered by 나라장터 OpenAPI</div>
</div>
""", unsafe_allow_html=True)