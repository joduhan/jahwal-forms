"""
app.py — 2분할 레이아웃 (입력폼 | 문서 미리보기)
실행: streamlit run app.py
"""

import sys
import tempfile
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

sys.path.insert(0, str(Path(__file__).parent))

from modules.common_forms import (
    make_agreement, make_privacy_consent,
    make_integrity_pledge, make_all_common,
)
from modules.startup_forms import (
    make_application, make_settlement_report, make_all_startup,
)
from modules.report_gen import generate_prompt, make_report_docx
from modules.doc_builder import OUTPUTS_DIR

DEFAULT_CENTER_WIDE = "제주특별자치도광역자활센터"

# ──────────────────────────────────────────────
# 페이지 설정 & CSS
# ──────────────────────────────────────────────

st.set_page_config(
    page_title="제주광역자활센터 업무 자동화",
    page_icon="📋", layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
.stButton > button { background-color:#1B4F8A; color:white; border-radius:6px; }
.stButton > button:hover { background-color:#163f6f; }
.sec { font-size:1rem; font-weight:700; color:#1B4F8A;
       border-left:4px solid #1B4F8A; padding-left:8px; margin:1rem 0 0.3rem; }
.preview-container {
    background: white; padding: 36px 40px;
    box-shadow: 0 2px 12px rgba(0,0,0,0.15);
    min-height: 600px; font-family: '맑은 고딕', sans-serif;
    font-size: 10pt; line-height: 160%; color: #000;
    border-radius: 2px;
}
.preview-title { text-align:center; font-size:15pt; font-weight:bold; margin-bottom:16px; }
.preview-sub { text-align:center; font-size:10pt; margin-bottom:14px; }
.preview-section { font-weight:bold; margin:12px 0 4px; font-size:10pt; }
.pt { width:100%; border-collapse:collapse; margin-bottom:12px; font-size:9pt; }
.pt td,.pt th { border:1px solid #000; padding:5px 8px; }
.pt th { background:#D9D9D9; font-weight:bold; text-align:center; }
.pt .lbl { background:#EFEFEF; font-weight:bold; text-align:center; width:35%; }
.sign { margin-top:20px; text-align:right; font-size:10pt; line-height:200%; }
.ph { color:#bbb; font-style:italic; }
</style>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────
# 공통 유틸
# ──────────────────────────────────────────────

def _sec(t): st.markdown(f'<p class="sec">{t}</p>', unsafe_allow_html=True)

def _dl_buttons(p: Path):
    c1, c2 = st.columns(2)
    with open(p, "rb") as f:
        c1.download_button("⬇️ DOCX", f.read(), p.name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    pdf = p.with_suffix(".pdf")
    if pdf.exists():
        with open(pdf, "rb") as f:
            c2.download_button("⬇️ PDF", f.read(), pdf.name, mime="application/pdf")
    else:
        c2.caption("PDF: LibreOffice 필요")

def _v(val, ph=""):
    """값이 있으면 값, 없으면 회색 placeholder span."""
    s = str(val).strip() if val else ""
    return s if s else f'<span class="ph">{ph}</span>'

def _preview(html: str):
    st.markdown(f'<div class="preview-container">{html}</div>', unsafe_allow_html=True)

def _tr2(label, val, ph=""):
    return f'<tr><td class="lbl">{label}</td><td>{_v(val, ph)}</td></tr>'

# ──────────────────────────────────────────────
# 미리보기 HTML 함수
# ──────────────────────────────────────────────

def preview_html_agreement(d: dict, party_type: str) -> str:
    is3 = party_type == "3way"
    party_rows = f"""
    <tr><td class="lbl">광역자활센터명</td><td>{_v(d.get('center_wide'),'광역자활센터명')}</td></tr>
    <tr><td class="lbl">광역 대표자</td><td>{_v(d.get('rep_wide'),'대표자')}</td></tr>
    <tr><td class="lbl">광역 사업자등록번호</td><td>{_v(d.get('reg_wide'),'000-00-00000')}</td></tr>
    <tr><td class="lbl">지역자활센터명</td><td>{_v(d.get('center_local'),'지역자활센터명')}</td></tr>
    <tr><td class="lbl">지역 대표자</td><td>{_v(d.get('rep_local'),'대표자')}</td></tr>
    <tr><td class="lbl">지역 사업자등록번호</td><td>{_v(d.get('reg_local'),'000-00-00000')}</td></tr>
    """
    if is3:
        party_rows += f"""
        <tr><td class="lbl">자활기업명</td><td>{_v(d.get('enterprise_name'),'자활기업명')}</td></tr>
        <tr><td class="lbl">기업 대표자</td><td>{_v(d.get('rep_enterprise'),'대표자')}</td></tr>
        <tr><td class="lbl">기업 사업자번호</td><td>{_v(d.get('reg_enterprise'),'000-00-00000')}</td></tr>
        """
    sign_line = f"""
    {_v(d.get('center_wide'),'광역자활센터')} 대표자: {_v(d.get('rep_wide'),'')}&nbsp;(인)<br>
    {_v(d.get('center_local'),'지역자활센터')} 대표자: {_v(d.get('rep_local'),'')}&nbsp;(인)
    """ + (f"<br>{_v(d.get('enterprise_name'),'자활기업')} 대표자: {_v(d.get('rep_enterprise'),'')}&nbsp;(인)" if is3 else "")

    return f"""
    <div class="preview-title">[{'3자' if is3 else '2자'}용] 표준협약서</div>
    <div style="text-align:right;font-size:9pt;margin-bottom:8px">
        협약번호: {_v(d.get('contract_no'),'제2025-___호')}
    </div>
    <p class="preview-section">■ 협약 당사자</p>
    <table class="pt"><tbody>{party_rows}</tbody></table>
    <p class="preview-section">■ 협약 내용</p>
    <table class="pt"><tbody>
        {_tr2('지원내용', d.get('support_content'), '지원내용')}
        <tr><td class="lbl">협약기간</td>
            <td>{_v(d.get('period_start'),'시작일')} ~ {_v(d.get('period_end'),'종료일')}</td></tr>
    </tbody></table>
    <p class="preview-section">■ 예산 현황</p>
    <table class="pt">
        <tr><th>중앙자활자금</th><th>자부담</th><th>자활기금</th><th>지역자활사업지원비</th></tr>
        <tr>
            <td style="text-align:center">{_v(d.get('amount_central'),'0')}</td>
            <td style="text-align:center">{_v(d.get('amount_self'),'0')}</td>
            <td style="text-align:center">{_v(d.get('amount_fund'),'0')}</td>
            <td style="text-align:center">{_v(d.get('amount_local_support'),'0')}</td>
        </tr>
    </table>
    <div class="sign">2025년&nbsp;&nbsp;&nbsp;월&nbsp;&nbsp;&nbsp;일<br>{sign_line}</div>
    """


def preview_html_privacy(d: dict) -> str:
    return f"""
    <div class="preview-title">개인정보 수집 및 이용 동의서</div>
    <div class="preview-sub">({_v(d.get('project_name'),'지원사업명')})</div>
    <p>아래와 같이 개인정보를 수집 및 이용하고자 하오니, 자세히 읽어보신 후 동의여부를 결정하여 주시기 바랍니다.</p>
    <p class="preview-section">○ 필수항목</p>
    <table class="pt"><tbody>
        <tr><td class="lbl">수집항목<br>(광역자활센터)</td>
            <td>성명, 연락처, 생년월일, 자격 및 경력사항(자격증 및 경력증명서 사본),<br>
                고용보험 가입이력, 가족관계(가족관계증명서)</td></tr>
        <tr><td class="lbl">수집항목<br>(한국자활복지개발원)</td>
            <td>성명, 연락처, 생년월일, 자격 및 경력사항</td></tr>
        <tr><td class="lbl">수집목적</td>
            <td>중앙자활자금 지원사업 참여를 위한 채용 승인 검토 및 사업 관리</td></tr>
        <tr><td class="lbl">보유기간</td><td>수집일로부터 5년</td></tr>
        <tr><td class="lbl">거부 시 불이익</td>
            <td>동의를 거부할 경우 중앙자활자금 지원사업 선정에 제외될 수 있습니다.</td></tr>
    </tbody></table>
    <p><strong>위와 같이 개인정보를 수집 및 이용하는 것에 동의하십니까?</strong></p>
    <p style="text-align:center">□ 동의함 &nbsp;&nbsp;&nbsp;&nbsp; □ 동의하지 않음</p>
    <div class="sign">
        {_v(d.get('sign_date'),'2025')}년&nbsp;&nbsp;&nbsp;월&nbsp;&nbsp;&nbsp;일<br>
        이&nbsp;&nbsp;&nbsp;&nbsp;름 :&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(날인)<br><br>
        한국자활복지개발원장 · {_v(d.get('center_wide'), DEFAULT_CENTER_WIDE)}장 귀하
    </div>
    """


def preview_html_integrity(d: dict) -> str:
    cw = _v(d.get('center_wide'), DEFAULT_CENTER_WIDE)
    return f"""
    <div class="preview-title">청&nbsp;&nbsp;렴&nbsp;&nbsp;이&nbsp;&nbsp;행&nbsp;&nbsp;각&nbsp;&nbsp;서</div>
    <p>우리 기관은 부패없는 투명하고 공정한 행정이 사회발전과 국가경쟁력에 중요한 관건이 됨을
    깊이 인식하고, 청렴이행각서 작성 취지에 적극 호응하며 한국자활복지개발원({cw})에서
    실시하는 지원사업을 참여함에 있어 우리 기관의 직원과 대리인은</p>
    <p style="padding-left:1em">1. 한국자활복지개발원({cw}) 및 관련기관 임직원에게 금품·향응 등을
    제공하지 않겠으며, 위반 시 향후 3년간 관련 지원사업 참여를 제한 받는 것에 동의합니다.</p>
    <p style="padding-left:1em">2. 상기와 같은 행위로 선정이 취소되거나 협약이 해제되더라도
    이에 대한 민·형사상 일체의 이의를 제기하지 않겠습니다.</p>
    <p style="padding-left:1em">3. 지원된 예산은 목적 외에 사용하지 않겠으며, 위반 사항이 발생할
    경우 반납조치를 감수하겠습니다.</p>
    <p>위 청렴이행각서는 상호신뢰를 바탕으로 한 약속으로서 반드시 지킬 것이며,
    이를 위반할 경우 어떠한 제재도 감수할 것을 서약합니다.</p>
    <div class="sign">
        {_v(d.get('sign_date'),'2025')}년&nbsp;&nbsp;&nbsp;월&nbsp;&nbsp;&nbsp;일<br>
        서약기관명: {_v(d.get('org_name'),'기관명')}&nbsp;&nbsp;대표 {_v(d.get('rep_name'),'성명')}&nbsp;(서명 또는 인)<br><br>
        {cw}장 귀중
    </div>
    """


def preview_html_application(d: dict) -> str:
    loan = d.get('amount_loan', '0')
    grant = d.get('amount_grant', '0')
    try:
        total = int(str(loan).replace(',','')) + int(str(grant).replace(',',''))
        total = f"{total:,}"
    except Exception:
        total = ''
    etype = d.get('enterprise_type', '')
    types = ['주식회사','법인사업자','협동조합','개인기업']
    cb_row = '  '.join(f"{'■' if t==etype else '□'} {t}" for t in types)
    rtype = d.get('repayment_type','')
    rc_row = '  '.join(f"{'■' if r==rtype else '□'} {r}" for r in ['분기균등분할상환','만기일시상환'])

    return f"""
    <div class="preview-title">자활기업 창업자금 사업신청서</div>
    <div class="preview-sub">(자활정보시스템 등록)</div>
    <p class="preview-section">■ 신청기관 정보</p>
    <table class="pt"><tbody>
        {_tr2('기관명', d.get('center_local'), '지역자활센터명')}
        {_tr2('사업자등록번호', d.get('reg_no'), '000-00-00000')}
        {_tr2('기관 대표자', d.get('rep_name'), '대표자명')}
        {_tr2('대표자 연락처', d.get('rep_phone'), '000-0000-0000')}
        {_tr2('기관주소', d.get('address'), '주소')}
        {_tr2('사업담당자', d.get('manager_name'), '담당자명')}
    </tbody></table>
    <p class="preview-section">■ 사업단 현황</p>
    <table class="pt"><tbody>
        {_tr2('사업단명', d.get('unit_name'), '사업단명')}
        {_tr2('사업유형', d.get('unit_type'), '유형')}
        {_tr2('운영기간', d.get('unit_period'), '기간')}
        {_tr2('3년간 매출실적(연평균)', d.get('sales_avg'), '금액')}
        {_tr2('최근 3개월 매출실적', d.get('sales_recent'), '금액')}
    </tbody></table>
    <p class="preview-section">■ 자활기업 정보</p>
    <table class="pt"><tbody>
        {_tr2('자활기업명', d.get('enterprise_name'), '기업명')}
        {_tr2('기업 대표자', d.get('enterprise_rep'), '대표자명')}
        {_tr2('기업형태', cb_row, '')}
    </tbody></table>
    <p class="preview-section">■ 신청금액</p>
    <table class="pt"><tbody>
        {_tr2('융자(임대보증금)', loan, '0')}
        {_tr2('무상지원(운영자금)', grant, '0')}
        {_tr2('합계', total, '자동계산')}
        {_tr2('융자상환계획', rc_row, '')}
    </tbody></table>
    """


def preview_html_settlement(d: dict, report_type: str) -> str:
    titles = {
        "1st": "자활기업 창업자금 지원 기업 1차 중간 집행정산 보고",
        "2nd": "자활기업 창업자금 지원 기업 2차 중간 집행정산 보고",
        "final": "자활기업 창업자금 지원 기업 최종 집행정산 보고",
    }
    exec_lbl = {"1st":"협약일로부터 6개월 이내","2nd":"협약일로부터 12개월 이내","final":"협약기간 18개월"}
    return f"""
    <div class="preview-title" style="font-size:13pt">{titles.get(report_type,'')}</div>
    <p class="preview-section">■ 기관 및 기업 정보</p>
    <table class="pt"><tbody>
        {_tr2('기  관  명', d.get('center_local'), '지역자활센터명')}
        {_tr2('대  표  자', d.get('rep_name'), '대표자')}
        {_tr2('주      소', d.get('address'), '주소')}
        {_tr2('자활기업명', d.get('enterprise_name'), '기업명')}
        {_tr2('기업인정일자', d.get('certified_date'), '날짜')}
        {_tr2('사업기간', d.get('business_period'), '기간')}
        {_tr2('협약기간', d.get('contract_period'), '기간')}
        {_tr2('집행보고기간', exec_lbl.get(report_type,''), '')}
    </tbody></table>
    <p class="preview-section">■ 지원금액 및 집행현황</p>
    <table class="pt">
        <tr><th>구분</th><th>지원금액</th><th>집행금액</th></tr>
        <tr>
            <td style="text-align:center">융자(임대보증금)</td>
            <td style="text-align:center">{_v(d.get('amount_loan'),'0')}</td>
            <td style="text-align:center">{_v(d.get('exec_loan'),'0')}</td>
        </tr>
        <tr>
            <td style="text-align:center">무상지원(운영자금)</td>
            <td style="text-align:center">{_v(d.get('amount_grant'),'0')}</td>
            <td style="text-align:center">{_v(d.get('exec_grant'),'0')}</td>
        </tr>
    </table>
    <div class="sign">
        {_v(d.get('sign_date'),'2025')}년&nbsp;&nbsp;&nbsp;월&nbsp;&nbsp;&nbsp;일<br>
        수행기관명: {_v(d.get('center_local'),'')}<br>
        한국자활복지개발원장 귀중
    </div>
    """

# ──────────────────────────────────────────────
# 사이드바
# ──────────────────────────────────────────────

MENU = {
    "🏠 홈": "home", "📄 표준협약서": "agreement",
    "📄 개인정보 동의서": "privacy", "📄 청렴이행각서": "pledge",
    "📋 사업신청서": "application", "📋 중간집행정산보고": "settlement",
    "📊 결과보고서": "report", "📦 일괄 생성": "batch",
}
with st.sidebar:
    st.markdown("### 🏛️ 제주광역자활센터")
    st.markdown("**업무 자동화 시스템**")
    st.divider()
    selected = st.radio("메뉴", list(MENU.keys()), label_visibility="collapsed")
    page = MENU[selected]
    st.divider()
    st.caption("ⓒ 2025 제주특별자치도광역자활센터")

# ──────────────────────────────────────────────
# 홈 (전체 너비)
# ──────────────────────────────────────────────

def page_home():
    st.title("📋 제주광역자활센터 업무 자동화 시스템")
    st.markdown("2025년도 중앙자활자금 개발원·광역 지원사업 서식을 자동으로 생성합니다.")
    _sec("생성 가능한 서식")
    st.dataframe(pd.DataFrame([
        ("공통3","표준협약서","📄 표준협약서"),
        ("공통4","개인정보 수집·이용 동의서","📄 개인정보 동의서"),
        ("공통5","청렴이행각서","📄 청렴이행각서"),
        ("창업1","사업신청서","📋 사업신청서"),
        ("창업2","집행정산보고","📋 중간집행정산보고"),
        ("보고","결과보고서","📊 결과보고서"),
    ], columns=["서식번호","서식명","메뉴"]), use_container_width=True, hide_index=True)
    _sec("최근 생성 파일")
    OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)
    recent = sorted(OUTPUTS_DIR.glob("*.docx"), key=lambda p: p.stat().st_mtime, reverse=True)[:5]
    for f in recent:
        st.text(f"📄 {f.name}  ({pd.Timestamp(f.stat().st_mtime, unit='s').strftime('%Y-%m-%d %H:%M')})")
    if not recent:
        st.info("아직 생성된 파일이 없습니다.")

# ──────────────────────────────────────────────
# 표준협약서
# ──────────────────────────────────────────────

def page_agreement():
    st.title("📄 표준협약서")
    party_type_label = st.radio("협약 유형", ["2자용 (광역 + 지역)","3자용 (광역 + 지역 + 자활기업)"], horizontal=True, key="ag_ptype")
    is3 = "3자용" in party_type_label
    party_type = "3way" if is3 else "2way"

    col1, col2 = st.columns([4, 6])
    with col1:
        _sec("협약 기본정보")
        center_wide     = st.text_input("광역자활센터명 *", value=DEFAULT_CENTER_WIDE, key="ag_cw")
        contract_no     = st.text_input("협약번호", placeholder="예: 제2025-001호", key="ag_no")
        support_content = st.text_input("지원내용 *", placeholder="예: 자활기업 창업자금 지원", key="ag_sc")
        c1, c2 = st.columns(2)
        period_start    = c1.text_input("협약기간 시작 *", placeholder="예: 2025.01.01.", key="ag_ps")
        period_end      = c2.text_input("협약기간 종료 *", placeholder="예: 2025.12.31.", key="ag_pe")
        _sec("광역자활센터 정보")
        ca, cb = st.columns(2)
        rep_wide = ca.text_input("광역 대표자 *", key="ag_rw")
        reg_wide = cb.text_input("광역 사업자등록번호 *", placeholder="123-45-67890", key="ag_rgw")
        _sec("지역자활센터 정보")
        center_local = st.text_input("지역자활센터명 *", key="ag_cl")
        cc, cd = st.columns(2)
        rep_local = cc.text_input("지역 대표자 *", key="ag_rl")
        reg_local = cd.text_input("지역 사업자등록번호 *", placeholder="234-56-78901", key="ag_rgl")
        enterprise_name = rep_enterprise = reg_enterprise = ""
        if is3:
            _sec("자활기업 정보")
            enterprise_name = st.text_input("자활기업명 *", key="ag_en")
            ce, cf = st.columns(2)
            rep_enterprise = ce.text_input("기업 대표자 *", key="ag_re")
            reg_enterprise = cf.text_input("기업 사업자등록번호 *", key="ag_rge")
        _sec("예산 현황 (원)")
        e1, e2, e3, e4 = st.columns(4)
        amount_central       = e1.text_input("중앙자활자금",         placeholder="5,000,000", key="ag_ac")
        amount_self          = e2.text_input("자부담",               placeholder="1,000,000", key="ag_as")
        amount_fund          = e3.text_input("자활기금",              placeholder="0",         key="ag_af")
        amount_local_support = e4.text_input("지역자활사업지원비",    placeholder="0",         key="ag_al")

        gen = st.button("📄 문서 생성", key="ag_gen")

    d = dict(center_wide=center_wide, center_local=center_local,
             enterprise_name=enterprise_name, rep_wide=rep_wide, rep_local=rep_local,
             rep_enterprise=rep_enterprise, reg_wide=reg_wide, reg_local=reg_local,
             reg_enterprise=reg_enterprise, support_content=support_content,
             period_start=period_start, period_end=period_end,
             amount_central=amount_central, amount_self=amount_self,
             amount_fund=amount_fund, amount_local_support=amount_local_support,
             contract_no=contract_no)

    with col2:
        st.markdown("#### 📄 문서 미리보기")
        _preview(preview_html_agreement(d, party_type))
        if "ag_path" in st.session_state:
            st.divider()
            _dl_buttons(st.session_state["ag_path"])

    if gen:
        required = [center_wide, support_content, period_start, period_end,
                    rep_wide, reg_wide, center_local, rep_local, reg_local]
        if is3: required += [enterprise_name, rep_enterprise, reg_enterprise]
        if not all(required):
            col1.error("⚠️ 필수 항목(*)을 모두 입력해 주세요.")
        else:
            with st.spinner("문서를 생성하는 중..."):
                try:
                    path = make_agreement(party_type=party_type, data=d)
                    st.session_state["ag_path"] = path
                    st.success(f"✅ 생성 완료: `{path.name}`")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ 오류: {e}")

# ──────────────────────────────────────────────
# 개인정보 동의서
# ──────────────────────────────────────────────

def page_privacy():
    st.title("📄 개인정보 수집·이용 동의서")
    col1, col2 = st.columns([4, 6])
    with col1:
        c1, c2 = st.columns(2)
        project_name = c1.text_input("지원사업명 *", placeholder="예: 자활기업 창업 지원사업", key="pv_pn")
        center_wide  = c2.text_input("광역자활센터명 *", value=DEFAULT_CENTER_WIDE, key="pv_cw")
        c3, c4 = st.columns(2)
        signee_name  = c3.text_input("동의자 성명 *", key="pv_sn")
        sign_date    = c4.text_input("날짜(연도)", value=str(date.today().year), key="pv_sd")
        gen = st.button("📄 문서 생성", key="pv_gen")

    d = dict(project_name=project_name, center_wide=center_wide,
             signee_name=signee_name, sign_date=sign_date)
    with col2:
        st.markdown("#### 📄 문서 미리보기")
        _preview(preview_html_privacy(d))
        if "pv_path" in st.session_state:
            st.divider()
            _dl_buttons(st.session_state["pv_path"])

    if gen:
        if not all([project_name, center_wide, signee_name]):
            st.error("⚠️ 필수 항목(*)을 모두 입력해 주세요.")
        else:
            with st.spinner("문서를 생성하는 중..."):
                try:
                    path = make_privacy_consent(d)
                    st.session_state["pv_path"] = path
                    st.success(f"✅ 생성 완료: `{path.name}`")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ 오류: {e}")

# ──────────────────────────────────────────────
# 청렴이행각서
# ──────────────────────────────────────────────

def page_pledge():
    st.title("📄 청렴이행각서")
    col1, col2 = st.columns([4, 6])
    with col1:
        center_wide = st.text_input("광역자활센터명 *", value=DEFAULT_CENTER_WIDE, key="pl_cw")
        c1, c2 = st.columns(2)
        org_name  = c1.text_input("서약기관명 *", placeholder="예: 제주시지역자활센터", key="pl_on")
        rep_name  = c2.text_input("대표자명 *", key="pl_rn")
        sign_date = st.text_input("날짜(연도)", value=str(date.today().year), key="pl_sd")
        gen = st.button("📄 문서 생성", key="pl_gen")

    d = dict(center_wide=center_wide, org_name=org_name, rep_name=rep_name, sign_date=sign_date)
    with col2:
        st.markdown("#### 📄 문서 미리보기")
        _preview(preview_html_integrity(d))
        if "pl_path" in st.session_state:
            st.divider()
            _dl_buttons(st.session_state["pl_path"])

    if gen:
        if not all([center_wide, org_name, rep_name]):
            st.error("⚠️ 필수 항목(*)을 모두 입력해 주세요.")
        else:
            with st.spinner("문서를 생성하는 중..."):
                try:
                    path = make_integrity_pledge(d)
                    st.session_state["pl_path"] = path
                    st.success(f"✅ 생성 완료: `{path.name}`")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ 오류: {e}")

# ──────────────────────────────────────────────
# 사업신청서
# ──────────────────────────────────────────────

def page_application():
    st.title("📋 자활기업 창업자금 사업신청서")
    col1, col2 = st.columns([4, 6])
    with col1:
        _sec("표1. 신청기관 정보")
        c1,c2=st.columns(2)
        center_local  = c1.text_input("지역자활센터명 *",  key="ap_cl")
        reg_no        = c2.text_input("사업자등록번호 *",  key="ap_rn", placeholder="123-45-67890")
        c3,c4=st.columns(2)
        rep_name      = c3.text_input("기관 대표자 *",     key="ap_rm")
        rep_phone     = c4.text_input("대표자 연락처 *",   key="ap_rp", placeholder="064-000-0000")
        address       = st.text_input("기관주소 *",        key="ap_ad")
        c5,c6=st.columns(2)
        email         = c5.text_input("이메일",            key="ap_em")
        manager_name  = c6.text_input("사업담당자 *",      key="ap_mn")
        manager_phone = st.text_input("담당자 휴대전화 *", key="ap_mp", placeholder="010-0000-0000")

        _sec("표2. 사업단 현황")
        c7,c8=st.columns(2)
        unit_name     = c7.text_input("사업단명 *",         key="ap_un")
        unit_type     = c8.text_input("사업유형 *",         key="ap_ut", placeholder="청소·환경")
        unit_period   = st.text_input("운영기간 *",         key="ap_up", placeholder="2022.01 ~ 2024.12")
        c9,c10=st.columns(2)
        sales_avg     = c9.text_input("3년간 매출실적(연평균)", key="ap_sa", placeholder="12,000,000")
        sales_recent  = c10.text_input("최근 3개월 매출실적",  key="ap_sr", placeholder="3,500,000")
        c11,c12=st.columns(2)
        participants_avg      = c11.text_input("참여자수(월평균)",       key="ap_pa", placeholder="5명")
        startup_fund_saved    = c12.text_input("기적립 창업자금(100%)", key="ap_sf", placeholder="1,000,000")
        c13,c14=st.columns(2)
        startup_fund_approved = c13.text_input("기승인 창업자금 금액",  key="ap_sfa", placeholder="0")
        approval_doc_no       = c14.text_input("지자체 승인 문서번호",  key="ap_adn")
        approved_items        = st.text_input("기승인 창업자금 사용항목", key="ap_ai")

        _sec("표3. 자활기업 정보")
        c15,c16=st.columns(2)
        enterprise_name           = c15.text_input("자활기업명 *",       key="ap_en")
        enterprise_reg_no         = c16.text_input("기업 사업자번호",    key="ap_ern")
        c17,c18=st.columns(2)
        enterprise_rep            = c17.text_input("기업 대표자 *",      key="ap_er")
        enterprise_certified_date = c18.text_input("기업인정(예정)일",   key="ap_ecd", placeholder="2025.01.01.")
        c19,c20=st.columns(2)
        members_basic     = c19.text_input("기초생활보장 수급자", key="ap_mb", placeholder="3명(60%)")
        members_near_poor = c20.text_input("차상위자",            key="ap_mn2", placeholder="2명(40%)")
        enterprise_type   = st.selectbox("기업형태 *", ["주식회사","법인사업자","협동조합","개인기업"], key="ap_et")

        _sec("표4. 신청금액 (원)")
        c21,c22=st.columns(2)
        amount_loan   = c21.text_input("융자(임대보증금)",  key="ap_al", placeholder="2,000,000")
        amount_grant  = c22.text_input("무상지원(운영자금)", key="ap_ag", placeholder="3,000,000")
        repayment_type = st.radio("융자상환계획", ["분기균등분할상환","만기일시상환"], horizontal=True, key="ap_rt")
        repayment_reason = ""
        if repayment_type == "만기일시상환":
            repayment_reason = st.text_input("만기일시상환 사유 *", key="ap_rr")

        gen = st.button("📄 문서 생성", key="ap_gen")

    d = dict(
        center_local=center_local, reg_no=reg_no, rep_name=rep_name, rep_phone=rep_phone,
        address=address, email=email, manager_name=manager_name, manager_phone=manager_phone,
        unit_name=unit_name, unit_type=unit_type, unit_period=unit_period,
        sales_avg=sales_avg, sales_recent=sales_recent, participants_avg=participants_avg,
        startup_fund_saved=startup_fund_saved, startup_fund_approved=startup_fund_approved,
        approval_doc_no=approval_doc_no, approved_items=approved_items,
        enterprise_name=enterprise_name, enterprise_reg_no=enterprise_reg_no,
        enterprise_rep=enterprise_rep, enterprise_certified_date=enterprise_certified_date,
        members_basic=members_basic, members_near_poor=members_near_poor,
        enterprise_type=enterprise_type, amount_loan=amount_loan, amount_grant=amount_grant,
        repayment_type=repayment_type, repayment_reason=repayment_reason,
    )
    with col2:
        st.markdown("#### 📄 문서 미리보기")
        _preview(preview_html_application(d))
        if "ap_path" in st.session_state:
            st.divider()
            _dl_buttons(st.session_state["ap_path"])

    if gen:
        required = [center_local, reg_no, rep_name, rep_phone, address,
                    manager_name, manager_phone, unit_name, unit_type, unit_period,
                    enterprise_name, enterprise_rep]
        if repayment_type == "만기일시상환": required.append(repayment_reason)
        if not all(required):
            st.error("⚠️ 필수 항목(*)을 모두 입력해 주세요.")
        else:
            with st.spinner("문서를 생성하는 중..."):
                try:
                    path = make_application(d)
                    st.session_state["ap_path"] = path
                    st.success(f"✅ 생성 완료: `{path.name}`")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ 오류: {e}")

# ──────────────────────────────────────────────
# 집행정산보고
# ──────────────────────────────────────────────

def page_settlement():
    st.title("📋 집행정산보고")
    rtype_label = st.radio("보고 유형", ["1차 중간 집행정산","2차 중간 집행정산","최종 집행정산"],
                           horizontal=True, key="st_rt")
    RMAP = {"1차 중간 집행정산":"1st","2차 중간 집행정산":"2nd","최종 집행정산":"final"}
    report_type = RMAP[rtype_label]

    col1, col2 = st.columns([4, 6])
    with col1:
        _sec("기관 정보")
        c1,c2=st.columns(2)
        center_local = c1.text_input("지역자활센터명 *", key="st_cl")
        rep_name     = c2.text_input("대표자 *",         key="st_rn")
        address      = st.text_input("주소 *",           key="st_ad")
        c3,c4=st.columns(2)
        phone = c3.text_input("전화번호 *", key="st_ph", placeholder="064-000-0000")
        email = c4.text_input("E-mail",     key="st_em")
        _sec("자활기업 정보")
        c5,c6=st.columns(2)
        enterprise_name  = c5.text_input("자활기업명 *",  key="st_en")
        certified_date   = c6.text_input("기업인정일자",  key="st_cd", placeholder="2025.01.01.")
        business_period  = st.text_input("사업기간",      key="st_bp", placeholder="2025.01.01. ~ 2025.12.31.")
        contract_period  = st.text_input("협약기간",      key="st_cp", placeholder="2025.01.01. ~ 2026.06.30.")
        report_period    = st.text_input("사업보고기간",  key="st_rp", placeholder="2025.01.01. ~ 2025.06.30.")
        _sec("지원금액 및 집행현황 (원)")
        c7,c8,c9,c10=st.columns(4)
        amount_loan  = c7.text_input("융자 지원금액",  key="st_al",  placeholder="2,000,000")
        amount_grant = c8.text_input("무상 지원금액",  key="st_ag",  placeholder="3,000,000")
        exec_loan    = c9.text_input("융자 집행금액",  key="st_el",  placeholder="2,000,000")
        exec_grant   = c10.text_input("무상 집행금액", key="st_eg",  placeholder="1,500,000")
        _sec("작성자 정보")
        c11,c12=st.columns(2)
        manager_name  = c11.text_input("작성자(담당자) *", key="st_mn")
        manager_phone = c12.text_input("작성자 연락처 *",  key="st_mp", placeholder="010-0000-0000")
        sign_date = st.text_input("날짜(연도)", value=str(date.today().year), key="st_sd")
        gen = st.button("📄 문서 생성", key="st_gen")

    d = dict(center_local=center_local, rep_name=rep_name, address=address,
             phone=phone, email=email, enterprise_name=enterprise_name,
             certified_date=certified_date, business_period=business_period,
             contract_period=contract_period, report_period=report_period,
             amount_loan=amount_loan, amount_grant=amount_grant,
             exec_loan=exec_loan, exec_grant=exec_grant,
             manager_name=manager_name, manager_phone=manager_phone, sign_date=sign_date)
    with col2:
        st.markdown("#### 📄 문서 미리보기")
        _preview(preview_html_settlement(d, report_type))
        if "st_path" in st.session_state:
            st.divider()
            _dl_buttons(st.session_state["st_path"])

    if gen:
        if not all([center_local, rep_name, address, phone, enterprise_name,
                    manager_name, manager_phone]):
            st.error("⚠️ 필수 항목(*)을 모두 입력해 주세요.")
        else:
            with st.spinner("문서를 생성하는 중..."):
                try:
                    path = make_settlement_report(report_type, d)
                    st.session_state["st_path"] = path
                    st.success(f"✅ 생성 완료: `{path.name}`")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ 오류: {e}")

# ──────────────────────────────────────────────
# 결과보고서 (3단계, Claude.ai 플랜 방식)
# ──────────────────────────────────────────────

def page_report():
    st.title("📊 결과보고서")
    if "report_step" not in st.session_state:
        st.session_state.report_step = 1
    step = st.session_state.report_step
    labels = ["1단계: 데이터 입력","2단계: Claude.ai 연동","3단계: 문서 완성"]
    st.progress((step-1)/2, text=labels[step-1])

    col1, col2 = st.columns([4, 6])

    if step == 1:
        with col1:
            c1,c2=st.columns(2)
            사업명   = c1.text_input("사업명 *",   key="rp_sn")
            사업기간  = c2.text_input("사업기간 *", key="rp_sp", placeholder="2025-01-01 ~ 2025-12-31")
            c3,c4=st.columns(2)
            지원기업수 = c3.text_input("지원기업수 *", key="rp_gi", placeholder="10")
            완료기업수 = c4.text_input("완료기업수 *", key="rp_ci", placeholder="8")
            c5,c6=st.columns(2)
            지원금총액 = c5.text_input("지원금총액 *", key="rp_ta", placeholder="50,000,000")
            집행금액  = c6.text_input("집행금액 *",   key="rp_ea", placeholder="42,000,000")
            주요성과  = st.text_area("주요 성과 *", height=80, key="rp_re")
            애로사항  = st.text_area("애로사항 *",  height=80, key="rp_is")
            향후계획  = st.text_area("향후 계획 *", height=80, key="rp_fp")
            go = st.button("📝 프롬프트 생성", key="rp_go")
        with col2:
            st.markdown("#### 📋 입력 데이터 요약")
            preview_data = {
                "사업명": 사업명, "사업기간": 사업기간,
                "지원기업수": f"{지원기업수}개소", "완료기업수": f"{완료기업수}개소",
                "지원금총액": f"{지원금총액}원", "집행금액": f"{집행금액}원",
            }
            rows = "".join(f"<tr><td class='lbl'>{k}</td><td>{_v(v,k)}</td></tr>" for k,v in preview_data.items())
            _preview(f"""
            <div class="preview-title" style="font-size:13pt">2025년 {_v(사업명,'사업명')} 결과보고서</div>
            <table class="pt"><tbody>{rows}</tbody></table>
            <p class="preview-section">■ 주요 성과</p>
            <p>{_v(주요성과,'주요 성과를 입력하세요.')}</p>
            <p class="preview-section">■ 향후 계획</p>
            <p>{_v(향후계획,'향후 계획을 입력하세요.')}</p>
            """)
        if go:
            if not all([사업명,사업기간,지원기업수,완료기업수,지원금총액,집행금액,주요성과,애로사항,향후계획]):
                st.error("⚠️ 필수 항목(*)을 모두 입력해 주세요.")
            else:
                st.session_state.report_data = dict(
                    사업명=사업명, 사업기간=사업기간, 지원기업수=지원기업수,
                    완료기업수=완료기업수, 지원금총액=지원금총액, 집행금액=집행금액,
                    주요성과=주요성과, 애로사항=애로사항, 향후계획=향후계획,
                )
                st.session_state.report_prompt = generate_prompt(st.session_state.report_data)
                st.session_state.report_step = 2
                st.rerun()

    elif step == 2:
        with col1:
            st.markdown("### 아래 프롬프트를 Claude.ai에 붙여넣으세요.")
            st.text_area("생성된 프롬프트", value=st.session_state.report_prompt, height=350, key="rp_prompt")
            st.link_button("🌐 Claude.ai 열기", "https://claude.ai")
            b1,b2=st.columns(2)
            if b1.button("⬅️ 이전"): st.session_state.report_step=1; st.rerun()
            if b2.button("다음 ➡️"): st.session_state.report_step=3; st.rerun()
        with col2:
            st.markdown("#### 📋 생성된 프롬프트 미리보기")
            _preview(f"<pre style='white-space:pre-wrap;font-size:9pt'>{st.session_state.report_prompt}</pre>")

    elif step == 3:
        with col1:
            st.markdown("### Claude.ai 응답을 붙여넣으세요.")
            response_text = st.text_area("Claude.ai 응답 텍스트", height=350, key="rp_resp")
            b1,b2=st.columns(2)
            if b1.button("⬅️ 이전"): st.session_state.report_step=2; st.rerun()
            if b2.button("📄 DOCX 생성"):
                if not response_text.strip():
                    st.error("⚠️ 응답 텍스트를 붙여넣어 주세요.")
                else:
                    with st.spinner("문서를 생성하는 중..."):
                        try:
                            path = make_report_docx(response_text, st.session_state.report_data)
                            st.session_state["rp_path"] = path
                            st.success(f"✅ 생성 완료: `{path.name}`")
                            st.session_state.report_step = 1
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ 오류: {e}")
        with col2:
            st.markdown("#### 📄 응답 미리보기")
            resp = st.session_state.get("rp_resp_preview","")
            content = response_text if "response_text" in dir() and response_text else '<span class="ph">Claude.ai 응답이 여기 표시됩니다.</span>'
            _preview(f"<div style='white-space:pre-wrap;font-size:9.5pt'>{content}</div>")
            if "rp_path" in st.session_state:
                st.divider()
                _dl_buttons(st.session_state["rp_path"])

# ──────────────────────────────────────────────
# 일괄 생성
# ──────────────────────────────────────────────

def page_batch():
    st.title("📦 일괄 생성")
    FORM_OPTIONS = {
        "공통 - 표준협약서":     ("공통","협약서"),
        "공통 - 개인정보 동의서": ("공통","동의서"),
        "공통 - 청렴이행각서":   ("공통","각서"),
        "창업 - 사업신청서":     ("창업","신청서"),
        "창업 - 집행정산보고":   ("창업","정산보고"),
    }
    COL_GUIDE = {
        "협약서":  "party_type, center_wide, center_local, rep_wide, rep_local, reg_wide, reg_local, support_content, period_start, period_end, contract_no, amount_central, amount_self, amount_fund, amount_local_support",
        "동의서":  "project_name, center_wide, signee_name, sign_date",
        "각서":    "center_wide, org_name, rep_name, sign_date",
        "신청서":  "center_local, reg_no, rep_name, rep_phone, address, email, manager_name, manager_phone, unit_name, unit_type, unit_period, ...",
        "정산보고": "report_type(1st/2nd/final), center_local, rep_name, address, phone, email, enterprise_name, ...",
    }
    selected = st.selectbox("서식 종류", list(FORM_OPTIONS.keys()))
    group, form_type = FORM_OPTIONS[selected]
    st.info(f"📌 CSV 필수 컬럼: `{COL_GUIDE[form_type]}`")
    uploaded = st.file_uploader("CSV 파일 업로드 (UTF-8)", type=["csv"])
    if uploaded:
        df = pd.read_csv(uploaded, encoding="utf-8-sig", dtype=str).fillna("")
        st.dataframe(df.head(10), use_container_width=True)
        if st.button("🚀 전체 생성"):
            with tempfile.NamedTemporaryFile(suffix=".csv", delete=False, mode="wb") as tmp:
                tmp.write(uploaded.getvalue()); tmp_path = tmp.name
            with st.spinner(f"{len(df)}개 문서 생성 중..."):
                try:
                    results = (make_all_common if group=="공통" else make_all_startup)(tmp_path, form_type)
                    st.success(f"✅ {len(results)}개 문서 생성 완료")
                    for p in results:
                        c1,c2=st.columns([5,1]); c1.text(p.name)
                        with open(p,"rb") as f:
                            c2.download_button("⬇️", f.read(), p.name, key=f"dl_{p.name}",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                except Exception as e:
                    st.error(f"❌ 오류: {e}")
            Path(tmp_path).unlink(missing_ok=True)

# ──────────────────────────────────────────────
# 라우팅
# ──────────────────────────────────────────────

{
    "home": page_home, "agreement": page_agreement,
    "privacy": page_privacy, "pledge": page_pledge,
    "application": page_application, "settlement": page_settlement,
    "report": page_report, "batch": page_batch,
}[page]()
