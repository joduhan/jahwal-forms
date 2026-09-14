"""
app.py
제주광역자활센터 업무 자동화 시스템 — Streamlit 웹 UI
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
    make_agreement,
    make_privacy_consent,
    make_integrity_pledge,
    make_all_common,
)
from modules.startup_forms import (
    make_application,
    make_settlement_report,
    make_all_startup,
)
from modules.report_gen import generate_prompt, make_report_docx
from modules.doc_builder import OUTPUTS_DIR

DEFAULT_CENTER_WIDE = "제주특별자치도광역자활센터"

# ──────────────────────────────────────────────
# 페이지 설정
# ──────────────────────────────────────────────

st.set_page_config(
    page_title="제주광역자활센터 업무 자동화",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
.stButton > button { background-color: #1B4F8A; color: white; border-radius: 6px; }
.stButton > button:hover { background-color: #163f6f; color: white; }
.sec { font-size:1.05rem; font-weight:700; color:#1B4F8A;
       border-left:4px solid #1B4F8A; padding-left:8px; margin:1rem 0 0.4rem; }
</style>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────
# 공통 유틸
# ──────────────────────────────────────────────

def _dl_buttons(docx_path: Path):
    col1, col2 = st.columns(2)
    with col1:
        with open(docx_path, "rb") as f:
            st.download_button("⬇️ DOCX 다운로드", f.read(), docx_path.name,
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    pdf = docx_path.with_suffix(".pdf")
    with col2:
        if pdf.exists():
            with open(pdf, "rb") as f:
                st.download_button("⬇️ PDF 다운로드", f.read(), pdf.name, mime="application/pdf")
        else:
            st.caption("PDF: LibreOffice 설치 후 사용 가능")


def _ok(path: Path):
    st.success(f"✅ 생성 완료: `{path.name}`")
    _dl_buttons(path)


def _sec(label: str):
    st.markdown(f'<p class="sec">{label}</p>', unsafe_allow_html=True)


# ──────────────────────────────────────────────
# 사이드바
# ──────────────────────────────────────────────

MENU = {
    "🏠 홈": "home",
    "📄 표준협약서": "agreement",
    "📄 개인정보 동의서": "privacy",
    "📄 청렴이행각서": "pledge",
    "📋 사업신청서": "application",
    "📋 중간집행정산보고": "settlement",
    "📊 결과보고서": "report",
    "📦 일괄 생성": "batch",
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
# 홈
# ──────────────────────────────────────────────

def page_home():
    st.title("📋 제주광역자활센터 업무 자동화 시스템")
    st.markdown(
        "2025년도 중앙자활자금 개발원·광역 지원사업 서식을 자동으로 생성합니다.  \n"
        "왼쪽 메뉴에서 서식을 선택하세요."
    )
    _sec("생성 가능한 서식")
    st.dataframe(
        pd.DataFrame([
            ("공통3", "표준협약서",              "📄 표준협약서"),
            ("공통4", "개인정보 수집·이용 동의서", "📄 개인정보 동의서"),
            ("공통5", "청렴이행각서",             "📄 청렴이행각서"),
            ("창업1", "사업신청서",               "📋 사업신청서"),
            ("창업2", "집행정산보고",             "📋 중간집행정산보고"),
            ("보고",  "결과보고서",               "📊 결과보고서"),
        ], columns=["서식번호", "서식명", "메뉴"]),
        use_container_width=True, hide_index=True,
    )
    _sec("최근 생성 파일")
    OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)
    recent = sorted(OUTPUTS_DIR.glob("*.docx"), key=lambda p: p.stat().st_mtime, reverse=True)[:5]
    if recent:
        for f in recent:
            mtime = pd.Timestamp(f.stat().st_mtime, unit="s").strftime("%Y-%m-%d %H:%M")
            st.text(f"📄 {f.name}  ({mtime})")
    else:
        st.info("아직 생성된 파일이 없습니다.")

# ──────────────────────────────────────────────
# 표준협약서
# ──────────────────────────────────────────────

def page_agreement():
    st.title("📄 표준협약서")

    party_type_label = st.radio(
        "협약 유형", ["2자용 (광역 + 지역)", "3자용 (광역 + 지역 + 자활기업)"],
        horizontal=True,
    )
    is_3way = "3자용" in party_type_label
    party_type = "3way" if is_3way else "2way"

    with st.form("form_agreement"):
        _sec("협약 기본정보")
        c1, c2 = st.columns(2)
        center_wide  = c1.text_input("광역자활센터명 *", value=DEFAULT_CENTER_WIDE)
        contract_no  = c2.text_input("협약번호", placeholder="예: 제2025-001호")
        support_content = st.text_input("지원내용 *", placeholder="예: 자활기업 창업자금 지원")
        c3, c4 = st.columns(2)
        period_start = c3.text_input("협약기간 시작 *", placeholder="예: 2025.01.01.")
        period_end   = c4.text_input("협약기간 종료 *", placeholder="예: 2025.12.31.")

        _sec("광역자활센터 정보")
        c5, c6 = st.columns(2)
        rep_wide = c5.text_input("광역 대표자 *")
        reg_wide = c6.text_input("광역 사업자등록번호 *", placeholder="예: 123-45-67890")

        _sec("지역자활센터 정보")
        c7, c8 = st.columns(2)
        center_local = c7.text_input("지역자활센터명 *")
        c9, c10 = st.columns(2)
        rep_local = c9.text_input("지역 대표자 *")
        reg_local = c10.text_input("지역 사업자등록번호 *", placeholder="예: 234-56-78901")

        if is_3way:
            _sec("자활기업 정보")
            enterprise_name = st.text_input("자활기업명 *")
            c11, c12 = st.columns(2)
            rep_enterprise = c11.text_input("기업 대표자 *")
            reg_enterprise = c12.text_input("기업 사업자등록번호 *")
        else:
            enterprise_name = rep_enterprise = reg_enterprise = ""

        _sec("예산 현황 (원)")
        ca, cb, cc, cd = st.columns(4)
        amount_central       = ca.text_input("중앙자활자금",         placeholder="예: 5,000,000")
        amount_self          = cb.text_input("자부담",               placeholder="예: 1,000,000")
        amount_fund          = cc.text_input("자활기금",              placeholder="예: 0")
        amount_local_support = cd.text_input("지역자활사업지원비",    placeholder="예: 0")

        submitted = st.form_submit_button("📄 문서 생성")

    if submitted:
        required = [center_wide, support_content, period_start, period_end,
                    rep_wide, reg_wide, center_local, rep_local, reg_local]
        if is_3way:
            required += [enterprise_name, rep_enterprise, reg_enterprise]
        if not all(required):
            st.error("⚠️ 필수 항목(*)을 모두 입력해 주세요.")
            return
        with st.spinner("문서를 생성하는 중..."):
            try:
                path = make_agreement(party_type=party_type, data={
                    "center_wide": center_wide, "center_local": center_local,
                    "enterprise_name": enterprise_name,
                    "rep_wide": rep_wide, "rep_local": rep_local, "rep_enterprise": rep_enterprise,
                    "reg_wide": reg_wide, "reg_local": reg_local, "reg_enterprise": reg_enterprise,
                    "support_content": support_content,
                    "period_start": period_start, "period_end": period_end,
                    "amount_central": amount_central, "amount_self": amount_self,
                    "amount_fund": amount_fund, "amount_local_support": amount_local_support,
                    "contract_no": contract_no,
                })
                _ok(path)
            except Exception as e:
                st.error(f"❌ 오류: {e}")

# ──────────────────────────────────────────────
# 개인정보 동의서
# ──────────────────────────────────────────────

def page_privacy():
    st.title("📄 개인정보 수집·이용 동의서")
    with st.form("form_privacy"):
        c1, c2 = st.columns(2)
        project_name = c1.text_input("지원사업명 *", placeholder="예: 자활기업 창업 지원사업")
        center_wide  = c2.text_input("광역자활센터명 *", value=DEFAULT_CENTER_WIDE)
        c3, c4 = st.columns(2)
        signee_name  = c3.text_input("동의자 성명 *")
        sign_date    = c4.text_input("날짜(연도)", value=str(date.today().year),
                                     placeholder="예: 2025")
        submitted = st.form_submit_button("📄 문서 생성")

    if submitted:
        if not all([project_name, center_wide, signee_name]):
            st.error("⚠️ 필수 항목(*)을 모두 입력해 주세요.")
            return
        with st.spinner("문서를 생성하는 중..."):
            try:
                path = make_privacy_consent({
                    "project_name": project_name,
                    "center_wide": center_wide,
                    "signee_name": signee_name,
                    "sign_date": sign_date,
                })
                _ok(path)
            except Exception as e:
                st.error(f"❌ 오류: {e}")

# ──────────────────────────────────────────────
# 청렴이행각서
# ──────────────────────────────────────────────

def page_pledge():
    st.title("📄 청렴이행각서")
    with st.form("form_pledge"):
        center_wide = st.text_input("광역자활센터명 *", value=DEFAULT_CENTER_WIDE)
        c1, c2 = st.columns(2)
        org_name  = c1.text_input("서약기관명 *", placeholder="예: 제주시지역자활센터")
        rep_name  = c2.text_input("대표자명 *")
        sign_date = st.text_input("날짜(연도)", value=str(date.today().year))
        submitted = st.form_submit_button("📄 문서 생성")

    if submitted:
        if not all([center_wide, org_name, rep_name]):
            st.error("⚠️ 필수 항목(*)을 모두 입력해 주세요.")
            return
        with st.spinner("문서를 생성하는 중..."):
            try:
                path = make_integrity_pledge({
                    "center_wide": center_wide,
                    "org_name": org_name,
                    "rep_name": rep_name,
                    "sign_date": sign_date,
                })
                _ok(path)
            except Exception as e:
                st.error(f"❌ 오류: {e}")

# ──────────────────────────────────────────────
# 사업신청서
# ──────────────────────────────────────────────

def page_application():
    st.title("📋 자활기업 창업자금 사업신청서")

    with st.form("form_application"):
        # 표1: 신청기관 정보
        _sec("표1. 신청기관 정보")
        c1, c2 = st.columns(2)
        center_local = c1.text_input("지역자활센터명 *")
        reg_no       = c2.text_input("사업자등록번호 *", placeholder="예: 123-45-67890")
        c3, c4 = st.columns(2)
        rep_name  = c3.text_input("기관 대표자 *")
        rep_phone = c4.text_input("대표자 연락처 *", placeholder="예: 064-000-0000")
        address   = st.text_input("기관주소 *")
        c5, c6 = st.columns(2)
        email        = c5.text_input("이메일")
        manager_name = c6.text_input("사업담당자 *")
        manager_phone = st.text_input("담당자 휴대전화 *", placeholder="예: 010-0000-0000")

        # 표2: 사업단 현황
        _sec("표2. 사업단 현황")
        c7, c8 = st.columns(2)
        unit_name   = c7.text_input("사업단명 *")
        unit_type   = c8.text_input("사업유형 *", placeholder="예: 청소·환경")
        unit_period = st.text_input("운영기간 *", placeholder="예: 2022.01 ~ 2024.12")
        c9, c10 = st.columns(2)
        sales_avg    = c9.text_input("3년간 매출실적(연평균)", placeholder="예: 12,000,000")
        sales_recent = c10.text_input("최근 3개월 매출실적",  placeholder="예: 3,500,000")
        c11, c12 = st.columns(2)
        participants_avg       = c11.text_input("참여자수(월평균)", placeholder="예: 5명")
        startup_fund_saved     = c12.text_input("기적립 창업자금(100%)", placeholder="예: 1,000,000")
        c13, c14 = st.columns(2)
        startup_fund_approved  = c13.text_input("기승인 창업자금 금액", placeholder="없으면 0")
        approval_doc_no        = c14.text_input("지자체 승인 문서번호")
        approved_items         = st.text_input("기승인 창업자금 사용항목")

        # 표3: 자활기업 정보
        _sec("표3. 자활기업 정보")
        c15, c16 = st.columns(2)
        enterprise_name          = c15.text_input("자활기업명 *")
        enterprise_reg_no        = c16.text_input("기업 사업자번호")
        c17, c18 = st.columns(2)
        enterprise_rep           = c17.text_input("기업 대표자 *")
        enterprise_certified_date= c18.text_input("기업인정(예정)일", placeholder="예: 2025.01.01.")
        c19, c20 = st.columns(2)
        members_basic            = c19.text_input("기초생활보장 수급자", placeholder="예: 3명(60%)")
        members_near_poor        = c20.text_input("차상위자",            placeholder="예: 2명(40%)")
        enterprise_type          = st.selectbox(
            "기업형태 *", ["주식회사", "법인사업자", "협동조합", "개인기업"]
        )

        # 표4: 신청금액
        _sec("표4. 신청금액 (원)")
        c21, c22 = st.columns(2)
        amount_loan  = c21.text_input("융자(임대보증금)", placeholder="예: 2,000,000")
        amount_grant = c22.text_input("무상지원(운영자금)", placeholder="예: 3,000,000")
        repayment_type = st.radio(
            "융자상환계획", ["분기균등분할상환", "만기일시상환"], horizontal=True
        )
        repayment_reason = ""
        if repayment_type == "만기일시상환":
            repayment_reason = st.text_input("만기일시상환 사유 *")

        submitted = st.form_submit_button("📄 문서 생성")

    if submitted:
        required = [center_local, reg_no, rep_name, rep_phone, address,
                    manager_name, manager_phone, unit_name, unit_type, unit_period,
                    enterprise_name, enterprise_rep]
        if repayment_type == "만기일시상환":
            required.append(repayment_reason)
        if not all(required):
            st.error("⚠️ 필수 항목(*)을 모두 입력해 주세요.")
            return
        with st.spinner("문서를 생성하는 중..."):
            try:
                path = make_application({
                    "center_local": center_local, "reg_no": reg_no,
                    "rep_name": rep_name, "rep_phone": rep_phone,
                    "address": address, "email": email,
                    "manager_name": manager_name, "manager_phone": manager_phone,
                    "unit_name": unit_name, "unit_type": unit_type, "unit_period": unit_period,
                    "sales_avg": sales_avg, "sales_recent": sales_recent,
                    "participants_avg": participants_avg,
                    "startup_fund_saved": startup_fund_saved,
                    "startup_fund_approved": startup_fund_approved,
                    "approval_doc_no": approval_doc_no, "approved_items": approved_items,
                    "enterprise_name": enterprise_name, "enterprise_reg_no": enterprise_reg_no,
                    "enterprise_rep": enterprise_rep,
                    "enterprise_certified_date": enterprise_certified_date,
                    "members_basic": members_basic, "members_near_poor": members_near_poor,
                    "enterprise_type": enterprise_type,
                    "amount_loan": amount_loan, "amount_grant": amount_grant,
                    "repayment_type": repayment_type, "repayment_reason": repayment_reason,
                })
                _ok(path)
            except Exception as e:
                st.error(f"❌ 오류: {e}")

# ──────────────────────────────────────────────
# 중간집행정산보고
# ──────────────────────────────────────────────

def page_settlement():
    st.title("📋 집행정산보고")

    report_type_label = st.radio(
        "보고 유형",
        ["1차 중간 집행정산", "2차 중간 집행정산", "최종 집행정산"],
        horizontal=True,
    )
    RTYPE_MAP = {
        "1차 중간 집행정산": "1st",
        "2차 중간 집행정산": "2nd",
        "최종 집행정산":     "final",
    }
    report_type = RTYPE_MAP[report_type_label]

    with st.form("form_settlement"):
        _sec("기관 정보")
        c1, c2 = st.columns(2)
        center_local = c1.text_input("지역자활센터명 *")
        rep_name     = c2.text_input("대표자 *")
        address      = st.text_input("주소 *")
        c3, c4 = st.columns(2)
        phone = c3.text_input("전화번호 *", placeholder="예: 064-000-0000")
        email = c4.text_input("E-mail")

        _sec("자활기업 정보")
        c5, c6 = st.columns(2)
        enterprise_name  = c5.text_input("자활기업명 *")
        certified_date   = c6.text_input("기업인정일자", placeholder="예: 2025.01.01.")
        business_period  = st.text_input("사업기간", placeholder="예: 2025.01.01. ~ 2025.12.31.")
        contract_period  = st.text_input("협약기간", placeholder="예: 2025.01.01. ~ 2026.06.30.")
        report_period    = st.text_input("사업보고기간", placeholder="예: 2025.01.01. ~ 2025.06.30.")

        _sec("지원금액 및 집행현황 (원)")
        c7, c8, c9, c10 = st.columns(4)
        amount_loan  = c7.text_input("융자 지원금액",  placeholder="예: 2,000,000")
        amount_grant = c8.text_input("무상 지원금액",  placeholder="예: 3,000,000")
        exec_loan    = c9.text_input("융자 집행금액",  placeholder="예: 2,000,000")
        exec_grant   = c10.text_input("무상 집행금액", placeholder="예: 1,500,000")

        _sec("작성자 정보")
        c11, c12 = st.columns(2)
        manager_name  = c11.text_input("작성자(담당자) *")
        manager_phone = c12.text_input("작성자 연락처 *", placeholder="예: 010-0000-0000")
        sign_date = st.text_input("날짜(연도)", value=str(date.today().year))

        submitted = st.form_submit_button("📄 문서 생성")

    if submitted:
        if not all([center_local, rep_name, address, phone,
                    enterprise_name, manager_name, manager_phone]):
            st.error("⚠️ 필수 항목(*)을 모두 입력해 주세요.")
            return
        with st.spinner("문서를 생성하는 중..."):
            try:
                path = make_settlement_report(report_type, {
                    "center_local": center_local, "rep_name": rep_name,
                    "address": address, "phone": phone, "email": email,
                    "enterprise_name": enterprise_name, "certified_date": certified_date,
                    "business_period": business_period, "contract_period": contract_period,
                    "report_period": report_period,
                    "amount_loan": amount_loan, "amount_grant": amount_grant,
                    "exec_loan": exec_loan, "exec_grant": exec_grant,
                    "manager_name": manager_name, "manager_phone": manager_phone,
                    "sign_date": sign_date,
                })
                _ok(path)
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
    labels = ["1단계: 데이터 입력", "2단계: Claude.ai 연동", "3단계: 문서 완성"]
    st.progress((step - 1) / 2, text=labels[step - 1])

    # ── 1단계 ──
    if step == 1:
        with st.form("form_report"):
            c1, c2 = st.columns(2)
            사업명   = c1.text_input("사업명 *")
            사업기간  = c2.text_input("사업기간 *", placeholder="예: 2025-01-01 ~ 2025-12-31")
            c3, c4 = st.columns(2)
            지원기업수 = c3.text_input("지원기업수 *", placeholder="예: 10")
            완료기업수 = c4.text_input("완료기업수 *", placeholder="예: 8")
            c5, c6 = st.columns(2)
            지원금총액 = c5.text_input("지원금총액 *", placeholder="예: 50,000,000")
            집행금액  = c6.text_input("집행금액 *",   placeholder="예: 42,000,000")
            주요성과  = st.text_area("주요 성과 *", height=90)
            애로사항  = st.text_area("애로사항 *",  height=90)
            향후계획  = st.text_area("향후 계획 *", height=90)
            go = st.form_submit_button("📝 프롬프트 생성")

        if go:
            if not all([사업명, 사업기간, 지원기업수, 완료기업수,
                        지원금총액, 집행금액, 주요성과, 애로사항, 향후계획]):
                st.error("⚠️ 필수 항목(*)을 모두 입력해 주세요.")
                return
            st.session_state.report_data = {
                "사업명": 사업명, "사업기간": 사업기간,
                "지원기업수": 지원기업수, "완료기업수": 완료기업수,
                "지원금총액": 지원금총액, "집행금액": 집행금액,
                "주요성과": 주요성과, "애로사항": 애로사항, "향후계획": 향후계획,
            }
            st.session_state.report_prompt = generate_prompt(st.session_state.report_data)
            st.session_state.report_step = 2
            st.rerun()

    # ── 2단계 ──
    elif step == 2:
        st.markdown("### 아래 프롬프트를 복사해서 Claude.ai에 붙여넣으세요.")
        st.text_area("생성된 프롬프트", value=st.session_state.report_prompt, height=400)
        st.link_button("🌐 Claude.ai 열기", "https://claude.ai")
        st.info("💡 Claude.ai 응답을 받은 후 '다음 단계'를 눌러주세요.")
        col1, col2 = st.columns(2)
        if col1.button("⬅️ 이전"):
            st.session_state.report_step = 1
            st.rerun()
        if col2.button("다음 단계 ➡️"):
            st.session_state.report_step = 3
            st.rerun()

    # ── 3단계 ──
    elif step == 3:
        st.markdown("### Claude.ai 응답을 아래에 붙여넣으세요.")
        response_text = st.text_area("Claude.ai 응답 텍스트", height=350)
        col1, col2 = st.columns(2)
        if col1.button("⬅️ 이전"):
            st.session_state.report_step = 2
            st.rerun()
        if col2.button("📄 DOCX 생성"):
            if not response_text.strip():
                st.error("⚠️ Claude.ai 응답을 붙여넣어 주세요.")
                return
            with st.spinner("문서를 생성하는 중..."):
                try:
                    path = make_report_docx(response_text, st.session_state.report_data)
                    _ok(path)
                    st.session_state.report_step = 1
                except Exception as e:
                    st.error(f"❌ 오류: {e}")

# ──────────────────────────────────────────────
# 일괄 생성
# ──────────────────────────────────────────────

def page_batch():
    st.title("📦 일괄 생성")

    FORM_OPTIONS = {
        "공통 - 표준협약서":      ("공통", "협약서"),
        "공통 - 개인정보 동의서":  ("공통", "동의서"),
        "공통 - 청렴이행각서":    ("공통", "각서"),
        "창업 - 사업신청서":      ("창업", "신청서"),
        "창업 - 집행정산보고":    ("창업", "정산보고"),
    }
    COL_GUIDE = {
        "협약서":  "party_type, center_wide, center_local, rep_wide, rep_local, reg_wide, reg_local, support_content, period_start, period_end, contract_no, amount_central, amount_self, amount_fund, amount_local_support",
        "동의서":  "project_name, center_wide, signee_name, sign_date",
        "각서":    "center_wide, org_name, rep_name, sign_date",
        "신청서":  "center_local, reg_no, rep_name, rep_phone, address, email, manager_name, manager_phone, unit_name, ...",
        "정산보고": "report_type(1st/2nd/final), center_local, rep_name, address, phone, email, enterprise_name, ...",
    }

    selected = st.selectbox("서식 종류", list(FORM_OPTIONS.keys()))
    group, form_type = FORM_OPTIONS[selected]
    st.info(f"📌 CSV 필수 컬럼: `{COL_GUIDE[form_type]}`")

    uploaded = st.file_uploader("CSV 파일 업로드 (UTF-8)", type=["csv"])
    if uploaded:
        df = pd.read_csv(uploaded, encoding="utf-8-sig", dtype=str).fillna("")
        st.markdown("**미리보기 (최대 10행)**")
        st.dataframe(df.head(10), use_container_width=True)

        if st.button("🚀 전체 생성"):
            with tempfile.NamedTemporaryFile(suffix=".csv", delete=False, mode="wb") as tmp:
                tmp.write(uploaded.getvalue())
                tmp_path = tmp.name
            with st.spinner(f"{len(df)}개 문서 생성 중..."):
                try:
                    results = (make_all_common if group == "공통" else make_all_startup)(
                        tmp_path, form_type
                    )
                    st.success(f"✅ {len(results)}개 문서 생성 완료")
                    for p in results:
                        c1, c2 = st.columns([5, 1])
                        c1.text(p.name)
                        with open(p, "rb") as f:
                            c2.download_button("⬇️", f.read(), p.name, key=f"dl_{p.name}",
                                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                except Exception as e:
                    st.error(f"❌ 오류: {e}")
            Path(tmp_path).unlink(missing_ok=True)

# ──────────────────────────────────────────────
# 라우팅
# ──────────────────────────────────────────────

{
    "home":        page_home,
    "agreement":   page_agreement,
    "privacy":     page_privacy,
    "pledge":      page_pledge,
    "application": page_application,
    "settlement":  page_settlement,
    "report":      page_report,
    "batch":       page_batch,
}[page]()
