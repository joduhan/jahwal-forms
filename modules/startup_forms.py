"""
startup_forms.py — HWP 원본 서식 기준 재작성
창업1: 자활기업 창업자금 사업신청서
창업2: 중간/최종 집행정산보고
"""

import json
from datetime import datetime
from pathlib import Path
from typing import List

import pandas as pd
from docx.shared import Pt, Cm, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from modules.doc_builder import create_doc, save_docx, docx_to_pdf

FONT_NAME = "맑은 고딕"


# ──────────────────────────────────────────────
# 내부 헬퍼 (common_forms와 동일 패턴)
# ──────────────────────────────────────────────

def _today() -> str:
    return datetime.now().strftime("%Y%m%d")


def _apply_font(run, name=FONT_NAME, size=10, bold=False):
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), name)
    run.font.size = Pt(size)
    run.font.bold = bold


def _para(doc, text="", size=10, bold=False, align="left", spacing=True, indent_cm=0.0):
    ALIGN = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
    }
    p = doc.add_paragraph()
    p.alignment = ALIGN.get(align, WD_ALIGN_PARAGRAPH.LEFT)
    if spacing:
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        p.paragraph_format.line_spacing = 1.6
    if indent_cm:
        p.paragraph_format.left_indent = Cm(indent_cm)
    if text:
        run = p.add_run(text)
        _apply_font(run, size=size, bold=bold)
    return p


def _set_tbl_borders(table, outer_pt=1.0, inner_pt=0.5):
    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    tblBorders = OxmlElement("w:tblBorders")
    for side, pt in [
        ("top", outer_pt), ("left", outer_pt),
        ("bottom", outer_pt), ("right", outer_pt),
        ("insideH", inner_pt), ("insideV", inner_pt),
    ]:
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(int(pt * 8)))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        tblBorders.append(el)
    tblPr.append(tblBorders)


def _cell(cell, text, size=10, bold=False, align="left", bg=None):
    cell.text = ""
    ALIGN = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
    }
    p = cell.paragraphs[0]
    p.alignment = ALIGN.get(align, WD_ALIGN_PARAGRAPH.LEFT)
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(str(text))
    _apply_font(run, size=size, bold=bold)
    if bg:
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), bg)
        tcPr.append(shd)


def _page_break(doc):
    p = doc.add_paragraph()
    run = p.add_run()
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    run._r.append(br)


def _fmt(value) -> str:
    try:
        return f"{int(str(value).replace(',', '')):,}"
    except (ValueError, TypeError):
        return str(value)


def _checkbox_row(options: list, selected: str) -> str:
    """선택된 항목에 ■, 나머지는 □로 표시한 문자열 반환."""
    parts = []
    for opt in options:
        mark = "■" if opt == selected else "□"
        parts.append(f"{mark} {opt}")
    return "  ".join(parts)


def _2col_table(doc, rows_data: list, col_widths=(Cm(4.5), Cm(12.0))):
    """(레이블, 값) 쌍 목록으로 2열 테이블 생성."""
    tbl = doc.add_table(rows=len(rows_data), cols=2)
    tbl.style = "Table Grid"
    _set_tbl_borders(tbl)
    for i, (label, val) in enumerate(rows_data):
        tbl.rows[i].cells[0].width = col_widths[0]
        tbl.rows[i].cells[1].width = col_widths[1]
        _cell(tbl.rows[i].cells[0], label, bold=True, align="center", bg="D9D9D9")
        _cell(tbl.rows[i].cells[1], val)
    return tbl


def _section_title(doc, text):
    _para(doc, text, size=11, bold=True, spacing=False)


# ──────────────────────────────────────────────
# 창업1: 사업신청서
# ──────────────────────────────────────────────

def make_application(data: dict = None) -> Path:
    """자활기업 창업자금 사업신청서(창업1)를 생성한다."""
    if data is None:
        data = {}
    doc = create_doc()

    # ── 제목 ──
    _para(doc, "자활기업 창업자금 사업신청서", size=16, bold=True, align="center", spacing=False)
    _para(doc, "(자활정보시스템 등록)", size=10, align="center", spacing=False)
    _para(doc)

    # ── 표1: 신청기관 정보 ──
    _section_title(doc, "■ 신청기관 정보")
    _2col_table(doc, [
        ("구분",         "지역자활센터"),
        ("기관명",       data.get("center_local", "")),
        ("사업자등록번호", data.get("reg_no", "")),
        ("기관 대표자",   data.get("rep_name", "")),
        ("대표자 연락처", data.get("rep_phone", "")),
        ("기관주소",      data.get("address", "")),
        ("이메일",        data.get("email", "")),
        ("사업담당자",    data.get("manager_name", "")),
        ("휴대전화",      data.get("manager_phone", "")),
    ])
    _para(doc)

    # ── 표2: 사업단 현황 ──
    _section_title(doc, "■ 사업단 현황")
    _2col_table(doc, [
        ("사업단명",              data.get("unit_name", "")),
        ("사업유형",              data.get("unit_type", "")),
        ("운영기간",              data.get("unit_period", "")),
        ("3년간 매출실적(연평균)", data.get("sales_avg", "")),
        ("최근 3개월 매출실적",   data.get("sales_recent", "")),
        ("참여자수(월평균)",       data.get("participants_avg", "")),
        ("기적립 창업자금(100%)",  data.get("startup_fund_saved", "")),
        ("기승인 창업자금 금액",   data.get("startup_fund_approved", "")),
        ("지자체 승인 문서번호",   data.get("approval_doc_no", "")),
        ("기승인 창업자금 사용항목", data.get("approved_items", "")),
    ])
    _para(doc)

    # ── 표3: 자활기업 정보 ──
    _section_title(doc, "■ 자활기업 정보")
    _para(doc, "※ 창업예정 자활기업의 경우 자활기업설립계획 정보에 대해 기입", size=9)

    # 기업형태 체크박스
    ent_type_selected = data.get("enterprise_type", "")
    ent_type_str = _checkbox_row(
        ["주식회사", "법인사업자", "협동조합", "개인기업"], ent_type_selected
    )

    # 구성원 현황
    mb = data.get("members_basic", "")
    mn = data.get("members_near_poor", "")
    members_str = f"기초생활보장 수급자 {mb}  /  차상위자 {mn}"

    _2col_table(doc, [
        ("자활기업명",       data.get("enterprise_name", "")),
        ("사업자번호",       data.get("enterprise_reg_no", "")),
        ("기업 대표자",      data.get("enterprise_rep", "")),
        ("기업인정(예정)일", data.get("enterprise_certified_date", "")),
        ("창업구성원 현황",  members_str),
        ("기업형태",         ent_type_str),
    ])
    _para(doc)

    # ── 표4: 신청금액 ──
    _section_title(doc, "■ 신청금액")
    loan = data.get("amount_loan", 0)
    grant = data.get("amount_grant", 0)
    try:
        total = int(str(loan).replace(",", "")) + int(str(grant).replace(",", ""))
        total_str = _fmt(total)
    except (ValueError, TypeError):
        total_str = ""

    repayment_type = data.get("repayment_type", "")
    repayment_reason = data.get("repayment_reason", "")
    repayment_str = _checkbox_row(["분기균등분할상환", "만기일시상환"], repayment_type)
    if repayment_reason:
        repayment_str += f"  (사유: {repayment_reason})"

    _2col_table(doc, [
        ("구분",           "금액"),
        ("융자(임대보증금)",  _fmt(loan)),
        ("무상지원(운영자금)", _fmt(grant)),
        ("합계",           total_str),
        ("융자상환계획",    repayment_str),
    ])
    _para(doc)

    # ── 다음 페이지: 첨부서류 ──
    _page_break(doc)
    _section_title(doc, "■ 첨부서류 목록")
    attachments = [
        "1. 사업계획서(첨부등록) [창업2]",
        "2. 사업예산서(시스템에 내역 등록)",
        "3. 사업자등록증(첨부등록) [창업3]",
        "4. 자활기업인정서(첨부등록)",
        "5. 구성원 명단(첨부등록)",
        "6. 기업형태 인증 서류",
        "7. 매출실적 자료",
        "8. 인센티브 증빙서류",
        "9. 개인정보 수집 및 이용 동의서 [창업13]",
        "10. 청렴이행각서 [공통5]",
        "11. 지자체 의견서 [창업14]",
    ]
    for item in attachments:
        _para(doc, item, size=10, indent_cm=0.5)

    filename = f"창업1_사업신청서_{data.get('center_local','기관')}_{_today()}.docx"
    path = save_docx(doc, filename)
    docx_to_pdf(path)
    return path


# ──────────────────────────────────────────────
# 창업2: 집행정산보고
# ──────────────────────────────────────────────

_REPORT_TITLES = {
    "1st":   "자활기업 창업자금 지원 기업 1차 중간 집행정산 보고",
    "2nd":   "자활기업 창업자금 지원 기업 2차 중간 집행정산 보고",
    "final": "자활기업 창업자금 지원 기업 최종 집행정산 보고",
}

_EXEC_PERIOD_LABEL = {
    "1st":   "협약일로부터 6개월 이내",
    "2nd":   "협약일로부터 12개월 이내",
    "final": "협약기간 18개월",
}

_CONTRACT_PERIOD_DEFAULT = {
    "1st":   "",
    "2nd":   "",
    "final": "18개월",
}

_MONTH_COUNT = {"1st": 6, "2nd": 12, "final": 18}


def make_settlement_report(report_type: str = "1st", data: dict = None) -> Path:
    """
    집행정산보고(창업2)를 생성한다.

    Args:
        report_type: '1st' / '2nd' / 'final'
        data: 보고서 데이터 dict
    """
    if data is None:
        data = {}
    doc = create_doc()

    title = _REPORT_TITLES.get(report_type, _REPORT_TITLES["1st"])

    # ── 표지: 제목 ──
    _para(doc, title, size=14, bold=True, align="center", spacing=False)
    _para(doc)

    # ── 표지 테이블 ──
    contract_period = (
        data.get("contract_period", "") or _CONTRACT_PERIOD_DEFAULT[report_type]
    )
    exec_period_label = _EXEC_PERIOD_LABEL[report_type]

    _2col_table(doc, [
        ("기  관  명", data.get("center_local", "")),
        ("대  표  자", data.get("rep_name", "")),
        ("주      소", data.get("address", "")),
        ("전화번호",   data.get("phone", "")),
        ("E-mail",    data.get("email", "")),
        ("자활기업명", data.get("enterprise_name", "")),
        ("기업인정일자", data.get("certified_date", "")),
        ("사업기간",   data.get("business_period", "")),
        ("협약기간",   contract_period),
        ("사업보고기간", data.get("report_period", "")),
        ("집행보고기간", exec_period_label),
    ])
    _para(doc)

    # ── 지원금액 테이블 ──
    _section_title(doc, "■ 지원금액 및 집행현황")
    amount_tbl = doc.add_table(rows=3, cols=3)
    amount_tbl.style = "Table Grid"
    _set_tbl_borders(amount_tbl)
    headers = ["구분", "지원금액", "집행금액"]
    for i, h in enumerate(headers):
        _cell(amount_tbl.rows[0].cells[i], h, bold=True, align="center", bg="D9D9D9")
    _cell(amount_tbl.rows[1].cells[0], "융자(임대보증금)", align="center")
    _cell(amount_tbl.rows[1].cells[1], _fmt(data.get("amount_loan", "")), align="center")
    _cell(amount_tbl.rows[1].cells[2], _fmt(data.get("exec_loan", "")), align="center")
    _cell(amount_tbl.rows[2].cells[0], "무상지원(운영자금)", align="center")
    _cell(amount_tbl.rows[2].cells[1], _fmt(data.get("amount_grant", "")), align="center")
    _cell(amount_tbl.rows[2].cells[2], _fmt(data.get("exec_grant", "")), align="center")

    col_w = [Cm(5.0), Cm(5.5), Cm(5.5)]
    for r in amount_tbl.rows:
        for i, w in enumerate(col_w):
            r.cells[i].width = w
    _para(doc)

    # ── 하단 항목 ──
    _para(doc, f"작성자(담당자): {data.get('manager_name','')}      연락처: {data.get('manager_phone','')}", size=10)
    _para(doc)

    _section_title(doc, "■ 첨부서류")
    for item in ["1. 사업결과보고서 [붙임1]", "2. 증빙서류", "3. 기타 관련 서류"]:
        _para(doc, item, size=10, indent_cm=0.5)

    _para(doc, "※ 증빙서류 범위:", size=9, bold=True)
    notes = [
        "① 지출결의서, 영수증, 세금계산서 등 지출 증빙서류",
        "② 계약서, 견적서 등 계약 관련 서류",
        "③ 임대차계약서(임대보증금 해당 시)",
        "④ 자동차 취득 관련 서류(해당 시)",
        "⑤ 기타 집행 관련 증빙서류 일체",
    ]
    for note in notes:
        _para(doc, note, size=9, indent_cm=1.0)
    _para(doc)

    sign_year = data.get("sign_date", "2025")
    _para(doc, f"{sign_year}년     월     일", size=10, align="center")
    _para(doc)
    _para(doc, f"수행기관명: {data.get('center_local','')}", size=10, align="right")
    _para(doc, "한국자활복지개발원장 귀중", size=10, align="right")

    # ── 붙임1: 사업결과보고서 ──
    _page_break(doc)
    _para(doc, "[붙임1] 사업결과보고서", size=14, bold=True, align="center", spacing=False)
    _para(doc)

    # 섹션1: 사업개요
    _section_title(doc, "1. 사업개요")
    _2col_table(doc, [
        ("사업목적", ""),
        ("사업내용", ""),
    ])
    _para(doc)

    # 섹션2: 사업 추진 내용
    _section_title(doc, "2. 사업 추진 내용")

    _para(doc, "가. 예산변경 현황", size=10, bold=True)
    change_tbl = doc.add_table(rows=2, cols=3)
    change_tbl.style = "Table Grid"
    _set_tbl_borders(change_tbl)
    for i, h in enumerate(["변경내용", "변경금액", "승인일자"]):
        _cell(change_tbl.rows[0].cells[i], h, bold=True, align="center", bg="D9D9D9")
        _cell(change_tbl.rows[1].cells[i], "", align="center")
    _para(doc)

    _para(doc, "나. 계획대비 추진결과", size=10, bold=True)
    progress_tbl = doc.add_table(rows=6, cols=3)
    progress_tbl.style = "Table Grid"
    _set_tbl_borders(progress_tbl)
    prog_headers = ["구분", "계획", "실적"]
    for i, h in enumerate(prog_headers):
        _cell(progress_tbl.rows[0].cells[i], h, bold=True, align="center", bg="D9D9D9")
    for i, label in enumerate(
        ["임대차계약", "자동차구입", "교육 및 컨설팅", "시설구축", "장비구입"], start=1
    ):
        _cell(progress_tbl.rows[i].cells[0], label, align="center")
        _cell(progress_tbl.rows[i].cells[1], "", align="center")
        _cell(progress_tbl.rows[i].cells[2], "", align="center")
    _para(doc)

    _para(doc, "다. 주요자산 관리내역", size=10, bold=True)
    asset_tbl = doc.add_table(rows=2, cols=4)
    asset_tbl.style = "Table Grid"
    _set_tbl_borders(asset_tbl)
    for i, h in enumerate(["품목", "취득일", "취득금액", "현재가치"]):
        _cell(asset_tbl.rows[0].cells[i], h, bold=True, align="center", bg="D9D9D9")
        _cell(asset_tbl.rows[1].cells[i], "", align="center")
    _para(doc)

    # 섹션3: 사업 추진 효과 및 중간평가
    _section_title(doc, "3. 사업 추진 효과 및 중간평가")

    _para(doc, "가. 매출현황 (단위: 천원)", size=10, bold=True)
    sales_tbl = doc.add_table(rows=3, cols=4)
    sales_tbl.style = "Table Grid"
    _set_tbl_borders(sales_tbl)
    for i, h in enumerate(["구분", "매출액", "매출원가", "순이익"]):
        _cell(sales_tbl.rows[0].cells[i], h, bold=True, align="center", bg="D9D9D9")
    _cell(sales_tbl.rows[1].cells[0], "당기", align="center")
    _cell(sales_tbl.rows[2].cells[0], "전기 대비", align="center")
    for r in [1, 2]:
        for c in [1, 2, 3]:
            _cell(sales_tbl.rows[r].cells[c], "", align="center")
    _para(doc)

    _para(doc, "나. 구성원 현황", size=10, bold=True)
    months = _MONTH_COUNT[report_type]
    member_cols = ["월"] + [f"{i}월" for i in range(1, months + 1)]
    mem_tbl = doc.add_table(rows=5, cols=len(member_cols))
    mem_tbl.style = "Table Grid"
    _set_tbl_borders(mem_tbl)
    for i, h in enumerate(member_cols):
        _cell(mem_tbl.rows[0].cells[i], h, bold=True, align="center", bg="D9D9D9")
    for r, label in enumerate(["총인원", "기초수급자", "차상위", "일반"], start=1):
        _cell(mem_tbl.rows[r].cells[0], label, align="center", bold=True)
        for c in range(1, len(member_cols)):
            _cell(mem_tbl.rows[r].cells[c], "", align="center")
    _para(doc)

    # 섹션4: 향후계획 및 기타 애로사항
    _section_title(doc, "4. 향후계획 및 기타 애로사항")
    _2col_table(doc, [
        ("향후계획", ""),
        ("기타\n애로사항", ""),
    ])

    filename = f"창업2_집행정산_{report_type}_{data.get('enterprise_name','기업')}_{_today()}.docx"
    path = save_docx(doc, filename)
    docx_to_pdf(path)
    return path


# ──────────────────────────────────────────────
# 일괄 생성
# ──────────────────────────────────────────────

def make_all_startup(csv_path: str, form_type: str) -> List[Path]:
    """
    CSV 일괄 생성.
    form_type: '신청서' / '정산보고'
    정산보고 CSV는 report_type 컬럼(1st/2nd/final) 필수.
    """
    csv_path = Path(csv_path)
    if not csv_path.exists():
        raise FileNotFoundError(f"CSV 파일 없음: {csv_path}")

    if form_type not in ("신청서", "정산보고"):
        raise ValueError(f"지원하지 않는 서식: '{form_type}'. 가능: ['신청서', '정산보고']")

    df = pd.read_csv(csv_path, encoding="utf-8-sig", dtype=str).fillna("")
    results: List[Path] = []

    for idx, row in df.iterrows():
        row_dict = row.to_dict()
        try:
            if form_type == "신청서":
                path = make_application(row_dict)
            else:
                rtype = row_dict.pop("report_type", "1st")
                if "items" in row_dict:
                    try:
                        row_dict["items"] = json.loads(row_dict["items"]) if row_dict["items"] else []
                    except json.JSONDecodeError:
                        row_dict["items"] = []
                path = make_settlement_report(rtype, row_dict)
            results.append(path)
            print(f"[완료] {path.name}")
        except Exception as e:
            print(f"[경고] {idx+1}행 실패: {e}")
    return results
