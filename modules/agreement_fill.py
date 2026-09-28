"""
표준협약서 템플릿 채우기
원본 HWPX → DOCX 변환 파일을 템플릿으로 사용하여 실제 값 삽입.
"""
from __future__ import annotations
import copy
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
from lxml import etree

TEMPLATE_DOCX = Path(__file__).parent.parent / "data" / "표준_업무협약서.docx"
OUTPUTS_DIR = Path(__file__).parent.parent / "outputs"


# ---------------------------------------------------------------------------
# 헬퍼
# ---------------------------------------------------------------------------

def _replace_run_text(para, search: str, replacement: str) -> bool:
    """단락에서 search 텍스트를 포함하는 run을 찾아 교체. True = 성공."""
    for run in para.runs:
        if search in run.text:
            run.text = run.text.replace(search, replacement)
            return True
    return False


def _set_runs_keep_bullet(para, label: str, value: str) -> None:
    """run0(‣ 기호)을 유지하고 라벨+값을 올바르게 설정."""
    runs = para.runs
    if len(runs) >= 3:
        # run0='‣', run1=라벨, run2=값
        runs[1].text = label
        runs[2].text = value
    elif len(runs) == 2:
        # run0='‣', run1=라벨만 → run1에 라벨+값 합치기
        runs[1].text = label + value
    elif len(runs) == 1:
        runs[0].text = "▸" + label + value


def _set_org_cell(cell, rep: str, reg_no: str, addr: str) -> None:
    """협약당사자 기관 정보 셀 — 라벨+값 형태로 삽입.
    결과: '‣ 대표자: {rep}', '‣ 사업자등록번호: {reg_no}', '‣ 주  소: {addr}'
    """
    for para in cell.paragraphs:
        txt = para.text
        if "대표자:" in txt:
            _set_runs_keep_bullet(para, " 대표자:", f" {rep}")
        elif "사업자등록번호:" in txt:
            _set_runs_keep_bullet(para, " 사업자등록번호:", f" {reg_no}")
        elif "주  소:" in txt or "주소:" in txt:
            _set_runs_keep_bullet(para, " 주  소:", f" {addr}")


def _set_cell_text_all_runs(cell, new_text: str) -> None:
    """셀의 첫 번째 단락 전체 텍스트를 교체 (단일 run으로)."""
    para = cell.paragraphs[0]
    for run in para.runs:
        run.text = ""
    if para.runs:
        para.runs[0].text = new_text
    else:
        para.add_run(new_text)


def _replace_text_in_paras(paragraphs, replacements: dict) -> None:
    """단락 목록에서 replacements 딕셔너리의 키를 값으로 교체."""
    for para in paragraphs:
        for search, repl in replacements.items():
            full_text = para.text
            if search in full_text:
                # 런 병합 없이 단순 교체: 전체 텍스트를 하나의 런으로 재구성
                if para.runs:
                    new_text = full_text.replace(search, repl)
                    for i, run in enumerate(para.runs):
                        run.text = new_text if i == 0 else ""


def _remove_table(doc, tbl_index: int) -> None:
    """문서에서 특정 인덱스의 표 제거."""
    tbl = doc.tables[tbl_index]
    tbl._element.getparent().remove(tbl._element)


# ---------------------------------------------------------------------------
# 메인 함수
# ---------------------------------------------------------------------------

def fill_agreement(
    data: dict,
    output_path: str | Path | None = None,
    template_docx: str | Path | None = None,
) -> Path:
    """
    협약서 템플릿에 실제 값을 채워 DOCX 생성.

    data 키:
        party_type      : '2way' | '3way'
        agreement_no    : 협약번호 (예: '2025-001')
        support_name    : 사업명 (예: '스타트업 육성 지원사업')
        gwangnyeok      : {'name': str, 'rep': str, 'reg_no': str, 'addr': str}
        jiyeok          : {'name': str, 'rep': str, 'reg_no': str, 'addr': str}
        gieop           : {'name': str, 'rep': str, 'reg_no': str, 'addr': str}  # 3way만
        support_content : 지원내용 (예: '사업비 지원')
        period_start    : 협약기간 시작 (예: '2025.03.01')
        period_end      : 협약기간 종료 (예: '2025.12.31')
        budget          : 중앙자활자금 예산 (예: '10,000,000')
        budget_gicup    : 자활기금 (3way, 예: '5,000,000')
        budget_jiyeok   : 지역자활사업지원비 (3way, 예: '2,000,000')
        date            : 협약 작성 날짜 (예: '2025.  3.  1.')
    """
    tpl = Path(template_docx) if template_docx else TEMPLATE_DOCX
    doc = Document(str(tpl))

    party_type = data.get("party_type", "2way")
    support_name = data.get("support_name", "000 지원사업")
    agreement_no = data.get("agreement_no", "")
    gw = data.get("gwangnyeok", {})
    jy = data.get("jiyeok", {})
    gi = data.get("gieop", {})
    support_content = data.get("support_content", "사업비 지원")
    period_start = data.get("period_start", "2025.00.00")
    period_end = data.get("period_end", "2025.00.00")
    budget = data.get("budget", "")
    budget_gicup = data.get("budget_gicup", "")
    budget_jiyeok = data.get("budget_jiyeok", "")
    date_str = data.get("date", "2025.  00.  00.")

    period_str = f"{period_start}. ~ {period_end}. (모니터링 기간 포함)"

    # -----------------------------------------------------------------------
    # 전역 텍스트 교체 규칙 (단락 수준)
    # -----------------------------------------------------------------------
    global_replacements = {
        "000 지원사업": support_name,
        "00광역자활센터": gw.get("name", "00광역자활센터"),
        "00지역자활센터": jy.get("name", "00지역자활센터"),
    }
    if party_type == "3way" and gi:
        global_replacements["00자활기업"] = gi.get("name", "00자활기업")
    # 수정 3: 본문 단락 "중앙자활자금 000원" → "중앙자활자금 {budget}원"
    if budget:
        global_replacements["중앙자활자금 000원"] = f"중앙자활자금 {budget}원"

    # -----------------------------------------------------------------------
    # 표 처리
    # 표 0: 헤딩 (변경 없음)
    # 표 1: 2자용 협약서 표  (index 1)
    # 표 2: 3자용 협약서 표  (index 2)
    # 표 3: 예산 표 2자용     (index 3)
    # 표 4: 예산 표 3자용     (index 4)
    # 표 5: 배경 이미지       (변경 없음)
    # -----------------------------------------------------------------------

    # 현재 표 목록 (제거 전)
    tables = doc.tables

    if party_type == "2way":
        # 2자용 표 채우기
        tbl = tables[1]
        _fill_2way_table(tbl, support_name, agreement_no, gw, jy,
                         support_content, period_str, date_str)
        # 예산 표 3 채우기
        _fill_budget_2way(tables[3], budget)
        # 3자용 표/예산 제거 (역순으로 인덱스 유지)
        _remove_table(doc, 4)  # 예산 3자용
        _remove_table(doc, 2)  # 협약 3자용
    else:
        # 3자용 표 채우기
        tbl = tables[2]
        _fill_3way_table(tbl, support_name, agreement_no, gw, jy, gi,
                         support_content, period_str, date_str)
        # 예산 표 4 채우기
        _fill_budget_3way(tables[4], budget, budget_gicup, budget_jiyeok)
        # 2자용 표/예산 제거
        _remove_table(doc, 3)  # 예산 2자용 (먼저 제거해야 인덱스 유지)
        _remove_table(doc, 1)  # 협약 2자용

    # -----------------------------------------------------------------------
    # 단락 전역 교체
    # -----------------------------------------------------------------------
    _replace_text_in_paras(doc.paragraphs, global_replacements)

    # -----------------------------------------------------------------------
    # 저장
    # -----------------------------------------------------------------------
    OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)
    if output_path is None:
        from datetime import date
        today = date.today().strftime("%Y%m%d")
        fname = f"공통3_표준협약서_템플릿방식_{today}.docx"
        output_path = OUTPUTS_DIR / fname
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    return output_path


# ---------------------------------------------------------------------------
# 표 채우기 헬퍼
# ---------------------------------------------------------------------------

def _fill_2way_table(tbl, support_name, agreement_no, gw, jy,
                     support_content, period_str, date_str):
    rows = tbl.rows

    # 행 0: 제목 "000 지원사업 협약서(2자용)" → "{support_name} 협약서(2자용)"
    title_cell = _unique_cell(rows[0], 0)
    if title_cell and title_cell.paragraphs[0].runs:
        title_cell.paragraphs[0].runs[0].text = f"{support_name} 협약서(2자용)"
        for run in title_cell.paragraphs[0].runs[1:]:
            run.text = ""

    # 행 1: 동일 제목 두 번째 행 비우기 (한 줄 표시를 위해)
    title_cell2 = _unique_cell(rows[1], 0)
    if title_cell2 and title_cell2.paragraphs:
        for run in title_cell2.paragraphs[0].runs:
            run.text = ""

    # 행 0, 열 3: 협약번호
    no_cell = _unique_cell(rows[0], 3)
    if no_cell and agreement_no:
        para = no_cell.paragraphs[0]
        if para.runs:
            para.runs[-1].text = f"협약번호 : {agreement_no}"

    # 행 2: 광역자활센터
    gw_name_cell = _unique_cell(rows[2], 1)
    if gw_name_cell and gw.get("name"):
        name = gw["name"]
        # "00광역\n자활센터" 구조 → 기관명 2줄로
        paras = gw_name_cell.paragraphs
        if len(paras) >= 2:
            paras[0].runs[0].text = name if "\n" not in name else name.split("\n")[0]
            paras[1].runs[0].text = "" if "\n" not in name else name.split("\n")[1]
        elif paras:
            paras[0].runs[0].text = name

    gw_info_cell = _unique_cell(rows[2], 2)
    if gw_info_cell:
        _set_org_cell(gw_info_cell,
                      gw.get("rep", ""), gw.get("reg_no", ""), gw.get("addr", ""))

    # 행 3: 지역자활센터
    jy_name_cell = _unique_cell(rows[3], 1)
    if jy_name_cell and jy.get("name"):
        name = jy["name"]
        paras = jy_name_cell.paragraphs
        if len(paras) >= 2:
            paras[0].runs[0].text = name if "\n" not in name else name.split("\n")[0]
            paras[1].runs[0].text = "" if "\n" not in name else name.split("\n")[1]
        elif paras:
            paras[0].runs[0].text = name

    jy_info_cell = _unique_cell(rows[3], 2)
    if jy_info_cell:
        _set_org_cell(jy_info_cell,
                      jy.get("rep", ""), jy.get("reg_no", ""), jy.get("addr", ""))

    # 행 4: 지원내용
    sc_cell = _unique_cell(rows[4], 2)
    if sc_cell and support_content:
        if sc_cell.paragraphs and sc_cell.paragraphs[0].runs:
            sc_cell.paragraphs[0].runs[0].text = support_content

    # 행 5: 협약기간
    period_cell = _unique_cell(rows[5], 2)
    if period_cell:
        if period_cell.paragraphs and period_cell.paragraphs[0].runs:
            period_cell.paragraphs[0].runs[0].text = period_str

    # 행 6: 날짜 "2025.  00.  00."
    body_cell = _unique_cell(rows[6], 0)
    if body_cell:
        for para in body_cell.paragraphs:
            if "2025." in para.text and "00" in para.text:
                if para.runs:
                    para.runs[0].text = date_str


def _fill_3way_table(tbl, support_name, agreement_no, gw, jy, gi,
                     support_content, period_str, date_str):
    rows = tbl.rows

    # 행 0: 제목 → "{support_name} 협약서(3자용)"
    title_cell = _unique_cell(rows[0], 0)
    if title_cell and title_cell.paragraphs[0].runs:
        title_cell.paragraphs[0].runs[0].text = f"{support_name} 협약서(3자용)"
        for run in title_cell.paragraphs[0].runs[1:]:
            run.text = ""

    # 행 1: 동일 제목 두 번째 행 비우기
    title_cell2 = _unique_cell(rows[1], 0)
    if title_cell2 and title_cell2.paragraphs:
        for run in title_cell2.paragraphs[0].runs:
            run.text = ""

    # 협약번호
    no_cell = _unique_cell(rows[0], 3)
    if no_cell and agreement_no:
        para = no_cell.paragraphs[0]
        if para.runs:
            para.runs[-1].text = f"협약번호 : {agreement_no}"

    # 행 2: 광역자활센터
    gw_name_cell = _unique_cell(rows[2], 1)
    if gw_name_cell and gw.get("name"):
        _set_org_name_cell(gw_name_cell, gw["name"])
    gw_info_cell = _unique_cell(rows[2], 2)
    if gw_info_cell:
        _set_org_cell(gw_info_cell,
                      gw.get("rep", ""), gw.get("reg_no", ""), gw.get("addr", ""))

    # 행 3: 지역자활센터
    jy_name_cell = _unique_cell(rows[3], 1)
    if jy_name_cell and jy.get("name"):
        _set_org_name_cell(jy_name_cell, jy["name"])
    jy_info_cell = _unique_cell(rows[3], 2)
    if jy_info_cell:
        _set_org_cell(jy_info_cell,
                      jy.get("rep", ""), jy.get("reg_no", ""), jy.get("addr", ""))

    # 행 4: 자활기업
    gi_name_cell = _unique_cell(rows[4], 1)
    if gi_name_cell and gi.get("name"):
        _set_org_name_cell(gi_name_cell, gi["name"])
    gi_info_cell = _unique_cell(rows[4], 2)
    if gi_info_cell:
        _set_org_cell(gi_info_cell,
                      gi.get("rep", ""), gi.get("reg_no", ""), gi.get("addr", ""))

    # 행 5: 지원내용
    sc_cell = _unique_cell(rows[5], 2)
    if sc_cell and support_content:
        if sc_cell.paragraphs and sc_cell.paragraphs[0].runs:
            sc_cell.paragraphs[0].runs[0].text = support_content

    # 행 6: 협약기간
    period_cell = _unique_cell(rows[6], 2)
    if period_cell:
        if period_cell.paragraphs and period_cell.paragraphs[0].runs:
            period_cell.paragraphs[0].runs[0].text = period_str

    # 행 7: 날짜
    body_cell = _unique_cell(rows[7], 0)
    if body_cell:
        for para in body_cell.paragraphs:
            if "2025." in para.text and "00" in para.text:
                if para.runs:
                    para.runs[0].text = date_str


def _fill_budget_2way(tbl, budget: str) -> None:
    """예산 표 (2자용) - 중앙자활자금 금액 설정."""
    if not budget:
        return
    rows = tbl.rows
    # 행 1, 열 1: "원" → "X,XXX,XXX원"
    cell = _unique_cell(rows[1], 1)
    if cell and cell.paragraphs and cell.paragraphs[0].runs:
        cell.paragraphs[0].runs[0].text = f"{budget}원"


def _fill_budget_3way(tbl, budget: str, budget_gicup: str, budget_jiyeok: str) -> None:
    """예산 표 (3자용) - 각 항목 금액 설정."""
    rows = tbl.rows
    # 행 2: 지원금액 행
    if budget:
        cell = _unique_cell(rows[2], 1)
        if cell and cell.paragraphs and cell.paragraphs[0].runs:
            cell.paragraphs[0].runs[0].text = f"{budget}원"
    if budget_gicup:
        cell = _unique_cell(rows[2], 2)
        if cell and cell.paragraphs and cell.paragraphs[0].runs:
            cell.paragraphs[0].runs[0].text = f"{budget_gicup}원"
    if budget_jiyeok:
        cell = _unique_cell(rows[2], 3)
        if cell and cell.paragraphs and cell.paragraphs[0].runs:
            cell.paragraphs[0].runs[0].text = f"{budget_jiyeok}원"


def _set_org_name_cell(cell, name: str) -> None:
    """기관명 셀 설정 (2줄 구조: '00광역' / '자활센터')."""
    paras = cell.paragraphs
    if len(paras) >= 2 and paras[0].runs and paras[1].runs:
        # 기관명을 그대로 첫 단락에, 두 번째는 비움
        paras[0].runs[0].text = name
        paras[1].runs[0].text = ""
    elif paras and paras[0].runs:
        paras[0].runs[0].text = name


def _unique_cell(row, col_idx: int):
    """행에서 고유한 셀 반환 (병합 중복 제거). col_idx 번째 유니크 셀."""
    seen = []
    for cell in row.cells:
        if not any(c._tc is cell._tc for c in seen):
            seen.append(cell)
    if col_idx < len(seen):
        return seen[col_idx]
    return None


# ---------------------------------------------------------------------------
# 테스트
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    test_data_2way = {
        "party_type": "2way",
        "agreement_no": "2025-공3-001",
        "support_name": "스타트업 육성 지원사업",
        "gwangnyeok": {
            "name": "부산광역자활센터",
            "rep": "홍길동",
            "reg_no": "123-45-67890",
            "addr": "부산광역시 연제구 중앙대로 123",
        },
        "jiyeok": {
            "name": "부산진구자활센터",
            "rep": "김철수",
            "reg_no": "234-56-78901",
            "addr": "부산광역시 부산진구 시민공원로 45",
        },
        "support_content": "사업비 지원",
        "period_start": "2025.03.01",
        "period_end": "2025.12.31",
        "budget": "10,000,000",
        "date": "2025.  3.  1.",
    }

    out = fill_agreement(
        test_data_2way,
        output_path=r"C:\Users\PC\Desktop\광역자활센터 자동화\outputs\공통3_표준협약서_템플릿방식_test.docx"
    )
    print(f"저장 완료: {out}")
    import subprocess
    subprocess.Popen(["start", str(out)], shell=True)
