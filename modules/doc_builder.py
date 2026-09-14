"""
doc_builder.py
모든 서식 생성에서 공통으로 사용하는 DOCX 유틸리티 함수 모음
"""

import subprocess
import shutil
from pathlib import Path
from typing import List, Optional

from docx import Document
from docx.shared import Mm, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# 출력 폴더 경로
OUTPUTS_DIR = Path.home() / "Desktop" / "광역자활센터 자동화" / "outputs"

# 폰트 우선순위
PRIMARY_FONT = "나눔고딕"
FALLBACK_FONT = "맑은 고딕"


# ──────────────────────────────────────────────
# 내부 헬퍼
# ──────────────────────────────────────────────

def _set_cell_border(cell, border_size_pt: float = 0.5):
    """셀 4면에 실선 테두리를 XML로 직접 설정한다."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    # 1/8 pt 단위로 변환 (OOXML sz 단위)
    sz = str(int(border_size_pt * 8))

    for side in ("top", "left", "bottom", "right"):
        border_el = OxmlElement(f"w:{side}")
        border_el.set(qn("w:val"), "single")
        border_el.set(qn("w:sz"), sz)
        border_el.set(qn("w:space"), "0")
        border_el.set(qn("w:color"), "000000")
        tcPr.append(border_el)


def _set_cell_bg(cell, hex_color: str):
    """셀 배경색을 hex 문자열(예: 'D9D9D9')로 설정한다."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def _apply_font_xml(run, font_name: str):
    """
    한글 폰트를 XML rFonts 요소에 직접 설정한다.
    python-docx의 run.font.name은 ascii 폰트만 바꾸므로
    동아시아(hAnsi/eastAsia) 폰트도 함께 지정해야 한글에 적용된다.
    """
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), font_name)


def _align_const(align: str) -> WD_ALIGN_PARAGRAPH:
    mapping = {
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right":  WD_ALIGN_PARAGRAPH.RIGHT,
        "left":   WD_ALIGN_PARAGRAPH.LEFT,
    }
    return mapping.get(align.lower(), WD_ALIGN_PARAGRAPH.LEFT)


# ──────────────────────────────────────────────
# 공개 API
# ──────────────────────────────────────────────

def create_doc(top: float = 20, bottom: float = 20,
               left: float = 25, right: float = 25) -> Document:
    """
    A4 크기, 지정된 여백(mm)으로 설정된 Document 객체를 반환한다.

    Args:
        top:    위 여백 (mm), 기본 20
        bottom: 아래 여백 (mm), 기본 20
        left:   왼쪽 여백 (mm), 기본 25
        right:  오른쪽 여백 (mm), 기본 25
    """
    doc = Document()
    section = doc.sections[0]
    section.page_width  = Mm(210)   # A4 가로
    section.page_height = Mm(297)   # A4 세로
    section.top_margin    = Mm(top)
    section.bottom_margin = Mm(bottom)
    section.left_margin   = Mm(left)
    section.right_margin  = Mm(right)
    return doc


def set_font(run, name: str = PRIMARY_FONT, size: float = 10,
             bold: bool = False):
    """
    Run 객체에 폰트를 설정한다.
    나눔고딕이 지정된 경우 맑은 고딕을 fallback으로 함께 등록한다.

    Args:
        run:  docx Run 객체
        name: 폰트명 (기본 나눔고딕)
        size: 포인트 크기 (기본 10)
        bold: 굵게 여부
    """
    _apply_font_xml(run, name)
    # 나눔고딕 요청 시 hAnsi/cs 에만 fallback 추가 — eastAsia는 유지
    if name == PRIMARY_FONT:
        rPr = run._r.get_or_add_rPr()
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is not None:
            rFonts.set(qn("w:hAnsi"), FALLBACK_FONT)
    run.font.size = Pt(size)
    run.font.bold = bold


def add_title(doc: Document, text: str, size: float = 16,
              align: str = "center") -> None:
    """
    문서에 제목 문단을 추가한다.

    Args:
        doc:   Document 객체
        text:  제목 텍스트
        size:  글자 크기(pt), 기본 16
        align: 정렬 ('center' / 'left' / 'right'), 기본 'center'
    """
    para = doc.add_paragraph()
    para.alignment = _align_const(align)
    run = para.add_run(text)
    set_font(run, size=size, bold=True)


def add_paragraph(doc: Document, text: str, size: float = 10,
                  bold: bool = False, align: str = "left") -> None:
    """
    문서에 일반 문단을 추가한다.

    Args:
        doc:   Document 객체
        text:  문단 텍스트
        size:  글자 크기(pt), 기본 10
        bold:  굵게 여부
        align: 정렬 ('center' / 'left' / 'right'), 기본 'left'
    """
    para = doc.add_paragraph()
    para.alignment = _align_const(align)
    run = para.add_run(text)
    set_font(run, size=size, bold=bold)


def add_table(doc: Document, data: List[List[str]],
              col_widths: List[float], header: bool = True) -> None:
    """
    문서에 표를 추가한다.

    Args:
        doc:        Document 객체
        data:       List[List[str]] — 행/열 데이터 (첫 행이 헤더일 경우 header=True)
        col_widths: List[float] — 각 열 너비 (cm 단위)
        header:     True이면 첫 행에 배경색(#D9D9D9)과 볼드 적용
    """
    if not data:
        return

    rows = len(data)
    cols = len(data[0])
    table = doc.add_table(rows=rows, cols=cols)
    table.style = "Table Grid"

    for r_idx, row_data in enumerate(data):
        row = table.rows[r_idx]
        is_header_row = (header and r_idx == 0)

        for c_idx, cell_text in enumerate(row_data):
            cell = row.cells[c_idx]

            # 열 너비 설정
            if c_idx < len(col_widths):
                cell.width = Cm(col_widths[c_idx])

            # 테두리
            _set_cell_border(cell, border_size_pt=0.5)

            # 헤더 배경
            if is_header_row:
                _set_cell_bg(cell, "D9D9D9")

            # 텍스트 및 폰트
            cell.text = ""
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run(str(cell_text))
            set_font(run, size=9, bold=is_header_row)


def add_sign_section(doc: Document, left_label: str,
                     right_label: str) -> None:
    """
    서명란 문단을 추가한다.

    출력 형식:
        날짜: 2025년 __월 __일
        {left_label}: _____________ (인)
        {right_label}: _____________ (인)

    Args:
        doc:         Document 객체
        left_label:  왼쪽 서명자 레이블 (예: '갑(기관장)')
        right_label: 오른쪽 서명자 레이블 (예: '을(참여자)')
    """
    doc.add_paragraph()  # 여백
    lines = [
        "날짜: 2025년 __월 __일",
        f"{left_label}: _____________ (인)",
        f"{right_label}: _____________ (인)",
    ]
    for line in lines:
        add_paragraph(doc, line, size=10, align="right")


def save_docx(doc: Document, filename: str) -> Path:
    """
    Document를 outputs/ 폴더에 저장하고 전체 경로를 반환한다.
    outputs/ 폴더가 없으면 자동으로 생성한다.

    Args:
        doc:      저장할 Document 객체
        filename: 파일명 (확장자 .docx 포함 또는 미포함 모두 허용)

    Returns:
        저장된 파일의 전체 Path 객체
    """
    OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)

    if not filename.endswith(".docx"):
        filename += ".docx"

    save_path = OUTPUTS_DIR / filename
    doc.save(save_path)
    return save_path


def docx_to_pdf(docx_path: Path) -> Optional[Path]:
    """
    LibreOffice를 사용해 DOCX를 PDF로 변환한다.
    변환에 실패하면 예외를 발생시키지 않고 None을 반환한다.

    LibreOffice가 설치되어 있지 않거나 경로를 찾을 수 없는 경우도 None 반환.

    Args:
        docx_path: 변환할 .docx 파일 경로

    Returns:
        생성된 PDF 파일 경로, 실패 시 None
    """
    docx_path = Path(docx_path)
    pdf_path = docx_path.with_suffix(".pdf")

    # LibreOffice 실행 파일 후보 목록
    candidates = [
        "soffice",
        "libreoffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/usr/bin/libreoffice",
        "/usr/bin/soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]

    soffice = None
    for candidate in candidates:
        if shutil.which(candidate) or Path(candidate).exists():
            soffice = candidate
            break

    if soffice is None:
        return None

    try:
        result = subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf",
             "--outdir", str(docx_path.parent), str(docx_path)],
            capture_output=True,
            timeout=60,
        )
        if result.returncode == 0 and pdf_path.exists():
            return pdf_path
        return None
    except Exception:
        return None
