"""
표준협약서 HWPX 직접 수정 방식
HWPX(ZIP) → Contents/section0.xml XML 교체 → 새 HWPX 저장
"""
from __future__ import annotations
import hashlib
import base64
import zipfile
from pathlib import Path
from lxml import etree

TEMPLATE_HWPX = Path(__file__).parent.parent / "표준 업무협약서.hwpx"
OUTPUTS_DIR = Path(__file__).parent.parent / "outputs"

HP = "http://www.hancom.co.kr/hwpml/2011/paragraph"


# ─── manifest 해시 갱신 ────────────────────────────────────────────────────────

def _update_manifest_hash(manifest_bytes: bytes, target_path: str, new_content: bytes) -> bytes | None:
    """manifest.xml에서 target_path 항목의 SHA-1 checksum과 size를 갱신."""
    if not manifest_bytes:
        return None
    try:
        root = etree.fromstring(manifest_bytes)
        ns = {"m": "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"}
        MNS = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"

        for entry in root.iter(f"{{{MNS}}}file-entry"):
            fp = entry.get(f"{{{MNS}}}full-path", "")
            if fp != target_path:
                continue
            # size 갱신
            if f"{{{MNS}}}size" in entry.attrib:
                entry.set(f"{{{MNS}}}size", str(len(new_content)))
            # checksum 갱신 (SHA1/1K: 처음 1024바이트의 SHA-1)
            enc = entry.find(f"{{{MNS}}}encryption-data")
            if enc is not None:
                ctype = enc.get(f"{{{MNS}}}checksum-type", "")
                if "SHA1" in ctype:
                    chunk = new_content[:1024]
                    digest = hashlib.sha1(chunk).digest()
                    enc.set(f"{{{MNS}}}checksum", base64.b64encode(digest).decode())
                elif "SHA256" in ctype:
                    chunk = new_content[:1024]
                    digest = hashlib.sha256(chunk).digest()
                    enc.set(f"{{{MNS}}}checksum", base64.b64encode(digest).decode())
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    except Exception:
        return None


# ─── XML 헬퍼 ─────────────────────────────────────────────────────────────────

def _remove_markpen(root) -> None:
    """문서 전체에서 hp:markpenBegin / hp:markpenEnd 제거 (형광펜 표시 제거)."""
    HP_TAGS = {f"{{{HP}}}markpenBegin", f"{{{HP}}}markpenEnd"}
    for parent in root.iter():
        to_remove = [ch for ch in parent if ch.tag in HP_TAGS]
        for ch in to_remove:
            # tail 텍스트를 앞 형제나 parent.text 에 붙여서 내용 유지
            tail = ch.tail or ""
            idx = list(parent).index(ch)
            if idx > 0:
                prev = list(parent)[idx - 1]
                prev.tail = (prev.tail or "") + tail
            else:
                parent.text = (parent.text or "") + tail
            parent.remove(ch)


def _t_full_text(t_el) -> str:
    """<hp:t> 노드의 전체 가시 텍스트 (text + 자식 tail)."""
    s = t_el.text or ""
    for child in t_el:
        s += child.tail or ""
    return s


def _para_text(para) -> str:
    return "".join(_t_full_text(t) for t in para.iter(f"{{{HP}}}t"))


def _set_t_simple(t_el, value: str) -> None:
    """<hp:t> 자식 없는 경우 text 교체."""
    t_el.text = value


def _set_after_markpen_end(run_el, value: str) -> None:
    """run 내 <hp:t>의 markpenEnd 이후 tail을 value로, 없으면 마지막 텍스트 교체."""
    for t in run_el.findall(f"{{{HP}}}t"):
        for child in t:
            if child.tag.split("}")[-1] == "markpenEnd":
                child.tail = value
                return
        # markpenEnd 없으면 단순 교체
        t.text = value


def _replace_in_tc(tc, old: str, new: str) -> None:
    """tc 안의 모든 <hp:t>에서 old → new 교체 (자식 요소 제거 후 단순 텍스트로)."""
    for t in tc.iter(f"{{{HP}}}t"):
        txt = _t_full_text(t)
        if old in txt:
            new_txt = txt.replace(old, new)
            for child in list(t):
                t.remove(child)
            t.text = new_txt


def _replace_all_t(root, old: str, new: str) -> None:
    """문서 전체 텍스트에서 old → new 교체.
    단일 hp:t뿐 아니라 같은 hp:run 안의 여러 hp:t 합산도 처리.
    """
    HP_T = f"{{{HP}}}t"
    HP_RUN = f"{{{HP}}}run"

    # 단일 hp:t 교체
    for t in root.iter(HP_T):
        txt = _t_full_text(t)
        if old in txt:
            new_txt = txt.replace(old, new)
            for child in list(t):
                t.remove(child)
            t.text = new_txt

    # run 안 여러 hp:t 합산 교체 (분리된 텍스트 처리)
    for run in root.iter(HP_RUN):
        ts = run.findall(HP_T)
        if len(ts) < 2:
            continue
        combined = "".join(_t_full_text(t) for t in ts)
        if old not in combined:
            continue
        new_combined = combined.replace(old, new)
        for child in list(ts[0]):
            ts[0].remove(child)
        ts[0].text = new_combined
        for t in ts[1:]:
            for child in list(t):
                t.remove(child)
            t.text = ""


# ─── 셀별 채우기 ──────────────────────────────────────────────────────────────

def _fill_name_cell(tc, name: str) -> None:
    """기관명 셀: 첫 번째 <hp:t>에 전체 이름 설정, 셀 내 나머지 모든 <hp:t> 비움."""
    paras = list(tc.iter(f"{{{HP}}}p"))
    if not paras:
        return

    first_set = False
    for para in paras:
        for run in para.findall(f"{{{HP}}}run"):
            for t in run.findall(f"{{{HP}}}t"):
                for child in list(t):
                    t.remove(child)
                if not first_set:
                    t.text = name
                    first_set = True
                else:
                    t.text = ""


def _fill_info_cell(tc, rep: str, reg_no: str, addr: str) -> None:
    """‣ 대표자 / ‣ 사업자등록번호 / ‣ 주  소 값 삽입."""
    for para in tc.iter(f"{{{HP}}}p"):
        ptxt = _para_text(para)
        runs = para.findall(f"{{{HP}}}run")
        if "대표자:" in ptxt:
            if len(runs) >= 2:
                # run2의 <hp:t> text = rep
                ts = runs[1].findall(f"{{{HP}}}t")
                if ts:
                    ts[0].text = rep
                else:
                    t_new = etree.SubElement(runs[1], f"{{{HP}}}t")
                    t_new.text = rep
            elif runs:
                _set_after_markpen_end(runs[0], rep)
        elif "사업자등록번호:" in ptxt:
            if runs:
                _set_after_markpen_end(runs[0], reg_no)
        elif "주  소:" in ptxt or "주소:" in ptxt:
            if runs:
                _set_after_markpen_end(runs[0], addr)


def _fill_agreement_no_cell(tc, agreement_no: str) -> None:
    """빈 협약번호 값 셀에 텍스트 삽입."""
    if not agreement_no:
        return
    paras = list(tc.iter(f"{{{HP}}}p"))
    if not paras:
        return
    runs = paras[0].findall(f"{{{HP}}}run")
    if not runs:
        return
    ts = runs[0].findall(f"{{{HP}}}t")
    if ts:
        ts[0].text = agreement_no
    else:
        t_new = etree.SubElement(runs[0], f"{{{HP}}}t")
        t_new.text = agreement_no


def _fill_budget_cell(tc, budget: str) -> None:
    """예산 표 금액 셀: '원' → '{budget}원'."""
    if not budget:
        return
    for t in tc.iter(f"{{{HP}}}t"):
        txt = _t_full_text(t).strip()
        if txt == "원" or txt == "":
            for child in list(t):
                t.remove(child)
            t.text = f"{budget}원"
            return


# ─── 표 채우기 ────────────────────────────────────────────────────────────────

def _fill_2way_table(tbl, data: dict) -> None:
    tcs = list(tbl.iter(f"{{{HP}}}tc"))
    support_name = data.get("support_name", "000 지원사업")
    gw = data.get("gwangnyeok", {})
    jy = data.get("jiyeok", {})
    period_str = (
        f"{data.get('period_start', '2025.00.00')}."
        f" ~ {data.get('period_end', '2025.00.00')}. (모니터링 기간 포함)"
    )
    date_str = data.get("date", "2025.  00.  00.")
    support_content = data.get("support_content", "사업비 지원")

    # tc[0]: 제목
    _replace_in_tc(tcs[0], "000 지원사업 협약서(2자용)", f"{support_name} 협약서(2자용)")
    # tc[2]: 협약번호 값
    _fill_agreement_no_cell(tcs[2], data.get("agreement_no", ""))
    # tc[4]: 광역 기관명
    _fill_name_cell(tcs[4], gw.get("name", "00광역자활센터"))
    # tc[5]: 광역 정보
    _fill_info_cell(tcs[5], gw.get("rep", ""), gw.get("reg_no", ""), gw.get("addr", ""))
    # tc[6]: 지역 기관명
    _fill_name_cell(tcs[6], jy.get("name", "00지역자활센터"))
    # tc[7]: 지역 정보
    _fill_info_cell(tcs[7], jy.get("rep", ""), jy.get("reg_no", ""), jy.get("addr", ""))
    # tc[10]: 지원내용
    _replace_in_tc(tcs[10], "사업비 지원", support_content)
    # tc[12]: 협약기간
    _replace_in_tc(tcs[12], "2025.00.00. ~ 2025.00.00. (모니터링 기간 포함)", period_str)
    # tc[13]: 날짜 + 서명 기관명
    _replace_in_tc(tcs[13], "2025.  00.  00.", date_str)
    _replace_in_tc(tcs[13], "00광역자활센터", gw.get("name", "00광역자활센터"))
    # tc[14]: 서명란 기관명
    _replace_in_tc(tcs[14], "00광역자활센터", gw.get("name", "00광역자활센터"))


def _fill_3way_table(tbl, data: dict) -> None:
    tcs = list(tbl.iter(f"{{{HP}}}tc"))
    support_name = data.get("support_name", "000 지원사업")
    gw = data.get("gwangnyeok", {})
    jy = data.get("jiyeok", {})
    gi = data.get("gieop", {})
    period_str = (
        f"{data.get('period_start', '2025.00.00')}."
        f" ~ {data.get('period_end', '2025.00.00')}. (모니터링 기간 포함)"
    )
    date_str = data.get("date", "2025.  00.  00.")
    support_content = data.get("support_content", "사업비 지원")

    # 제목
    _replace_in_tc(tcs[0], "000 지원사업 협약서(3자용)", f"{support_name} 협약서(3자용)")
    # 협약번호
    _fill_agreement_no_cell(tcs[2], data.get("agreement_no", ""))
    # 광역 기관명/정보
    _fill_name_cell(tcs[4], gw.get("name", "00광역자활센터"))
    _fill_info_cell(tcs[5], gw.get("rep", ""), gw.get("reg_no", ""), gw.get("addr", ""))
    # 지역 기관명/정보
    _fill_name_cell(tcs[6], jy.get("name", "00지역자활센터"))
    _fill_info_cell(tcs[7], jy.get("rep", ""), jy.get("reg_no", ""), jy.get("addr", ""))
    # 기업 기관명/정보 (tc[8], tc[9])
    if len(tcs) > 8:
        _fill_name_cell(tcs[8], gi.get("name", "00자활기업"))
    if len(tcs) > 9:
        _fill_info_cell(tcs[9], gi.get("rep", ""), gi.get("reg_no", ""), gi.get("addr", ""))
    # 지원내용 (tc[12] 추정)
    for tc in tcs[10:16]:
        _replace_in_tc(tc, "인건비, 자산취득, 시설개선 지원", support_content)
    # 협약기간
    for tc in tcs:
        _replace_in_tc(tc, "2025.00.00. ~ 2025.00.00. (모니터링 기간 포함)", period_str)
    # 날짜 + 서명
    for tc in tcs:
        _replace_in_tc(tc, "2025.  00.  00.", date_str)
    # 서명란 기관명
    for tc in tcs:
        _replace_in_tc(tc, "00광역자활센터", gw.get("name", "00광역자활센터"))
        if gi:
            _replace_in_tc(tc, "00자활기업", gi.get("name", "00자활기업"))


def _fill_budget_2way_table(tbl, budget: str) -> None:
    """예산 표 2자용 — 지원예산 행 금액 셀 교체."""
    tcs = list(tbl.iter(f"{{{HP}}}tc"))
    # tc[3]: '원' 셀
    if len(tcs) > 3:
        _fill_budget_cell(tcs[3], budget)


def _fill_budget_3way_table(tbl, budget: str, budget_gicup: str, budget_jiyeok: str) -> None:
    """예산 표 3자용 — 지원금액 행 각 셀 교체."""
    tcs = list(tbl.iter(f"{{{HP}}}tc"))
    # 구조: 구분 | 중앙자활자금 | 자부담 | 자활기금 | 지역자활사업지원비 | 지원금액
    # 지원금액 행의 각 셀에 값 삽입
    # 행 2(인덱스) 이후 금액 셀들 처리
    for tc in tcs:
        ptxt = "".join(_t_full_text(t) for t in tc.iter(f"{{{HP}}}t")).strip()
        if ptxt == "원" or ptxt == "":
            pass  # 빈 금액 셀은 아래서 처리
    # 간단하게: "원"이 있는 셀들 순서대로 budget, budget_gicup, budget_jiyeok
    budget_vals = [v for v in [budget, budget_gicup, budget_jiyeok] if v]
    idx = 0
    for tc in tcs:
        if idx >= len(budget_vals):
            break
        for t in tc.iter(f"{{{HP}}}t"):
            if _t_full_text(t).strip() == "원":
                for child in list(t):
                    t.remove(child)
                t.text = f"{budget_vals[idx]}원"
                idx += 1
                break


def _remove_table(root, tbl_el) -> None:
    """문서에서 표 요소 제거."""
    parent = tbl_el.getparent()
    if parent is not None:
        parent.remove(tbl_el)


def _fix_title_page_break(root, blank_count: int = 7) -> None:
    """
    '표준 업무협약서' run을 표 단락에서 분리하여 독립 단락으로 만들고
    앞에 빈 단락 blank_count개를 삽입하여 2페이지로 자연스럽게 이동.
    header.xml / content.hpf 수정 없음.
    """
    HP_P = f"{{{HP}}}p"
    HP_RUN = f"{{{HP}}}run"
    HP_T = f"{{{HP}}}t"
    target_text = "표준 업무협약서"

    for t_el in root.iter(HP_T):
        if (t_el.text or "").strip() == target_text:
            run_el = t_el.getparent()
            para_el = run_el.getparent()
            sec_el = para_el.getparent()
            if sec_el is None:
                break

            para_idx = list(sec_el).index(para_el)
            para_el.remove(run_el)

            # 빈 단락들 삽입 (표 단락 바로 뒤, 제목 앞)
            insert_pos = para_idx + 1
            for _ in range(blank_count):
                blank = etree.Element(HP_P)
                blank.set("id", "2147483648")
                blank.set("paraPrIDRef", "26")   # 본문 단락 스타일
                blank.set("styleIDRef", "0")
                blank.set("pageBreak", "0")
                blank.set("columnBreak", "0")
                blank.set("merged", "0")
                run_b = etree.SubElement(blank, HP_RUN)
                run_b.set("charPrIDRef", "0")
                etree.SubElement(run_b, HP_T)
                sec_el.insert(insert_pos, blank)
                insert_pos += 1

            # 제목 단락 (가운데 정렬, paraPrIDRef=25)
            title_p = etree.Element(HP_P)
            title_p.set("id", "2147483648")
            title_p.set("paraPrIDRef", "25")
            title_p.set("styleIDRef", "0")
            title_p.set("pageBreak", "0")
            title_p.set("columnBreak", "0")
            title_p.set("merged", "0")
            title_p.append(run_el)
            sec_el.insert(insert_pos, title_p)
            break


# ─── 메인 함수 ────────────────────────────────────────────────────────────────

def fill_hwpx(
    data: dict,
    output_path: str | Path | None = None,
    template_hwpx: str | Path | None = None,
) -> Path:
    """
    HWPX 템플릿에 값을 채워 새 HWPX 파일 저장.

    data 키: party_type, agreement_no, support_name,
             gwangnyeok, jiyeok, gieop (3way),
             support_content, period_start, period_end,
             budget, budget_gicup, budget_jiyeok, date
    """
    tpl = Path(template_hwpx) if template_hwpx else TEMPLATE_HWPX
    if not tpl.exists():
        raise FileNotFoundError(f"템플릿 없음: {tpl}")

    party_type = data.get("party_type", "2way")
    budget = data.get("budget", "")
    budget_gicup = data.get("budget_gicup", "")
    budget_jiyeok = data.get("budget_jiyeok", "")

    # ── XML 로드 ───────────────────────────────────────────────
    with zipfile.ZipFile(tpl) as z:
        xml_bytes = z.read("Contents/section0.xml")
        zip_infos = {info.filename: info for info in z.infolist()}
        all_files = {name: z.read(name) for name in z.namelist()}

    root = etree.fromstring(xml_bytes)

    # ── 전역 텍스트 교체 (표 채우기 전에 먼저 실행) ──────────────
    # 먼저 실행해야 _fill_budget_cell 이후 "000원" 이중 교체 버그 방지
    support_name = data.get("support_name", "000 지원사업")
    gw_name = data.get("gwangnyeok", {}).get("name", "00광역자활센터")
    jy_name = data.get("jiyeok", {}).get("name", "00지역자활센터")

    _replace_all_t(root, "000 지원사업", support_name)
    _replace_all_t(root, "00광역자활센터", gw_name)
    _replace_all_t(root, "00지역자활센터", jy_name)
    if budget:
        _replace_all_t(root, "000원", f"{budget}원")
    if party_type == "3way":
        gi_name = data.get("gieop", {}).get("name", "00자활기업")
        _replace_all_t(root, "00자활기업", gi_name)

    # ── 표 목록 파악 ───────────────────────────────────────────
    # 표 순서: 0=헤더, 1=2자용협약, 2=2자용서명, 3=3자용협약, 4=3자용서명, 5=예산2자용, 6=예산3자용
    tbls = list(root.iter(f"{{{HP}}}tbl"))

    if party_type == "2way":
        _fill_2way_table(tbls[1], data)
        _fill_budget_2way_table(tbls[5], budget)
        # 3자용 관련 표 제거 (역순)
        for ti in [6, 4, 3]:
            if ti < len(tbls):
                _remove_table(root, tbls[ti])
    else:
        _fill_3way_table(tbls[3], data)
        _fill_budget_3way_table(tbls[6], budget, budget_gicup, budget_jiyeok)
        # 2자용 관련 표 제거
        for ti in [5, 2, 1]:
            if ti < len(tbls):
                _remove_table(root, tbls[ti])

    # ── "표준 업무협약서" 2페이지 + 가운데 정렬 ──────────────────
    _fix_title_page_break(root)

    # ── 형광펜(마크펜) 제거 ────────────────────────────────────
    _remove_markpen(root)

    # ── XML 직렬화 ─────────────────────────────────────────────
    # 원본과 동일한 XML 선언 형식 유지 (큰따옴표 + ?> 앞 공백)
    body = etree.tostring(root, encoding="unicode")
    new_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>' + body).encode("UTF-8")

    # ── 출력 경로 ─────────────────────────────────────────────
    OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)
    if output_path is None:
        from datetime import date
        today = date.today().strftime("%Y%m%d")
        fname = f"공통3_표준협약서_{party_type}_{today}.hwpx"
        output_path = OUTPUTS_DIR / fname
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    # ── manifest.xml 해시 갱신 (한글 변조 감지 우회) ──────────────
    updated_manifest = _update_manifest_hash(
        all_files.get("META-INF/manifest.xml", b""),
        "Contents/section0.xml",
        new_xml,
    )

    # ── ZIP 재압축 ─────────────────────────────────────────────
    with zipfile.ZipFile(output_path, "w") as zout:
        for name, data_bytes in all_files.items():
            orig_info = zip_infos[name]
            if name == "mimetype":
                info = zipfile.ZipInfo(name)
                info.compress_type = zipfile.ZIP_STORED
                zout.writestr(info, data_bytes)
            elif name == "Contents/section0.xml":
                zout.writestr(name, new_xml, compress_type=zipfile.ZIP_DEFLATED)
            elif name == "META-INF/manifest.xml" and updated_manifest:
                zout.writestr(name, updated_manifest, compress_type=orig_info.compress_type)
            else:
                zout.writestr(
                    name, data_bytes,
                    compress_type=orig_info.compress_type,
                )

    return output_path


# ─── 테스트 ───────────────────────────────────────────────────────────────────

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

    out = fill_hwpx(
        test_data_2way,
        output_path=r"C:\Users\PC\Desktop\광역자활센터 자동화\outputs\공통3_표준협약서_2자용_test_v4.hwpx"
    )
    print(f"저장 완료: {out}")
    import subprocess
    subprocess.Popen(["start", str(out)], shell=True)
