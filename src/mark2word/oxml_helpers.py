"""Centralized Word OXML helpers (python-docx public API + minimal private access)."""

from __future__ import annotations

from typing import Any

from docx.enum.text import WD_BREAK
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.shared import Pt

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def paragraph_element(paragraph) -> Any:
    return paragraph._p


def cell_element(cell) -> Any:
    return cell._tc


def add_run_text(parent, text: str, *, bold: bool = False, italic: bool = False) -> None:
    run = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")
    if bold:
        r_pr.append(OxmlElement("w:b"))
    if italic:
        r_pr.append(OxmlElement("w:i"))
    if len(r_pr):
        run.append(r_pr)
    text_elem = OxmlElement("w:t")
    text_elem.set(qn("xml:space"), "preserve")
    text_elem.text = text
    run.append(text_elem)
    parent.append(run)


def add_hyperlink(
    paragraph,
    text: str,
    url: str,
    style: dict[str, Any],
    *,
    bold: bool = False,
    italic: bool = False,
) -> None:
    """External (http) or internal (#anchor) hyperlink."""
    hyperlink = OxmlElement("w:hyperlink")
    if url.startswith("#"):
        hyperlink.set(qn("w:anchor"), url[1:])
    else:
        part = paragraph.part
        r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)
        hyperlink.set(qn("r:id"), r_id)

    run = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")
    if style.get("color"):
        color = OxmlElement("w:color")
        color.set(qn("w:val"), str(style["color"]).lstrip("#"))
        r_pr.append(color)
    if style.get("underline") is not False:
        underline = OxmlElement("w:u")
        underline.set(qn("w:val"), "single")
        r_pr.append(underline)
    if bold:
        r_pr.append(OxmlElement("w:b"))
    if italic:
        r_pr.append(OxmlElement("w:i"))
    if len(r_pr):
        run.append(r_pr)
    text_elem = OxmlElement("w:t")
    text_elem.set(qn("xml:space"), "preserve")
    text_elem.text = text
    run.append(text_elem)
    hyperlink.append(run)
    paragraph_element(paragraph).append(hyperlink)


def set_paragraph_border_bottom(paragraph, spec: dict[str, Any]) -> None:
    sz = max(1, round(float(str(spec.get("size", "0.5pt")).rstrip("pt")) * 8))
    color = str(spec.get("color", "000000")).lstrip("#")
    p_pr = paragraph_element(paragraph).get_or_add_pPr()
    p_pr.append(parse_xml(
        f"<w:pBdr {nsdecls('w')}>"
        f'<w:bottom w:val="single" w:sz="{sz}" w:space="2" w:color="{color}"/>'
        f"</w:pBdr>"
    ))


def add_bookmark(paragraph, bookmark_id: int, name: str) -> None:
    p = paragraph_element(paragraph)
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bookmark_id))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bookmark_id))
    p.insert(0, start)
    p.append(end)


def add_page_break(paragraph) -> None:
    paragraph.add_run().add_break(WD_BREAK.PAGE)


def add_field(paragraph, field_code: str) -> None:
    """Insert a Word field (e.g. PAGE, NUMPAGES)."""
    run = paragraph.add_run()
    r = run._r
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" {field_code} "
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    r.extend([fld_begin, instr, fld_sep, fld_end])


def set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell_element(cell).get_or_add_tcPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:val"), "clear")
    shading.set(qn("w:color"), "auto")
    shading.set(qn("w:fill"), str(fill).lstrip("#"))
    tc_pr.append(shading)


def set_cell_margins(cell, padding: dict[str, Any]) -> None:
    tc_pr = cell_element(cell).get_or_add_tcPr()
    tc_mar = OxmlElement("w:tcMar")
    for edge, key in (("top", "top"), ("bottom", "bottom"), ("left", "left"), ("right", "right")):
        if key not in padding:
            continue
        pt = _pt_twips(padding[key])
        node = OxmlElement(f"w:{edge}")
        node.set(qn("w:w"), str(pt))
        node.set(qn("w:type"), "dxa")
        tc_mar.append(node)
    if len(tc_mar):
        tc_pr.append(tc_mar)


def set_table_borders(table, spec: dict[str, Any]) -> None:
    sz = max(1, round(float(str(spec.get("size", "0.5pt")).rstrip("pt")) * 8))
    color = str(spec.get("color", "000000")).lstrip("#")
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    if tbl_pr is None:
        tbl_pr = OxmlElement("w:tblPr")
        tbl.insert(0, tbl_pr)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:color"), color)
        el.set(qn("w:space"), "0")
        borders.append(el)
    tbl_pr.append(borders)


def set_image_alt(run, alt: str) -> None:
    """Set accessibility description on an inline picture."""
    drawing = run._r.find(f".//{{{W_NS}}}drawing")
    if drawing is None:
        return
    doc_pr = drawing.find(f".//{{{W_NS}}}docPr")
    if doc_pr is not None:
        doc_pr.set("descr", alt)
        doc_pr.set("title", alt)


def _pt_twips(value) -> int:
    if isinstance(value, (int, float)):
        pt = float(value)
    else:
        s = str(value).strip().lower()
        pt = float(s[:-2]) if s.endswith("pt") else float(s)
    return int(pt * 20)
