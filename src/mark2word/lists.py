"""List numbering via Word OOXML."""

from __future__ import annotations

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

_DECIMAL_ABSTRACT_ATTR = "_mark2word_decimal_abstract_id"
_DECIMAL_LEVELS = 9


def _get_list_abstract_id(doc, style_name: str) -> int:
    style = doc.styles[style_name]
    num_id = style._element.pPr.numPr.numId.val
    numbering = doc.part.numbering_part.numbering_definitions._numbering
    ct_num = numbering.num_having_numId(num_id)
    return ct_num.abstractNumId.val


def _next_abstract_num_id(numbering) -> int:
    existing = [int(x) for x in numbering.xpath("./w:abstractNum/@w:abstractNumId")]
    for candidate in range(max(existing, default=-1) + 2):
        if candidate not in existing:
            return candidate
    return 0


def _build_decimal_level(ilvl: int) -> OxmlElement:
    lvl = OxmlElement("w:lvl")
    lvl.set(qn("w:ilvl"), str(ilvl))
    start = OxmlElement("w:start")
    start.set(qn("w:val"), "1")
    num_fmt = OxmlElement("w:numFmt")
    num_fmt.set(qn("w:val"), "decimal")
    lvl_text = OxmlElement("w:lvlText")
    lvl_text.set(qn("w:val"), f"%{ilvl + 1}.")
    lvl_jc = OxmlElement("w:lvlJc")
    lvl_jc.set(qn("w:val"), "left")
    lvl.append(start)
    lvl.append(num_fmt)
    lvl.append(lvl_text)
    lvl.append(lvl_jc)
    return lvl


def _build_decimal_multilevel_abstract(abstract_id: int) -> OxmlElement:
    """Multilevel abstract numbering with decimal format at every level."""
    abstract = OxmlElement("w:abstractNum")
    abstract.set(qn("w:abstractNumId"), str(abstract_id))
    multi = OxmlElement("w:multiLevelType")
    multi.set(qn("w:val"), "multilevel")
    abstract.append(multi)
    for ilvl in range(_DECIMAL_LEVELS):
        abstract.append(_build_decimal_level(ilvl))
    return abstract


def _ensure_decimal_multilevel_abstract(doc) -> int:
    cached = getattr(doc, _DECIMAL_ABSTRACT_ATTR, None)
    if cached is not None:
        return cached
    numbering = doc.part.numbering_part.numbering_definitions._numbering
    abstract_id = _next_abstract_num_id(numbering)
    abstract = _build_decimal_multilevel_abstract(abstract_id)
    numbering.insert(0, abstract)
    setattr(doc, _DECIMAL_ABSTRACT_ATTR, abstract_id)
    return abstract_id


def _new_list_num_id(doc, abstract_id: int, *, start: int | None = None, level: int = 0) -> int:
    numbering = doc.part.numbering_part.numbering_definitions._numbering
    ct_num = numbering.add_num(abstract_id)
    if start is not None and start != 1:
        override = ct_num.add_lvlOverride(level)
        override.add_startOverride(start)
    return ct_num.numId


def _same_list_run(prev: Paragraph | None, *, ordered: bool, run_num_id: int | None) -> bool:
    if prev is None or run_num_id is None:
        return False
    expected_style = "List Number" if ordered else "List Bullet"
    return prev.style.name == expected_style


def apply_list_numbering(
    paragraph: Paragraph,
    doc,
    *,
    ordered: bool,
    start: int | None,
    level: int,
    prev: Paragraph | None,
    run_num_id: int | None,
) -> int:
    """Apply list numbering and return the numId for the current list run."""
    if ordered:
        abstract_id = _ensure_decimal_multilevel_abstract(doc)
    else:
        abstract_id = _get_list_abstract_id(doc, "List Bullet")

    if _same_list_run(prev, ordered=ordered, run_num_id=run_num_id):
        num_id = run_num_id
    else:
        list_start = start if ordered else None
        num_id = _new_list_num_id(doc, abstract_id, start=list_start, level=level)

    p_pr = paragraph._p.get_or_add_pPr()
    num_pr = p_pr.get_or_add_numPr()
    num_pr.get_or_add_numId().val = num_id
    num_pr.get_or_add_ilvl().val = level
    return num_id
