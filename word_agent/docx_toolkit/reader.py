"""Read-only helpers: document summaries and robust text lookup.

These are the only APIs the planner-facing layer uses to *see* a document.
"""

from __future__ import annotations

import difflib

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn

from word_agent.docx_toolkit.snapshot import snapshot_document


class LookupError_(Exception):
    """Raised when a paragraph or table cannot be found as described."""


def iter_body_blocks(doc: Document):
    """Yield ('p', Paragraph) / ('t', Table) items in document order."""

    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            yield "p", Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            yield "t", Table(child, doc)


def body_paragraphs(doc: Document) -> list[Paragraph]:
    """Body-level paragraphs in document order (the canonical numbering)."""

    return [
        Paragraph(child, doc)
        for child in doc.element.body.iterchildren()
        if child.tag == qn("w:p")
    ]


def body_paragraph_texts(doc: Document) -> list[str]:
    """Texts of :func:`body_paragraphs`."""

    return [p.text for p in body_paragraphs(doc)]


def find_paragraph_index(
    doc: Document, text: str, *, fuzzy: bool = False, cutoff: float = 0.9
) -> int:
    """Find the index of the paragraph whose full text equals ``text``.

    Returns the 0-based index among *body-level* paragraphs only (the same
    numbering as :func:`snapshot_document`). With ``fuzzy=True`` an exact match
    is tried first, then a high-similarity match as a fallback.
    """

    paragraphs = [
        Paragraph(child, doc)
        for child in doc.element.body.iterchildren()
        if child.tag == qn("w:p")
    ]
    target = text.strip()

    for i, p in enumerate(paragraphs):
        if p.text.strip() == target:
            return i

    if not fuzzy:
        return -1

    scores = [(i, difflib.SequenceMatcher(None, p.text.strip(), target).ratio()) for i, p in enumerate(paragraphs)]
    scores.sort(key=lambda x: x[1], reverse=True)
    if scores and scores[0][1] >= cutoff:
        return scores[0][0]
    return -1


def find_all_paragraph_indices(doc: Document, text: str) -> list[int]:
    """All body-level paragraph indices whose full text equals ``text``."""

    target = text.strip()
    out = []
    idx = 0
    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            if Paragraph(child, doc).text.strip() == target:
                out.append(idx)
            idx += 1
    return out


def get_table(doc: Document, index: int) -> Table:
    """Return the table at 0-based index ``index`` (document order)."""

    tables = [
        Table(child, doc)
        for child in doc.element.body.iterchildren()
        if child.tag == qn("w:tbl")
    ]
    if not (0 <= index < len(tables)):
        raise LookupError_(
            f"Table index {index} out of range (document has {len(tables)} table(s))."
        )
    return tables[index]


def document_summary(doc: Document, max_paragraphs: int = 60, max_chars: int = 400) -> str:
    """Build a compact text summary of the document for the planner's context.

    Paragraphs are numbered (matching snapshot/executor numbering) so the
    planner can refer to them. Tables are summarised by their cell grid.
    """

    snap = snapshot_document(doc).to_dict()
    lines: list[str] = []

    lines.append(f"PARAGRAPHS ({len(snap['paragraphs'])}):")
    for i, p in enumerate(snap["paragraphs"][:max_paragraphs]):
        text = p["text"]
        if len(text) > max_chars:
            text = text[: max_chars - 1] + "…"
        style = f" [style={p['style']}]" if p.get("style") and p["style"] != "Normal" else ""
        lines.append(f"{i}: {text!r}{style}")
    if len(snap["paragraphs"]) > max_paragraphs:
        lines.append(f"... ({len(snap['paragraphs']) - max_paragraphs} more paragraphs)")

    if snap["tables"]:
        lines.append("")
        lines.append(f"TABLES ({len(snap['tables'])}):")
        for t_i, t in enumerate(snap["tables"]):
            lines.append(f"table {t_i}: {t['rows']} rows x {t['cols']} cols")
            for r_i, row in enumerate(t["cellgrid"][:10]):
                lines.append(f"  row {r_i}: {row}")
            if t["rows"] > 10:
                lines.append(f"  ... ({t['rows'] - 10} more rows)")

    if snap["images"]:
        lines.append("")
        lines.append(f"IMAGES: {len(snap['images'])} embedded image(s) present.")

    return "\n".join(lines)
