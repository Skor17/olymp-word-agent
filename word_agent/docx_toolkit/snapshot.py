"""Snapshot / fingerprinting of document content and formatting.

A snapshot is a deterministic, ordered digest of everything that matters for
integrity:

* every body paragraph with its text, style name, and the formatting of each
  run (bold/italic/underline/size/color/font/strike),
* every table with its cell grid and per-cell formatting,
* the set of embedded image parts (relationship ids) in the package.

The verifier (:mod:`word_agent.docx_toolkit.verify`) compares snapshots before
and after an edit and requires that the diff is exactly what the plan said it
would do.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any

from docx import Document
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn

_WHITESPACE_RE = re.compile(r"\s+")


def _norm(s: str | None) -> str:
    return _WHITESPACE_RE.sub(" ", (s or "")).strip()


# --------------------------------------------------------------------------- #
# Formatting digests
# --------------------------------------------------------------------------- #
def run_format_digest(run: Any) -> dict[str, Any]:
    """Return a stable dict of the *direct* formatting on a run.

    Only properties that are set on the run itself are recorded (inherited
    styles are resolved separately through the style name, so two runs with
    the same effective appearance but different style contexts still compare
    fairly — and, importantly, untouched runs compare *identically*).
    """

    rpr = run._element.rPr
    d: dict[str, Any] = {
        "bold": run.bold,
        "italic": run.italic,
        "underline": run.underline,
        "strike": run.font.strike,
        "name": run.font.name,
        "color": None,
        "size": None,
    }
    try:
        color = run.font.color
        if color and color.type is not None:
            d["color"] = str(color.rgb) if color.type == 1 else f"theme:{color.theme_color}"
    except Exception:  # pragma: no cover - defensive
        d["color"] = "unresolvable"
    size = run.font.size
    d["size"] = size.pt if size is not None else None

    if rpr is not None:
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is not None:
            d["eastAsia"] = rfonts.get(qn("w:eastAsia"))
        vert = rpr.find(qn("w:vertAlign"))
        d["vertAlign"] = vert.get(qn("w:val")) if vert is not None else None
    return d


def paragraph_format_digest(p: Paragraph) -> dict[str, Any]:
    """Digest for one paragraph: text, style, per-run text+formatting."""

    style = p.style.name if p.style is not None else None
    runs = [
        {"text": run.text, "fmt": run_format_digest(run)}
        for run in p.runs
        if run.text != ""
    ]
    return {
        "text": p.text,
        "style": style,
        "runs": runs,
    }


def _cell_digest(cell: _Cell) -> dict[str, Any]:
    paras = [paragraph_format_digest(p) for p in cell.paragraphs]
    # A cell always has at least one (possibly empty) paragraph.
    if not paras:
        paras = [{"text": "", "style": None, "runs": []}]
    return {
        "text": _norm(cell.text),
        "paragraphs": paras,
        "tables": [table_digest(t) for t in cell.tables],
    }


def table_digest(t: Table) -> dict[str, Any]:
    rows = [
        {"cells": [_cell_digest(c) for c in row.cells]}
        for row in t.rows
    ]
    return {
        "rows": len(t.rows),
        "cols": len(t.columns),
        "cellgrid": [[_norm(c.text) for c in row.cells] for row in t.rows],
        "rows_detail": rows,
    }


# --------------------------------------------------------------------------- #
# Document-level snapshot
# --------------------------------------------------------------------------- #
@dataclass
class Snapshot:
    """An ordered digest of a document's content and formatting."""

    paragraphs: list[dict[str, Any]] = field(default_factory=list)
    tables: list[dict[str, Any]] = field(default_factory=list)
    images: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "paragraphs": self.paragraphs,
            "tables": self.tables,
            "images": self.images,
        }


def _image_ids(doc: Document) -> list[str]:
    """Collect the ``r:embed`` ids of every image *rendered* in the document.

    We count ``a:blip`` references in the document body (which covers body
    paragraphs, tables, text boxes and headers/footers that share the body
    tree) rather than the package-level relationships: a removed drawing
    leaves its relationship behind, but the rendered reference disappears —
    which is exactly what integrity verification needs to detect.
    """

    ids: list[str] = []
    for blip in doc.element.body.iter(qn("a:blip")):
        embed = blip.get(qn("r:embed"))
        if embed and embed not in ids:
            ids.append(embed)
    return ids


def snapshot_document(doc: Document) -> Snapshot:
    """Build the full content+formatting snapshot of an open document."""

    snap = Snapshot(images=_image_ids(doc))
    body = doc.element.body

    # Walk the body in document order, handling paragraphs and tables as they
    # appear (this keeps tables interleaved with paragraphs in a stable order).
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            snap.paragraphs.append(paragraph_format_digest(Paragraph(child, doc)))
        elif child.tag == qn("w:tbl"):
            snap.tables.append(table_digest(Table(child, doc)))
        # sectPr etc. are structural and ignored for content integrity.

    return snap
