"""Deterministic execution of an :class:`EditPlan` against an in-memory Document.

This module is the *only* place in the project that mutates a document. It
implements a small, fixed set of typed operations:

* ``replace_text``        – replace a substring inside one paragraph
* ``replace_paragraph``   – replace a paragraph's full text (style preserved)
* ``insert_paragraph``    – insert a new paragraph before/after an anchor
* ``remove_paragraph``    – remove one (or, explicitly, all) paragraphs
* ``add_table_row``       – append a row to a table
* ``remove_table_row``    – remove a row from a table

Reference resolution (the subtle part)
--------------------------------------
A plan is authored against the *original* document — that is what the planner
was shown. When several operations touch the same paragraph, the second
operation still quotes the original text even though the first operation has
already changed it. The :class:`Executor` therefore resolves each paragraph
reference against **two frames**:

1. the **current** live text of the document (so ops may also reference text
   produced by an earlier op), and
2. the **original** reference frame captured when execution started.

If a reference is ambiguous or found in neither frame, the operation fails
with a precise error — the executor never guesses.

Safety properties
-----------------
* The first failing operation aborts execution; a partially applied plan is
  never reported as success.
* Every operation keeps untouched XML (paragraph properties, sibling runs,
  other blocks) intact. Unchanged content is verified byte-level by
  :mod:`word_agent.docx_toolkit.verify`.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Callable

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.text.paragraph import Paragraph

from word_agent.docx_toolkit.reader import (
    LookupError_,
    body_paragraphs,
    body_paragraph_texts,
    get_table,
)
from word_agent.docx_toolkit.snapshot import run_format_digest
from word_agent.plan.schema import (
    AddTableRowOp,
    EditPlan,
    InsertParagraphOp,
    Op,
    RemoveParagraphOp,
    RemoveTableRowOp,
    ReplaceParagraphOp,
    ReplaceTextOp,
)

# --------------------------------------------------------------------------- #
# Data classes
# --------------------------------------------------------------------------- #
@dataclass
class OpResult:
    """Outcome of one planned operation."""

    index: int
    op_type: str
    status: str  # "applied" | "skipped" | "error"
    message: str


@dataclass
class ExecutionResult:
    """Aggregate outcome of applying a plan."""

    results: list[OpResult] = field(default_factory=list)

    @property
    def applied(self) -> int:
        return sum(1 for r in self.results if r.status == "applied")

    @property
    def skipped(self) -> int:
        return sum(1 for r in self.results if r.status == "skipped")

    @property
    def failed(self) -> int:
        return sum(1 for r in self.results if r.status == "error")

    @property
    def success(self) -> bool:
        return self.failed == 0

    def report(self) -> str:
        lines = []
        for r in self.results:
            lines.append(f"[{r.status.upper():>7}] op {r.index} ({r.op_type}): {r.message}")
        lines.append(
            f"Total: {self.applied} applied, {self.skipped} skipped, {self.failed} failed."
        )
        return "\n".join(lines)


# --------------------------------------------------------------------------- #
# Formatting re-application
# --------------------------------------------------------------------------- #
def _apply_run_format(run, fmt: dict) -> None:
    """Best-effort re-application of a run-format digest onto a new run."""

    run.bold = fmt.get("bold")
    run.italic = fmt.get("italic")
    underline = fmt.get("underline")
    if underline is not None:
        run.underline = underline
    strike = fmt.get("strike")
    if strike is not None:
        run.font.strike = strike
    if fmt.get("name"):
        run.font.name = fmt["name"]
    if fmt.get("size"):
        run.font.size = Pt(fmt["size"])
    color = fmt.get("color")
    if color and isinstance(color, str) and not color.startswith("theme:"):
        try:
            from docx.shared import RGBColor

            run.font.color.rgb = RGBColor.from_string(color)
        except Exception:  # pragma: no cover - defensive
            pass


# --------------------------------------------------------------------------- #
# Executor
# --------------------------------------------------------------------------- #
class Executor:
    """Applies an :class:`EditPlan` to one in-memory :class:`Document`."""

    def __init__(self, doc: Document) -> None:
        self.doc = doc
        # Original reference frame: the paragraph texts the plan was written
        # against. Kept in sync (insert/remove) so indices stay aligned with
        # the live document, but never rewritten by text replacements.
        self._orig_texts: list[str] = body_paragraph_texts(doc)

    # ------------------------------------------------------------------ #
    # Reference resolution
    # ------------------------------------------------------------------ #
    def _resolve_paragraph(self, text: str):
        """Resolve a quoted text to a live paragraph.

        Returns ``(paragraph, live_index)`` on success or ``(None, error)``.
        """

        target = text.strip()
        cur = body_paragraphs(self.doc)
        cur_stripped = [p.text.strip() for p in cur]
        orig_stripped = [t.strip() for t in self._orig_texts]

        # Frame 1: current live text.
        hits = [i for i, t in enumerate(cur_stripped) if t == target]
        if len(hits) == 1:
            return cur[hits[0]], hits[0]
        if len(hits) > 1:
            return None, f"ambiguous: {len(hits)} paragraphs currently equal {target!r}"

        # Frame 2: original reference frame (the plan was written against the
        # original document; earlier ops may have changed this paragraph).
        # Indexes stay aligned because insert/remove keep both frames in sync.
        ohits = [i for i, t in enumerate(orig_stripped) if t == target]
        if len(ohits) == 1 and ohits[0] < len(cur):
            return cur[ohits[0]], ohits[0]
        if len(ohits) > 1:
            return None, f"ambiguous: {target!r} matches {len(ohits)} original paragraphs"

        return None, f"paragraph not found: {text!r}"

    # ------------------------------------------------------------------ #
    # Operation handlers. Each returns (status, message).
    # ------------------------------------------------------------------ #
    def _op_replace_text(self, op: ReplaceTextOp) -> tuple[str, str]:
        p, err = self._resolve_paragraph(op.paragraph)
        if p is None:
            return "error", err

        full = p.text
        if op.old not in full:
            return (
                "error",
                f"substring {op.old!r} not found in paragraph: {full[:80]!r}...",
            )

        match_start = full.index(op.old)
        match_end = match_start + len(op.old)

        run_spans = []  # (run, start, end)
        pos = 0
        for run in p.runs:
            t = run.text
            run_spans.append((run, pos, pos + len(t)))
            pos += len(t)

        try:
            start_run, start_span, start_end = next(
                (r, s, e) for (r, s, e) in run_spans if s <= match_start < e
            )
        except StopIteration:
            return "error", "match starts in a non-run element (hyperlink/smartTag); not supported"

        # Rewrite the starting run: keep the part before the match, add the
        # new text, and — if the match ends inside this same run — keep the
        # part after the match.
        local = match_start - start_span
        if start_end >= match_end:
            start_run.text = (
                start_run.text[:local] + op.new + start_run.text[match_end - start_span :]
            )
        else:
            start_run.text = start_run.text[:local] + op.new

        # Consume the remainder of the match in runs that overlap the match
        # region (keeping the trailing part of each such run, if any).
        for run, s, e in run_spans:
            if run is start_run:
                continue
            if s < match_end and e > match_start:
                run.text = run.text[match_end - s :]

        return "applied", f"replaced {op.old!r} with {op.new!r}"

    def _op_replace_paragraph(self, op: ReplaceParagraphOp) -> tuple[str, str]:
        p, err = self._resolve_paragraph(op.text)
        if p is None:
            return "error", err

        # Capture direct formatting of the first non-empty run (if any) so the
        # replacement text keeps the paragraph's typographic face.
        fmt: dict | None = None
        first_run = next((r for r in p.runs if r.text != ""), None)
        if first_run is not None:
            fmt = run_format_digest(first_run)

        # Remove run containers, keeping paragraph properties (pPr) intact.
        for child in list(p._element):
            if child.tag in (qn("w:r"), qn("w:hyperlink"), qn("w:smartTag")):
                p._element.remove(child)

        new_run = p.add_run(op.new_text)
        if fmt:
            _apply_run_format(new_run, fmt)
        return "applied", f"replaced paragraph text with {op.new_text[:60]!r}..."

    def _op_insert_paragraph(self, op: InsertParagraphOp) -> tuple[str, str]:
        anchor, anchor_idx = self._resolve_paragraph(op.anchor)
        if anchor is None:
            return "error", f"anchor {anchor_idx}"

        if op.position == "before":
            new_p = anchor.insert_paragraph_before()
        else:
            el = OxmlElement("w:p")
            anchor._element.addnext(el)
            new_p = Paragraph(el, anchor._parent)

        # Inherit the anchor's style so the new paragraph fits its context.
        try:
            if anchor.style is not None and anchor.style.style_id:
                new_p.style = anchor.style
        except Exception:  # pragma: no cover - defensive
            pass

        new_p.add_run(op.text)

        # Keep the reference frame index-aligned with the live document and
        # make the new paragraph referenceable by later operations.
        at = anchor_idx + 1 if op.position == "after" else anchor_idx
        self._orig_texts.insert(at, op.text)
        where = "before" if op.position == "before" else "after"
        return "applied", f"inserted paragraph {where} anchor: {op.text[:60]!r}..."

    def _op_remove_paragraph(self, op: RemoveParagraphOp) -> tuple[str, str]:
        if op.text is None:
            paras = body_paragraphs(self.doc)
            if not paras:
                return "skipped", "no paragraphs to remove"
            for p in paras:
                p._element.getparent().remove(p._element)
            self._orig_texts = []
            return "applied", f"removed all {len(paras)} body paragraphs"

        p, idx = self._resolve_paragraph(op.text)
        if p is None:
            return "error", idx
        p._element.getparent().remove(p._element)

        # Keep the reference frame index-aligned with the live document.
        if idx < len(self._orig_texts):
            del self._orig_texts[idx]
        return "applied", f"removed paragraph: {op.text[:60]!r}..."

    def _op_add_table_row(self, op: AddTableRowOp) -> tuple[str, str]:
        tbl = get_table(self.doc, op.table)  # may raise LookupError_
        row = tbl.add_row()
        cells = row.cells
        if len(op.cells) != len(cells):
            return (
                "error",
                f"table {op.table} has {len(cells)} column(s) but {len(op.cells)} cell value(s) given",
            )
        for cell, text in zip(cells, op.cells):
            cell.text = text
        return "applied", f"appended row to table {op.table}: {op.cells}"

    def _op_remove_table_row(self, op: RemoveTableRowOp) -> tuple[str, str]:
        tbl = get_table(self.doc, op.table)  # may raise LookupError_
        if not (0 <= op.row < len(tbl.rows)):
            return (
                "error",
                f"row {op.row} out of range for table {op.table} ({len(tbl.rows)} rows)",
            )
        tr = tbl.rows[op.row]._tr
        tbl._tbl.remove(tr)
        return "applied", f"removed row {op.row} from table {op.table}"


_HANDLERS: dict[str, str] = {
    "replace_text": "_op_replace_text",
    "replace_paragraph": "_op_replace_paragraph",
    "insert_paragraph": "_op_insert_paragraph",
    "remove_paragraph": "_op_remove_paragraph",
    "add_table_row": "_op_add_table_row",
    "remove_table_row": "_op_remove_table_row",
}


# --------------------------------------------------------------------------- #
# Public API
# --------------------------------------------------------------------------- #
def execute_plan(doc: Document, plan: EditPlan) -> ExecutionResult:
    """Apply ``plan`` to the in-memory ``doc``.

    The document is *not* saved here — the caller decides when to persist.
    Execution stops at the first failing operation.
    """

    ex = Executor(doc)
    result = ExecutionResult()
    for i, op in enumerate(plan.operations):
        handler_name = _HANDLERS.get(op.type)
        if handler_name is None:  # pragma: no cover - schema forbids unknown types
            result.results.append(OpResult(i, op.type, "error", "unknown operation type"))
            continue
        handler: Callable[[Op], tuple[str, str]] = getattr(ex, handler_name)
        try:
            status, message = handler(op)
        except LookupError_ as e:
            status, message = "error", str(e)
        except Exception as e:  # defensive: a tool bug should not kill the graph
            status, message = "error", f"{type(e).__name__}: {e}"
        result.results.append(OpResult(i, op.type, status, message))
        if status == "error":
            break
    return result
