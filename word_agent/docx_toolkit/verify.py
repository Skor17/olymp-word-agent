"""Integrity verification: prove that only the planned changes happened.

Approach
--------
Rather than diffing two documents and *guessing* which differences are
planned, we **simulate the plan** to reconstruct the exact expected result
(same dual reference frames as :class:`word_agent.docx_toolkit.mutate.Executor`)
and then require the actual post-edit document to match it:

1. **Paragraphs** — the simulation produces a *labeled* expected sequence:
   each item is either ``kept`` (an untouched original paragraph, by index),
   ``replaced`` (a paragraph the plan modified), or ``inserted`` (new text).
   The actual document must equal this sequence in text, order, and count;
   every ``kept`` paragraph must additionally be *formatting-identical* to
   the original (style + per-run formatting fingerprints). This catches
   unplanned text changes, formatting changes, reordering, insertions and
   deletions in one pass.
2. **Tables** — untouched tables must be formatting-identical; planned
   row additions/removals must match the expected cell grid exactly.
3. **Images** — every drawing (``a:blip``) present before must still be
   present after; no new drawings may appear unless the plan implies them
   (it never does — the op set has no image insertion).

The result is a :class:`VerificationResult` with a machine-checkable ``ok``
flag and a human-readable report. The graph refuses to save a document that
fails verification.
"""

from __future__ import annotations

from docx import Document

from word_agent.docx_toolkit.snapshot import snapshot_document
from word_agent.plan.schema import (
    AddTableRowOp,
    EditPlan,
    InsertParagraphOp,
    RemoveParagraphOp,
    RemoveTableRowOp,
    ReplaceParagraphOp,
    ReplaceTextOp,
)

# A labeled expected paragraph.
# kind: "kept" | "replaced" | "inserted"
# src_index: index into the BEFORE paragraph list (kept/replaced)
ExpectedPara = tuple[str, str, int | None]  # (kind, text, src_index)


def _resolve_frame(live: list[str], orig: list[str], quoted: str) -> int | None:
    """Mirror of the executor's paragraph reference resolution.

    Frame 1: unique exact (stripped) match in the live frame.
    Frame 2: unique exact match in the original frame -> live index by position.
    Returns the live index, or ``None`` if not resolvable.
    """

    target = quoted.strip()
    hits = [i for i, t in enumerate(live) if t.strip() == target]
    if len(hits) == 1:
        return hits[0]
    if not hits:
        ohits = [i for i, t in enumerate(orig) if t.strip() == target]
        if len(ohits) == 1 and ohits[0] < len(live):
            return ohits[0]
    return None


def _simulate_paragraphs(before_texts: list[str], plan: EditPlan) -> list[ExpectedPara]:
    """Reconstruct the labeled expected paragraph sequence after ``plan``.

    Faithfully mirrors :class:`word_agent.docx_toolkit.mutate.Executor`:
    identical dual reference frames and resolution order, so a plan that
    executes successfully is guaranteed to produce exactly this sequence.
    """

    live: list[ExpectedPara] = [
        ("kept", t, i) for i, t in enumerate(before_texts)
    ]
    orig_texts = list(before_texts)

    for op in plan.operations:
        if isinstance(op, ReplaceTextOp):
            i = _resolve_frame([x[1] for x in live], orig_texts, op.paragraph)
            if i is None:
                continue
            kind, text, src = live[i]
            if op.old in text:
                live[i] = ("replaced", text.replace(op.old, op.new, 1), src)
        elif isinstance(op, ReplaceParagraphOp):
            i = _resolve_frame([x[1] for x in live], orig_texts, op.text)
            if i is None:
                continue
            _, _, src = live[i]
            live[i] = ("replaced", op.new_text, src)
        elif isinstance(op, InsertParagraphOp):
            i = _resolve_frame([x[1] for x in live], orig_texts, op.anchor)
            if i is None:
                continue
            at = i + 1 if op.position == "after" else i
            live.insert(at, ("inserted", op.text, None))
            orig_texts.insert(at, op.text)
        elif isinstance(op, RemoveParagraphOp):
            if op.text is None:
                live, orig_texts = [], []
            else:
                i = _resolve_frame([x[1] for x in live], orig_texts, op.text)
                if i is None:
                    continue
                live.pop(i)
                if i < len(orig_texts):
                    orig_texts.pop(i)

    return live


def _para_fmt_key(p: dict) -> str:
    """Stable fingerprint of a paragraph's formatting (style + per-run formatting)."""

    parts = [p.get("style") or ""]
    for r in p.get("runs", []):
        fmt = r.get("fmt", {})
        parts.append(
            "|".join(
                str(fmt.get(k))
                for k in (
                    "bold",
                    "italic",
                    "underline",
                    "strike",
                    "name",
                    "color",
                    "size",
                    "eastAsia",
                    "vertAlign",
                )
            )
        )
    return "\u00a7".join(parts)


def _simulate_table_grids(before: list[dict], plan: EditPlan) -> list[list[list[str]]]:
    """Reconstruct the expected cell grid (texts) of every table after ``plan``."""

    grids = [[list(r) for r in t["cellgrid"]] for t in before]
    for op in plan.operations:
        if isinstance(op, AddTableRowOp):
            if 0 <= op.table < len(grids):
                ncols = len(grids[op.table][0]) if grids[op.table] else len(op.cells)
                if len(op.cells) == ncols:
                    grids[op.table].append(list(op.cells))
        elif isinstance(op, RemoveTableRowOp):
            if 0 <= op.table < len(grids) and 0 <= op.row < len(grids[op.table]):
                grids[op.table].pop(op.row)
    return grids


def _table_detail_key(t: dict) -> str:
    """Full formatting fingerprint of a table."""

    return repr(t["rows_detail"])


# --------------------------------------------------------------------------- #
# Result
# --------------------------------------------------------------------------- #
class VerificationResult:
    """Machine- and human-readable outcome of :func:`verify`."""

    def __init__(self, ok: bool, problems: list[str], notes: list[str]) -> None:
        self.ok = ok
        self.problems = problems
        self.notes = notes

    def report(self) -> str:
        status = "PASS" if self.ok else "FAIL"
        lines = [f"Integrity verification: {status}"]
        for n in self.notes:
            lines.append(f"  ok: {n}")
        for p in self.problems:
            lines.append(f"  PROBLEM: {p}")
        return "\n".join(lines)

    def __bool__(self) -> bool:
        return self.ok


# --------------------------------------------------------------------------- #
# Main verification
# --------------------------------------------------------------------------- #
def verify(before_path: str, after_path: str, plan: EditPlan) -> VerificationResult:
    """Verify that ``after_path`` equals ``before_path`` except for the plan's edits."""

    before = snapshot_document(Document(before_path))
    after = snapshot_document(Document(after_path))

    problems: list[str] = []
    notes: list[str] = []

    # -- 1. Images -------------------------------------------------------------
    if before.images != after.images:
        lost = [x for x in before.images if x not in after.images]
        gained = [x for x in after.images if x not in before.images]
        if lost:
            problems.append(f"images lost (drawings removed): {lost}")
        if gained:
            problems.append(f"unexpected images added: {gained}")
    else:
        notes.append(f"images intact ({len(before.images)} drawing(s))")

    # -- 2. Paragraphs ----------------------------------------------------------
    expected = _simulate_paragraphs([p["text"] for p in before.paragraphs], plan)
    actual = [p["text"] for p in after.paragraphs]

    if len(actual) != len(expected):
        problems.append(
            f"paragraph count mismatch: expected {len(expected)}, found {len(actual)}"
        )

    for j, (kind, exp_text, src_idx) in enumerate(expected):
        if j >= len(actual):
            break
        act_text = actual[j]
        if act_text != exp_text:
            problems.append(
                f"paragraph {j}: unexpected text {act_text[:60]!r} (expected {exp_text[:60]!r} per plan)"
            )
            continue
        if kind == "kept" and src_idx is not None:
            bp = before.paragraphs[src_idx]
            if _para_fmt_key(bp) != _para_fmt_key(after.paragraphs[j]):
                problems.append(
                    f"formatting changed in untouched paragraph {j}: {bp['text'][:60]!r}"
                )
        elif kind == "replaced":
            notes.append(f"paragraph {j} changed as planned: {exp_text[:50]!r}")
        elif kind == "inserted":
            notes.append(f"paragraph {j} inserted as planned: {exp_text[:50]!r}")

    # Extra paragraphs beyond the expected sequence (survivors of a partial
    # simulation, or injected content).
    for j in range(len(expected), len(actual)):
        problems.append(f"unexpected paragraph {j}: {actual[j][:60]!r}")

    # -- 3. Tables ----------------------------------------------------------------
    touched_tables: dict[int, None] = {}
    for op in plan.operations:
        if isinstance(op, (AddTableRowOp, RemoveTableRowOp)):
            touched_tables[op.table] = None

    expected_grids = _simulate_table_grids(before.tables, plan)
    if len(after.tables) != len(before.tables):
        problems.append(
            f"table count mismatch: expected {len(before.tables)}, found {len(after.tables)}"
        )

    for t_i in range(min(len(before.tables), len(after.tables))):
        bt, at = before.tables[t_i], after.tables[t_i]
        if t_i not in touched_tables:
            if _table_detail_key(bt) != _table_detail_key(at):
                problems.append(f"table {t_i} changed but the plan does not touch it")
        else:
            exp_grid = expected_grids[t_i]
            if at["cellgrid"] != exp_grid:
                problems.append(
                    f"table {t_i}: grid differs from plan — expected {exp_grid}, found {at['cellgrid']}"
                )
            else:
                notes.append(f"table {t_i} changed as planned ({at['rows']} rows)")

    ok = not problems
    return VerificationResult(ok=ok, problems=problems, notes=notes)


class VerificationError(Exception):
    """Raised when the post-edit document deviates from the plan."""


__all__ = [
    "VerificationError",
    "VerificationResult",
    "verify",
]
