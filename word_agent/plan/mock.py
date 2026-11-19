"""Deterministic mock planner.

This is a *test double* for the LLM planner: same interface, no network,
fully reproducible. It interprets a small, explicit instruction grammar with
regular expressions and is intentionally transparent about what it
understands — anything else raises :class:`PlannerError` with the offending
clause, so the graph's bounded repair loop reports a clean, human-readable
failure instead of guessing.

Supported clause forms (clauses are separated by ``;``, `` and ``, or
newlines; each clause is one operation):

* ``replace "OLD" with "NEW"``                          -> ``replace_text``
* ``replace the paragraph "FULL" with "NEW FULL"``      -> ``replace_paragraph``
* ``insert "TEXT" after "ANCHOR"``  (or ``before``)     -> ``insert_paragraph``
* ``remove the paragraph "FULL"``                       -> ``remove_paragraph``
* ``remove all paragraphs``                             -> ``remove_paragraph`` (null)
* ``add a row to table N: "c1", "c2"``                  -> ``add_table_row``
* ``remove row R from table N``                         -> ``remove_table_row``

Quoted strings are resolved against the actual document: for operations that
need an exact *full* paragraph text, a quote that is a unique substring of one
paragraph is resolved to that paragraph's full text. Ambiguous or unknown
references are errors, never guesses.
"""

from __future__ import annotations

import re

from word_agent.docx_toolkit.reader import (
    find_all_paragraph_indices,
    body_paragraph_texts,
)
from word_agent.plan.base import Planner, PlannerContext, PlannerError
from word_agent.plan.schema import (
    AddTableRowOp,
    EditPlan,
    InsertParagraphOp,
    Plan,
    RemoveParagraphOp,
    RemoveTableRowOp,
    ReplaceParagraphOp,
    ReplaceTextOp,
)

_QUOTED = r'"([^"]*)"'
_NUM = r"(\d+)"

# Order matters: more specific grammars first.
_PATTERNS: list[tuple[re.Pattern, str]] = [
    (
        re.compile(
            rf"^\s*(?:replace|change|update)\s+the\s+paragraph\s+{_QUOTED}\s+"
            rf"(?:with|to|as)\s+{_QUOTED}\s*$",
            re.IGNORECASE,
        ),
        "replace_paragraph",
    ),
    (
        re.compile(
            rf"^\s*(?:replace|change|update)\s+{_QUOTED}\s+(?:with|to|as)\s+{_QUOTED}\s*$",
            re.IGNORECASE,
        ),
        "replace_text",
    ),
    (
        re.compile(
            rf"^\s*(?:insert|add)\s+{_QUOTED}\s+(before|after)\s+{_QUOTED}\s*$",
            re.IGNORECASE,
        ),
        "insert_paragraph",
    ),
    (
        re.compile(r"^\s*remove\s+all\s+paragraphs\s*$", re.IGNORECASE),
        "remove_all",
    ),
    (
        re.compile(rf"^\s*(?:remove|delete)\s+(?:the\s+paragraph\s+)?{_QUOTED}\s*$"),
        "remove_paragraph",
    ),
    (
        re.compile(rf"^\s*add\s+(?:a\s+)?row\s+to\s+table\s+{_NUM}\s*:\s*(.+)$"),
        "add_table_row",
    ),
    (
        re.compile(rf"^\s*remove\s+row\s+{_NUM}\s+from\s+table\s+{_NUM}\s*$"),
        "remove_table_row",
    ),
]

def _split_clauses(instruction: str) -> list[str]:
    """Split an instruction into clauses on commas / semicolons / 'and' / newlines.

    Splits happen only *outside* double quotes, so quoted values that contain
    commas (e.g. table cell lists) stay intact.
    """

    clauses: list[str] = []
    buf: list[str] = []
    in_quote = False
    i = 0
    n = len(instruction)

    def flush() -> None:
        clauses.append("".join(buf))
        buf.clear()

    while i < n:
        ch = instruction[i]
        if ch == '"':
            in_quote = not in_quote
            buf.append(ch)
            i += 1
            continue
        if in_quote:
            buf.append(ch)
            i += 1
            continue
        if ch in ",;\n":
            # A comma between a closing quote and an opening quote is part of
            # a quoted list (e.g. table cell values) — keep it in the clause.
            last = next((c for c in reversed(buf) if c != " "), "")
            nxt = next((c for c in instruction[i + 1 : i + 10] if c != " "), "")
            if ch == "," and last == '"' and nxt == '"':
                buf.append(ch)
                i += 1
                continue
            flush()
            i += 1
            continue
        # Whole-word "and" used as a conjunction (spaces on both sides, or
        # start/end of the instruction).
        if ch == " " and instruction[i + 1 : i + 4].lower() == "and":
            rest = i + 4
            if rest >= n or instruction[rest] == " ":
                flush()
                i = rest if rest >= n else rest + 1
                continue
        buf.append(ch)
        i += 1
    flush()
    return [c.strip() for c in clauses if c.strip()]


def _resolve_paragraph(context: PlannerContext, needle: str, op: str) -> str:
    """Resolve a (possibly partial) quoted string to a full paragraph text."""

    texts = body_paragraph_texts(context.doc)

    exact = find_all_paragraph_indices(context.doc, needle)
    if len(exact) == 1:
        return texts[exact[0]]
    if len(exact) > 1:
        raise PlannerError(f"{op}: {needle!r} matches {len(exact)} paragraphs — ambiguous")

    matches = [t for t in texts if needle in t]
    if len(matches) == 1:
        return matches[0]
    if not matches:
        raise PlannerError(f"{op}: no paragraph found containing {needle!r}")
    raise PlannerError(f"{op}: {needle!r} matches {len(matches)} paragraphs — ambiguous")


def _split_cells(raw: str) -> list[str]:
    cells = re.findall(_QUOTED, raw)
    if not cells:
        raise PlannerError(f"add_table_row: expected quoted cell values, got {raw!r}")
    return cells


class MockPlanner(Planner):
    """Rule-based, fully deterministic planner (offline test double)."""

    name = "mock"

    def plan(self, instruction: str, context: PlannerContext) -> Plan:
        ops = []
        for clause in _split_clauses(instruction):
            ops.append(self._parse_clause(clause, context))

        if not ops:
            raise PlannerError(f"could not derive any operation from: {instruction!r}")
        return Plan(
            edit_plan=EditPlan(operations=ops, notes="planned by MockPlanner"),
            source=self.name,
        )

    # ------------------------------------------------------------------ #
    def _parse_clause(self, clause: str, context: PlannerContext):
        for pattern, kind in _PATTERNS:
            m = pattern.match(clause)
            if not m:
                continue
            try:
                if kind == "replace_text":
                    old, new = m.group(1), m.group(2)
                    full = _resolve_paragraph(context, old, "replace_text")
                    return ReplaceTextOp(paragraph=full, old=old, new=new)
                if kind == "replace_paragraph":
                    text, new_text = m.group(1), m.group(2)
                    full = _resolve_paragraph(context, text, "replace_paragraph")
                    return ReplaceParagraphOp(text=full, new_text=new_text)
                if kind == "insert_paragraph":
                    text, position, anchor = m.group(1), m.group(2), m.group(3)
                    full = _resolve_paragraph(context, anchor, "insert_paragraph")
                    return InsertParagraphOp(anchor=full, text=text, position=position)
                if kind == "remove_all":
                    return RemoveParagraphOp(text=None)
                if kind == "remove_paragraph":
                    full = _resolve_paragraph(context, m.group(1), "remove_paragraph")
                    return RemoveParagraphOp(text=full)
                if kind == "add_table_row":
                    return AddTableRowOp(table=int(m.group(1)), cells=_split_cells(m.group(2)))
                if kind == "remove_table_row":
                    return RemoveTableRowOp(row=int(m.group(1)), table=int(m.group(2)))
            except (IndexError, ValueError) as e:
                raise PlannerError(f"malformed clause {clause!r}: {e}") from e

        raise PlannerError(f"unrecognised clause: {clause!r}")
