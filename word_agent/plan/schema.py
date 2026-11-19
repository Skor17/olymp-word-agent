"""Pydantic models shared by the planner (LLM) and the executor (deterministic tools).

The LLM only ever produces an :class:`EditPlan`. Everything that touches the
document happens in :mod:`word_agent.docx_toolkit.mutate`, which understands
only these fixed, typed operations. This is what keeps the agent from being a
general-purpose code-execution harness.
"""

from __future__ import annotations

from typing import Annotated, Literal, Union

from pydantic import BaseModel, Field

OpType = Literal[
    "replace_text",
    "replace_paragraph",
    "insert_paragraph",
    "remove_paragraph",
    "add_table_row",
    "remove_table_row",
]

ALLOWED_OP_TYPES: tuple[str, ...] = (
    "replace_text",
    "replace_paragraph",
    "insert_paragraph",
    "remove_paragraph",
    "add_table_row",
    "remove_table_row",
)


class Op(BaseModel):
    """Base class for all editing operations."""

    type: str
    rationale: str = Field(
        default="",
        description="One-sentence explanation of why this operation is needed.",
    )


class ReplaceTextOp(Op):
    """Replace one occurrence of `old` inside a paragraph found by `paragraph`.

    The search is a substring search over the paragraph's concatenated run
    text. The formatting of the run the match starts in is preserved.
    """

    type: Literal["replace_text"] = "replace_text"
    paragraph: str = Field(
        description="Exact full text of the paragraph containing `old` (must be unique in the document)."
    )
    old: str = Field(description="Exact text to replace (substring of the paragraph).")
    new: str = Field(description="Replacement text.")


class ReplaceParagraphOp(Op):
    """Replace the full text of the paragraph found by `text`.

    The paragraph's style is preserved. If the paragraph's first run carries
    direct run formatting (bold/italic/...), that formatting is applied to the
    new text as well.
    """

    type: Literal["replace_paragraph"] = "replace_paragraph"
    text: str = Field(description="Exact full text of the paragraph to replace.")
    new_text: str = Field(description="The new full text of the paragraph.")


class InsertParagraphOp(Op):
    """Insert a new paragraph before or after the paragraph found by `anchor`.

    The new paragraph inherits the style of the anchor paragraph.
    """

    type: Literal["insert_paragraph"] = "insert_paragraph"
    anchor: str = Field(description="Exact full text of the paragraph next to which to insert.")
    text: str = Field(description="Text of the new paragraph.")
    position: Literal["after", "before"] = "after"


class RemoveParagraphOp(Op):
    """Remove the paragraph found by `text` (or every paragraph, if `text` is null)."""

    type: Literal["remove_paragraph"] = "remove_paragraph"
    text: str | None = Field(
        default=None,
        description=(
            "Exact full text of the paragraph to remove. Use null only when the instruction "
            "explicitly asks to remove *all* paragraphs."
        ),
    )


class AddTableRowOp(Op):
    """Append a new row to the table at 0-based index `table`."""

    type: Literal["add_table_row"] = "add_table_row"
    table: int = Field(description="0-based index of the table in document order.")
    cells: list[str] = Field(description="Cell texts, one per column.")


class RemoveTableRowOp(Op):
    """Remove the row at 0-based index `row` from the table at index `table`."""

    type: Literal["remove_table_row"] = "remove_table_row"
    table: int = Field(description="0-based index of the table in document order.")
    row: int = Field(description="0-based index of the row to remove within the table.")


OpModel = Annotated[
    Union[
        ReplaceTextOp,
        ReplaceParagraphOp,
        InsertParagraphOp,
        RemoveParagraphOp,
        AddTableRowOp,
        RemoveTableRowOp,
    ],
    Field(discriminator="type"),
]


class EditPlan(BaseModel):
    """The complete, deterministic edit program for one instruction."""

    operations: list[OpModel] = Field(default_factory=list)
    notes: str = Field(default="", description="Anything the executor should know (e.g. ambiguities).")


class Plan(BaseModel):
    """A plan plus its provenance (kept for debugging/audit logs)."""

    edit_plan: EditPlan
    source: str = Field(description="Which planner produced the plan (e.g. 'openai', 'mock').")
