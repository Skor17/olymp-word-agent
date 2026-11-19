"""Unit tests for the deterministic executor (docx_toolkit.mutate)."""

from __future__ import annotations

import sys
from pathlib import Path

import pytest
from docx import Document

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from word_agent.docx_toolkit.mutate import Executor, execute_plan
from word_agent.docx_toolkit.reader import body_paragraphs, body_paragraph_texts
from word_agent.plan.schema import (
    AddTableRowOp,
    EditPlan,
    InsertParagraphOp,
    RemoveParagraphOp,
    RemoveTableRowOp,
    ReplaceParagraphOp,
    ReplaceTextOp,
)


def _doc_with_text(tmp_path: Path, text: str = "Hello world from the test.") -> Path:
    doc = Document()
    doc.add_paragraph(text)
    p = tmp_path / "t.docx"
    doc.save(str(p))
    return p


class TestReplaceText:
    def test_single_run_replacement_preserves_siblings(self, tmp_path: Path):
        doc = Document()
        p1 = doc.add_paragraph("Hello world from the test.")
        # A multi-run paragraph for the actual replacement:
        p2 = doc.add_paragraph()
        p2.add_run("Pre: ")
        p2.add_run("OLD")
        p2.add_run(" post")

        plan = EditPlan(operations=[ReplaceTextOp(paragraph=p2.text, old="OLD", new="NEW")])
        result = execute_plan(doc, plan)
        assert result.success, result.report()
        assert p2.text == "Pre: NEW post"
        # Untouched paragraph is untouched.
        assert body_paragraphs(doc)[0].text == "Hello world from the test."

    def test_replacement_spanning_multiple_runs(self, tmp_path: Path):
        doc = Document()
        p = doc.add_paragraph()
        p.add_run("The ")
        p.add_run("quick brown")
        p.add_run(" fox")
        plan = EditPlan(operations=[ReplaceTextOp(paragraph=p.text, old="quick brown fox", new="lazy dog")])
        result = execute_plan(doc, plan)
        assert result.success, result.report()
        assert p.text == "The lazy dog"

    def test_two_replacements_same_paragraph_dual_frame(self, tmp_path: Path):
        """The plan is written against the ORIGINAL text; the second op must
        still resolve after the first op changed the paragraph."""

        doc = Document()
        p = doc.add_paragraph()
        p.add_run("Client: Acme")
        p.add_run(" | Status: Draft")
        original = p.text

        plan = EditPlan(
            operations=[
                ReplaceTextOp(paragraph=original, old="Acme", new="Globex"),
                ReplaceTextOp(paragraph=original, old="Draft", new="Final"),
            ]
        )
        result = execute_plan(doc, plan)
        assert result.success, result.report()
        assert p.text == "Client: Globex | Status: Final"

    def test_missing_substring_fails(self, tmp_path: Path):
        doc = Document()
        p = doc.add_paragraph("some text here")
        plan = EditPlan(operations=[ReplaceTextOp(paragraph=p.text, old="nope", new="x")])
        result = execute_plan(doc, plan)
        assert not result.success
        assert result.failed == 1

    def test_missing_paragraph_fails(self, tmp_path: Path):
        doc = Document()
        doc.add_paragraph("only one paragraph")
        plan = EditPlan(operations=[ReplaceTextOp(paragraph="does not exist", old="a", new="b")])
        result = execute_plan(doc, plan)
        assert not result.success

    def test_stops_at_first_failure(self, tmp_path: Path):
        doc = Document()
        p1 = doc.add_paragraph("first paragraph")
        p2 = doc.add_paragraph("second paragraph")
        plan = EditPlan(
            operations=[
                ReplaceTextOp(paragraph="nonexistent", old="a", new="b"),
                ReplaceTextOp(paragraph=p1.text, old="first", new="FIRST"),
            ]
        )
        result = execute_plan(doc, plan)
        assert not result.success
        assert len(result.results) == 1  # second op never attempted
        assert p1.text == "first paragraph"


class TestReplaceParagraph:
    def test_preserves_style_and_first_run_formatting(self, tmp_path: Path):
        doc = Document()
        p = doc.add_heading("Old Heading", level=1)
        # Force direct bold on the heading run (headings are already bold via
        # style; we check the style name is what's preserved).
        plan = EditPlan(operations=[ReplaceParagraphOp(text=p.text, new_text="New Heading")])
        result = execute_plan(doc, plan)
        assert result.success, result.report()
        assert p.text == "New Heading"
        assert p.style.name == "Heading 1"

    def test_mixed_runs_collapse_to_single_formatted_run(self):
        doc = Document()
        p = doc.add_paragraph()
        a = p.add_run("Bold part ")
        a.bold = True
        p.add_run("and normal part")
        plan = EditPlan(operations=[ReplaceParagraphOp(text=p.text, new_text="all new text")])
        result = execute_plan(doc, plan)
        assert result.success, result.report()
        assert p.text == "all new text"
        # First run's direct formatting (bold) is carried over.
        assert p.runs[0].bold is True

    def test_ambiguous_fails(self):
        doc = Document()
        doc.add_paragraph("dup")
        doc.add_paragraph("dup")
        plan = EditPlan(operations=[ReplaceParagraphOp(text="dup", new_text="x")])
        result = execute_plan(doc, plan)
        assert not result.success


class TestInsertParagraph:
    def test_insert_after_inherits_style(self, tmp_path: Path):
        doc = Document()
        anchor = doc.add_paragraph("anchor line", style="Intense Quote")
        plan = EditPlan(
            operations=[InsertParagraphOp(anchor=anchor.text, text="inserted line", position="after")]
        )
        result = execute_plan(doc, plan)
        assert result.success, result.report()
        texts = body_paragraph_texts(doc)
        assert texts == ["anchor line", "inserted line"]
        inserted = body_paragraphs(doc)[1]
        assert inserted.style.name == "Intense Quote"

    def test_insert_before(self):
        doc = Document()
        anchor = doc.add_paragraph("anchor line")
        plan = EditPlan(
            operations=[InsertParagraphOp(anchor=anchor.text, text="before line", position="before")]
        )
        result = execute_plan(doc, plan)
        assert result.success, result.report()
        assert body_paragraph_texts(doc) == ["before line", "anchor line"]

    def test_inserted_paragraph_referenceable_by_later_ops(self):
        doc = Document()
        doc.add_paragraph("anchor line")
        plan = EditPlan(
            operations=[
                InsertParagraphOp(anchor="anchor line", text="temp line", position="after"),
                ReplaceTextOp(paragraph="temp line", old="temp", new="TEMP"),
            ]
        )
        result = execute_plan(doc, plan)
        assert result.success, result.report()
        assert body_paragraph_texts(doc) == ["anchor line", "TEMP line"]


class TestRemoveParagraph:
    def test_remove_one(self):
        doc = Document()
        doc.add_paragraph("keep me")
        doc.add_paragraph("remove me")
        doc.add_paragraph("keep me too")
        plan = EditPlan(operations=[RemoveParagraphOp(text="remove me")])
        result = execute_plan(doc, plan)
        assert result.success, result.report()
        assert body_paragraph_texts(doc) == ["keep me", "keep me too"]

    def test_remove_all_explicit(self):
        doc = Document()
        doc.add_paragraph("a")
        doc.add_paragraph("b")
        plan = EditPlan(operations=[RemoveParagraphOp(text=None)])
        result = execute_plan(doc, plan)
        assert result.success, result.report()
        assert body_paragraph_texts(doc) == []

    def test_remove_missing_fails(self):
        doc = Document()
        doc.add_paragraph("a")
        plan = EditPlan(operations=[RemoveParagraphOp(text="ghost")])
        result = execute_plan(doc, plan)
        assert not result.success


class TestTableOps:
    def test_add_row(self):
        doc = Document()
        t = doc.add_table(rows=1, cols=2)
        t.rows[0].cells[0].text = "A"
        t.rows[0].cells[1].text = "B"
        plan = EditPlan(operations=[AddTableRowOp(table=0, cells=["1", "2"])])
        result = execute_plan(doc, plan)
        assert result.success, result.report()
        assert len(t.rows) == 2
        assert [c.text for c in t.rows[1].cells] == ["1", "2"]
        # Header untouched.
        assert [c.text for c in t.rows[0].cells] == ["A", "B"]

    def test_remove_row(self):
        doc = Document()
        t = doc.add_table(rows=3, cols=2)
        for i in range(3):
            t.rows[i].cells[0].text = f"r{i}"
        plan = EditPlan(operations=[RemoveTableRowOp(table=0, row=1)])
        result = execute_plan(doc, plan)
        assert result.success, result.report()
        assert [r.cells[0].text for r in t.rows] == ["r0", "r2"]

    def test_add_row_wrong_cell_count_fails(self):
        doc = Document()
        t = doc.add_table(rows=1, cols=3)
        plan = EditPlan(operations=[AddTableRowOp(table=0, cells=["only", "two"])])
        result = execute_plan(doc, plan)
        assert not result.success

    def test_table_index_out_of_range_fails(self):
        doc = Document()
        plan = EditPlan(operations=[RemoveTableRowOp(table=5, row=0)])
        result = execute_plan(doc, plan)
        assert not result.success


class TestFormattingPreservation:
    def test_replace_text_keeps_siblings_run_formatting(self, tmp_path: Path):
        """Replace inside a formatted paragraph; sibling runs keep bold/color."""

        from docx.shared import RGBColor

        doc = Document()
        p = doc.add_paragraph()
        b = p.add_run("Client: ")
        b.bold = True
        p.add_run("Acme Corporation")
        c = p.add_run(" | Status: ")
        c.bold = True
        d = p.add_run("Draft")
        d.bold = True
        d.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

        plan = EditPlan(operations=[ReplaceTextOp(paragraph=p.text, old="Acme Corporation", new="Globex Corporation")])
        result = execute_plan(doc, plan)
        assert result.success, result.report()
        assert p.text == "Client: Globex Corporation | Status: Draft"

        runs = p.runs
        assert runs[0].text == "Client: "
        assert runs[0].bold is True
        # The colored 'Draft' run must still be there, still red+bold.
        draft = [r for r in runs if r.text == "Draft"]
        assert draft, f"runs: {[r.text for r in runs]}"
        assert draft[0].bold is True
        assert str(draft[0].font.color.rgb) == "C00000"

    def test_executor_keeps_image_part(self, tmp_path: Path):
        """Mutation must not drop the image part from the package."""

        from tests.conftest import make_rich_doc

        src = tmp_path / "rich.docx"
        make_rich_doc(src)
        doc = Document(str(src))
        plan = EditPlan(
            operations=[
                ReplaceTextOp(
                    paragraph="Prepared by: Jane Analyst  (internal)",
                    old="Jane Analyst",
                    new="John Analyst",
                )
            ]
        )
        result = execute_plan(doc, plan)
        assert result.success, result.report()
        out = tmp_path / "out.docx"
        doc.save(str(out))

        from word_agent.docx_toolkit.verify import verify

        v = verify(str(src), str(out), plan)
        assert v.ok, v.report()
