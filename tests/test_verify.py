"""Tests for the integrity verifier (docx_toolkit.verify).

These are the *safety net* tests: they prove the verifier (a) accepts edits
that exactly match the plan, and (b) rejects any deviation — formatting
changes in untouched paragraphs, silent text changes, image loss, table
corruption, and extra insertions — i.e. the kinds of damage the case study
warns about.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest
from docx import Document
from docx.shared import RGBColor

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from word_agent.docx_toolkit.mutate import execute_plan
from word_agent.docx_toolkit.verify import verify
from word_agent.plan.schema import (
    AddTableRowOp,
    EditPlan,
    InsertParagraphOp,
    RemoveParagraphOp,
    ReplaceParagraphOp,
    ReplaceTextOp,
)


def _save(doc: Document, path: Path) -> None:
    doc.save(str(path))


def _apply_and_verify(tmp_path: Path, build, plan: EditPlan, out_name: str = "out.docx"):
    """Build doc -> save -> execute plan -> save -> verify; return (v, out)."""

    src = tmp_path / "in.docx"
    build(src)
    _save(Document(str(src)), src)

    doc = Document(str(src))
    result = execute_plan(doc, plan)
    assert result.success, result.report()
    out = tmp_path / out_name
    _save(doc, out)
    return verify(str(src), str(out), plan), out


# --------------------------------------------------------------------------- #
# Accept: planned edits pass
# --------------------------------------------------------------------------- #
class TestVerifyAccepts:
    def test_replace_text_passes(self, tmp_path: Path):
        def build(src: Path):
            doc = Document()
            doc.add_paragraph("Client: Acme Corporation")
            doc.add_paragraph("Unrelated paragraph.")
            doc.save(str(src))

        v, _ = _apply_and_verify(
            tmp_path,
            build,
            EditPlan(
                operations=[
                    ReplaceTextOp(
                        paragraph="Client: Acme Corporation",
                        old="Acme Corporation",
                        new="Globex Corporation",
                    )
                ]
            ),
        )
        assert v.ok, v.report()

    def test_insert_and_remove_pass(self, tmp_path: Path):
        def build(src: Path):
            doc = Document()
            doc.add_paragraph("alpha")
            doc.add_paragraph("beta")
            doc.add_paragraph("gamma")
            doc.save(str(src))

        v, _ = _apply_and_verify(
            tmp_path,
            build,
            EditPlan(
                operations=[
                    InsertParagraphOp(anchor="beta", text="inserted", position="after"),
                    RemoveParagraphOp(text="gamma"),
                ]
            ),
        )
        assert v.ok, v.report()

    def test_add_table_row_passes(self, tmp_path: Path):
        def build(src: Path):
            doc = Document()
            t = doc.add_table(rows=1, cols=2)
            t.rows[0].cells[0].text = "A"
            t.rows[0].cells[1].text = "B"
            doc.save(str(src))

        v, _ = _apply_and_verify(
            tmp_path, build, EditPlan(operations=[AddTableRowOp(table=0, cells=["x", "y"])])
        )
        assert v.ok, v.report()

    def test_image_intact_reported(self, tmp_path: Path):
        from tests.conftest import make_rich_doc

        v, _ = _apply_and_verify(
            tmp_path,
            lambda src: make_rich_doc(src),
            EditPlan(
                operations=[
                    ReplaceTextOp(
                        paragraph="Prepared by: Jane Analyst  (internal)",
                        old="Jane Analyst",
                        new="John Analyst",
                    )
                ]
            ),
        )
        assert v.ok, v.report()
        assert any("images intact" in n for n in v.notes)


# --------------------------------------------------------------------------- #
# Reject: deviations from the plan are caught
# --------------------------------------------------------------------------- #
class TestVerifyRejects:
    def test_unplanned_text_change_fails(self, tmp_path: Path):
        def build(src: Path):
            doc = Document()
            doc.add_paragraph("Client: Acme")
            doc.add_paragraph("Secret body paragraph.")
            doc.save(str(src))

        # The plan only touches the first paragraph.
        plan = EditPlan(
            operations=[ReplaceTextOp(paragraph="Client: Acme", old="Acme", new="Globex")]
        )
        src = tmp_path / "in.docx"
        build(src)

        doc = Document(str(src))
        execute_plan(doc, plan)
        # Sabotage: silently change an unrelated paragraph.
        doc.paragraphs[1].runs[0].text = "Altered body paragraph."
        out = tmp_path / "out.docx"
        _save(doc, out)

        v = verify(str(src), str(out), plan)
        assert not v.ok, "verifier must catch an unplanned paragraph change"
        assert any("unexpected text" in p for p in v.problems)

    def test_unplanned_formatting_change_fails(self, tmp_path: Path):
        def build(src: Path):
            doc = Document()
            doc.add_paragraph("Client: Acme")
            p = doc.add_paragraph()
            r = p.add_run("Keep this bold")
            r.bold = True
            doc.save(str(src))

        plan = EditPlan(
            operations=[ReplaceTextOp(paragraph="Client: Acme", old="Acme", new="Globex")]
        )
        src = tmp_path / "in.docx"
        build(src)

        doc = Document(str(src))
        execute_plan(doc, plan)
        # Sabotage: un-bold an unrelated paragraph.
        doc.paragraphs[1].runs[0].bold = False
        out = tmp_path / "out.docx"
        _save(doc, out)

        v = verify(str(src), str(out), plan)
        assert not v.ok, "verifier must catch an unplanned formatting change"
        assert any("formatting changed in untouched paragraph" in p for p in v.problems)

    def test_unplanned_color_change_fails(self, tmp_path: Path):
        def build(src: Path):
            doc = Document()
            doc.add_paragraph("Client: Acme")
            p = doc.add_paragraph()
            r = p.add_run("Keep this red")
            r.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
            doc.save(str(src))

        plan = EditPlan(
            operations=[ReplaceTextOp(paragraph="Client: Acme", old="Acme", new="Globex")]
        )
        src = tmp_path / "in.docx"
        build(src)

        doc = Document(str(src))
        execute_plan(doc, plan)
        doc.paragraphs[1].runs[0].font.color.rgb = RGBColor(0x00, 0xFF, 0x00)
        out = tmp_path / "out.docx"
        _save(doc, out)

        v = verify(str(src), str(out), plan)
        assert not v.ok, "verifier must catch an unplanned color change"

    def test_unplanned_deletion_fails(self, tmp_path: Path):
        def build(src: Path):
            doc = Document()
            doc.add_paragraph("Client: Acme")
            doc.add_paragraph("Do not delete me.")
            doc.save(str(src))

        plan = EditPlan(
            operations=[ReplaceTextOp(paragraph="Client: Acme", old="Acme", new="Globex")]
        )
        src = tmp_path / "in.docx"
        build(src)

        doc = Document(str(src))
        execute_plan(doc, plan)
        # Sabotage: delete an unrelated paragraph.
        p = doc.paragraphs[1]._element
        p.getparent().remove(p)
        out = tmp_path / "out.docx"
        _save(doc, out)

        v = verify(str(src), str(out), plan)
        assert not v.ok, "verifier must catch an unplanned deletion"
        assert any("count mismatch" in p for p in v.problems)

    def test_unplanned_insertion_fails(self, tmp_path: Path):
        def build(src: Path):
            doc = Document()
            doc.add_paragraph("Client: Acme")
            doc.save(str(src))

        plan = EditPlan(
            operations=[ReplaceTextOp(paragraph="Client: Acme", old="Acme", new="Globex")]
        )
        src = tmp_path / "in.docx"
        build(src)

        doc = Document(str(src))
        execute_plan(doc, plan)
        # Sabotage: inject a new paragraph.
        doc.add_paragraph("Injected by a buggy tool.")
        out = tmp_path / "out.docx"
        _save(doc, out)

        v = verify(str(src), str(out), plan)
        assert not v.ok, "verifier must catch an unplanned insertion"
        assert any("count mismatch" in p for p in v.problems)
        assert any("unexpected paragraph" in p for p in v.problems)

    def test_image_loss_fails(self, tmp_path: Path):
        from tests.conftest import make_rich_doc

        src = tmp_path / "in.docx"
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
        execute_plan(doc, plan)

        # Sabotage: remove the paragraph wrapping the drawing (simulate a
        # lossy save that drops the image).
        from docx.oxml.ns import qn

        for drawing in list(doc.element.body.iter(qn("w:drawing"))):
            # The drawing sits inside a w:r inside a w:p; remove the whole p.
            node = drawing
            while node is not None and not node.tag.endswith("}p"):
                node = node.getparent()
            if node is not None:
                node.getparent().remove(node)
        out = tmp_path / "out.docx"
        _save(doc, out)

        v = verify(str(src), str(out), plan)
        assert not v.ok, "verifier must catch lost images"
        assert any("images lost" in p for p in v.problems)

    def test_unplanned_table_change_fails(self, tmp_path: Path):
        def build(src: Path):
            doc = Document()
            doc.add_paragraph("Client: Acme")
            t = doc.add_table(rows=2, cols=2)
            t.rows[0].cells[0].text = "A"
            t.rows[0].cells[1].text = "B"
            t.rows[1].cells[0].text = "C"
            t.rows[1].cells[1].text = "D"
            doc.save(str(src))

        plan = EditPlan(
            operations=[ReplaceTextOp(paragraph="Client: Acme", old="Acme", new="Globex")]
        )
        src = tmp_path / "in.docx"
        build(src)

        doc = Document(str(src))
        execute_plan(doc, plan)
        # Sabotage: mutate a table cell.
        doc.tables[0].rows[1].cells[0].text = "SABOTAGED"
        out = tmp_path / "out.docx"
        _save(doc, out)

        v = verify(str(src), str(out), plan)
        assert not v.ok, "verifier must catch an unplanned table change"
        assert any("table 0 changed" in p for p in v.problems)

    def test_planned_table_change_is_ok_even_with_extra_rows(self, tmp_path: Path):
        """add_table_row is allowed to change the table; other tables are not."""

        def build(src: Path):
            doc = Document()
            doc.add_paragraph("Client: Acme")
            t = doc.add_table(rows=1, cols=2)
            t.rows[0].cells[0].text = "A"
            t.rows[0].cells[1].text = "B"
            doc.save(str(src))

        v, _ = _apply_and_verify(
            tmp_path,
            build,
            EditPlan(
                operations=[
                    ReplaceTextOp(paragraph="Client: Acme", old="Acme", new="Globex"),
                    AddTableRowOp(table=0, cells=["x", "y"]),
                ]
            ),
        )
        assert v.ok, v.report()
