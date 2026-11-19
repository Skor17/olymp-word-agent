"""End-to-end tests of the LangGraph agent (plan -> execute -> verify)."""

from __future__ import annotations

import sys
from pathlib import Path

import pytest
from docx import Document

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from word_agent.agent import run_agent
from word_agent.docx_toolkit.mutate import execute_plan
from word_agent.docx_toolkit.reader import body_paragraph_texts
from word_agent.plan.base import Planner, PlannerContext, PlannerError
from word_agent.plan.mock import MockPlanner
from word_agent.plan.schema import EditPlan, Plan, ReplaceTextOp


class TestHappyPath:
    def test_multi_op_instruction_succeeds(self, tmp_path: Path):
        from tests.conftest import make_rich_doc

        src = tmp_path / "in.docx"
        make_rich_doc(src)
        out = tmp_path / "out.docx"

        report = run_agent(
            instruction='Replace "Jane Analyst" with "John Analyst"',
            input_path=str(src),
            output_path=str(out),
            planner=MockPlanner(),
        )
        assert report.success, report.pretty()
        assert out.exists()
        texts = body_paragraph_texts(Document(str(out)))
        assert "Prepared by: John Analyst  (internal)" in texts
        # Everything else intact.
        assert "This is a body paragraph that should never be touched." in texts

    def test_table_and_text_in_one_instruction(self, tmp_path: Path):
        from tests.conftest import make_rich_doc

        src = tmp_path / "in.docx"
        make_rich_doc(src)
        out = tmp_path / "out.docx"

        report = run_agent(
            instruction='Replace "Jane Analyst" with "John Analyst", add a row to table 0: "South", "7", "9"',
            input_path=str(src),
            output_path=str(out),
            planner=MockPlanner(),
        )
        assert report.success, report.pretty()
        doc = Document(str(out))
        assert [c.text for c in doc.tables[0].rows[2].cells] == ["South", "7", "9"]

    def test_input_file_is_unchanged(self, tmp_path: Path):
        from tests.conftest import make_rich_doc

        src = tmp_path / "in.docx"
        make_rich_doc(src)
        before_bytes = src.read_bytes()
        out = tmp_path / "out.docx"

        report = run_agent(
            instruction='Replace "Jane Analyst" with "John Analyst"',
            input_path=str(src),
            output_path=str(out),
            planner=MockPlanner(),
        )
        assert report.success
        assert src.read_bytes() == before_bytes, "input document must not be modified in place"


class TestFailurePaths:
    def test_unparseable_instruction_fails_cleanly(self, tmp_path: Path):
        from tests.conftest import make_rich_doc

        src = tmp_path / "in.docx"
        make_rich_doc(src)
        out = tmp_path / "out.docx"

        report = run_agent(
            instruction="please rewrite everything in French, thank you very much",
            input_path=str(src),
            output_path=str(out),
            planner=MockPlanner(),
        )
        assert not report.success
        assert report.attempts == 3  # bounded: exactly max_attempts
        assert not out.exists(), "no output may be written on failure"
        assert "unrecognised clause" in " ".join(report.trace)

    def test_execution_failure_triggers_replan_then_failure(
        self, tmp_path: Path
    ):
        """A planner that emits an op the executor cannot satisfy must cause a
        clean failure with no output — and the graph must actually loop back
        to plan with feedback (proving the LangGraph routing works)."""

        from tests.conftest import make_rich_doc

        src = tmp_path / "in.docx"
        make_rich_doc(src)
        out = tmp_path / "out.docx"

        calls = {"n": 0}

        class BrokenPlanner(Planner):
            name = "broken"

            def plan(self, instruction: str, context: PlannerContext) -> Plan:
                calls["n"] += 1
                # Always propose a replace whose paragraph does not exist.
                return Plan(
                    edit_plan=EditPlan(
                        operations=[
                            ReplaceTextOp(paragraph="PARAGRAPH THAT DOES NOT EXIST", old="a", new="b")
                        ]
                    ),
                    source=self.name,
                )

        report = run_agent(
            instruction="anything",
            input_path=str(src),
            output_path=str(out),
            planner=BrokenPlanner(),
            max_attempts=3,
        )
        assert not report.success
        assert not out.exists()
        assert calls["n"] == 3, "planner must be re-invoked on execution failure"
        assert report.attempts == 3

    def test_verification_failure_never_writes_output(self, tmp_path: Path):
        """The deepest safety guarantee: even if execution 'succeeds', a
        document that fails integrity verification is never written.

        We simulate a lossy/broken editor by sabotaging the in-memory document
        *after* the planned edit is applied (injecting text into an unrelated
        paragraph). The verifier must reject it on every attempt and no output
        file may ever appear.
        """

        from tests.conftest import make_rich_doc

        src = tmp_path / "in.docx"
        make_rich_doc(src)
        out = tmp_path / "out.docx"

        import word_agent.agent as agent_mod
        from word_agent.docx_toolkit.reader import body_paragraphs

        original_execute = agent_mod.execute_plan

        def sabotaged_execute(doc, plan):
            result = original_execute(doc, plan)
            # Sabotage AFTER the planned edit: alter an unrelated paragraph.
            ps = body_paragraphs(doc)
            if len(ps) > 2:
                ps[2].add_run(" INJECTED DAMAGE")
            return result

        agent_mod.execute_plan = sabotaged_execute
        try:
            report = run_agent(
                instruction='Replace "Jane Analyst" with "John Analyst"',
                input_path=str(src),
                output_path=str(out),
                planner=MockPlanner(),
                max_attempts=2,
            )
        finally:
            agent_mod.execute_plan = original_execute

        assert not report.success, "sabotaged output must fail verification"
        assert not out.exists(), "verification failure must prevent the output write"
        assert any("integrity verification failed" in line for line in report.trace)

    def test_planner_error_surface(self, tmp_path: Path):
        from tests.conftest import make_rich_doc

        src = tmp_path / "in.docx"
        make_rich_doc(src)

        class ExplodingPlanner(Planner):
            name = "exploding"

            def plan(self, instruction: str, context: PlannerContext) -> Plan:
                raise PlannerError("simulated model outage")

        report = run_agent(
            instruction="whatever",
            input_path=str(src),
            output_path=str(tmp_path / "out.docx"),
            planner=ExplodingPlanner(),
            max_attempts=2,
        )
        assert not report.success
        assert not (tmp_path / "out.docx").exists()
        assert "simulated model outage" in " ".join(report.trace)
