"""The LangGraph agent.

Graph topology (bounded repair loop, at most ``max_attempts`` full tries):

    START -> plan -> execute -> verify -> END
              ^          |          |
              |          |          |
              +----------+----------+
              (on failure, if attempts remain: re-plan with feedback)

* **plan**    – the (possibly LLM-based) planner produces an :class:`EditPlan`
  for a *freshly loaded* copy of the document. Planner errors become feedback
  for the next attempt.
* **execute** – the deterministic executor applies the plan to the in-memory
  document. A failing operation aborts; the per-operation report becomes
  feedback for re-planning.
* **verify**  – the edited document is written to a temporary file and
  integrity-checked against the original (content *and* formatting, images,
  tables). Only if verification passes is the temporary file promoted to the
  requested output path. A failing verification produces feedback and the
  whole attempt is redone from a pristine copy of the original.

The key safety property: **an unverified document is never written to the
output path.**
"""

from __future__ import annotations

import os
import shutil
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional, TypedDict

from docx import Document
from langgraph.graph import END, START, StateGraph

from word_agent.docx_toolkit.mutate import ExecutionResult, execute_plan
from word_agent.docx_toolkit.reader import document_summary
from word_agent.docx_toolkit.verify import VerificationResult, verify
from word_agent.plan.base import Planner, PlannerContext, PlannerError
from word_agent.plan.schema import Plan

MAX_ATTEMPTS_DEFAULT = 3


# --------------------------------------------------------------------------- #
# State
# --------------------------------------------------------------------------- #
class AgentState(TypedDict, total=False):
    # Inputs (set once by the caller)
    instruction: str
    input_path: str
    output_path: str

    # Working data
    doc: Optional[Document]  # fresh copy of the input, mutated by execute
    plan: Optional[Plan]
    feedback: Optional[str]  # reason the previous attempt failed
    attempts: int

    # Results
    exec_result: Optional[ExecutionResult]
    verification: Optional[VerificationResult]
    status: str  # "ok" | "failed"
    output_written: bool
    trace: list  # human-readable per-node log lines
    max_attempts: int


# --------------------------------------------------------------------------- #
# Nodes
# --------------------------------------------------------------------------- #
def _log(state: AgentState, line: str) -> None:
    state.setdefault("trace", []).append(line)


def make_plan_node(planner: Planner, max_attempts: int):
    def plan_node(state: AgentState) -> dict[str, Any]:
        attempt = state.get("attempts", 0) + 1
        state["attempts"] = attempt
        doc = Document(state["input_path"])  # pristine copy every attempt
        state["doc"] = doc
        _log(state, f"--- attempt {attempt}/{max_attempts} ---")
        _log(state, f"[plan] planner={planner.name}")

        context = PlannerContext(doc=doc, summary=document_summary(doc))
        try:
            if state.get("feedback"):
                _log(state, f"[plan] replanning with feedback: {state['feedback'][:200]!r}")
                plan = planner.replan(state["instruction"], context, state["feedback"])
            else:
                plan = planner.plan(state["instruction"], context)
        except PlannerError as e:
            state["feedback"] = f"planner error: {e}"
            if attempt >= max_attempts:
                state["status"] = "failed"
                _log(state, f"[plan] FAILED (max attempts): {e}")
            else:
                _log(state, f"[plan] error, will retry: {e}")
            return dict(state)

        state["plan"] = plan
        ops_desc = "; ".join(o.model_dump_json() for o in plan.edit_plan.operations)
        _log(state, f"[plan] {len(plan.edit_plan.operations)} op(s): {ops_desc[:500]}")
        state["feedback"] = None
        return dict(state)

    return plan_node


def make_execute_node():
    def execute_node(state: AgentState) -> dict[str, Any]:
        doc = state["doc"]
        plan = state.get("plan")
        if plan is None:
            state["feedback"] = state.get("feedback") or "no plan available for execution"
            if state["attempts"] >= state["max_attempts"]:
                state["status"] = "failed"
                _log(state, "[execute] FAILED — no plan; output NOT written")
            return dict(state)
        # Called through the module-level name (a global of this module) so
        # tests can monkeypatch word_agent.agent.execute_plan.
        result = execute_plan(doc, plan.edit_plan)
        state["exec_result"] = result
        for line in result.report().splitlines():
            _log(state, f"[execute] {line}")

        if not result.success:
            # Feedback for re-planning: which ops failed and why.
            failed = [r for r in result.results if r.status == "error"]
            state["feedback"] = "execution failed: " + " | ".join(
                f"op {r.index} ({r.op_type}) {r.message}" for r in failed
            )
            if state["attempts"] >= state["max_attempts"]:
                state["status"] = "failed"
                _log(state, "[execute] FAILED — max attempts reached; output NOT written")
        return dict(state)

    return execute_node


def make_verify_node(max_attempts: int):
    def verify_node(state: AgentState) -> dict[str, Any]:
        assert state["plan"] is not None
        doc = state["doc"]
        input_path = Path(state["input_path"])
        output_path = Path(state["output_path"])

        # Write the edited document to a temp file, verify, then promote.
        fd, tmp_name = tempfile.mkstemp(suffix=".docx", prefix="word_agent_")
        os.close(fd)
        tmp_path = Path(tmp_name)
        try:
            doc.save(tmp_path)
            result = verify(str(input_path), str(tmp_path), state["plan"].edit_plan)
        except Exception as e:  # defensive: verification infra failure == failed
            tmp_path.unlink(missing_ok=True)
            state["feedback"] = f"verification error: {type(e).__name__}: {e}"
            state["status"] = "failed"
            _log(state, f"[verify] ERROR: {e}")
            return dict(state)

        state["verification"] = result
        for line in result.report().splitlines():
            _log(state, f"[verify] {line}")

        if result.ok:
            output_path.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(str(tmp_path), str(output_path))
            state["status"] = "ok"
            state["output_written"] = True
            _log(state, f"[verify] PASS — wrote {output_path}")
        else:
            tmp_path.unlink(missing_ok=True)
            state["feedback"] = "integrity verification failed: " + "; ".join(result.problems)
            if state["attempts"] < max_attempts:
                _log(state, f"[verify] FAIL — will re-plan (attempt {state['attempts']} of {max_attempts})")
            else:
                state["status"] = "failed"
                _log(state, f"[verify] FAIL — max attempts reached; output NOT written")
        return dict(state)

    return verify_node


# --------------------------------------------------------------------------- #
# Routing
# --------------------------------------------------------------------------- #
def _route_after_plan(state: AgentState) -> str:
    if state.get("status") == "failed":
        return "end"
    if state.get("plan") is None:
        # The planner errored out: retry planning (fresh attempt) if budget
        # remains, otherwise give up.
        if state["attempts"] < state["max_attempts"]:
            return "plan"
        state["status"] = "failed"
        return "end"
    return "execute"


def _route_after_execute(state: AgentState) -> str:
    if state.get("status") == "failed":
        return "end"
    result: ExecutionResult | None = state.get("exec_result")
    if result is not None and not result.success:
        return "plan"
    return "verify"


def _route_after_verify(state: AgentState) -> str:
    if state.get("status") == "ok":
        return "end"
    if state.get("attempts", 0) < state["max_attempts"]:
        return "plan"
    return "end"


# --------------------------------------------------------------------------- #
# Graph construction
# --------------------------------------------------------------------------- #
def build_agent(planner: Planner, max_attempts: int = MAX_ATTEMPTS_DEFAULT):
    """Compile the LangGraph for a given planner."""

    graph = StateGraph(AgentState)
    graph.add_node("plan", make_plan_node(planner, max_attempts))
    graph.add_node("execute", make_execute_node())
    graph.add_node("verify", make_verify_node(max_attempts))

    graph.add_edge(START, "plan")
    graph.add_conditional_edges(
        "plan",
        _route_after_plan,
        {"execute": "execute", "plan": "plan", "end": END},
    )
    graph.add_conditional_edges(
        "execute",
        _route_after_execute,
        {"verify": "verify", "plan": "plan", "end": END},
    )
    graph.add_conditional_edges(
        "verify",
        _route_after_verify,
        {"plan": "plan", "end": END},
    )
    return graph.compile()


# --------------------------------------------------------------------------- #
# Public run helper
# --------------------------------------------------------------------------- #
@dataclass
class AgentReport:
    """Everything the CLI/UI needs to show about a run."""

    instruction: str
    input_path: str
    output_path: str
    status: str
    attempts: int
    plan: Plan | None = None
    exec_result: ExecutionResult | None = None
    verification: VerificationResult | None = None
    trace: list[str] = field(default_factory=list)
    error: str | None = None

    @property
    def success(self) -> bool:
        return self.status == "ok"

    def pretty(self) -> str:
        lines = [
            f"Instruction : {self.instruction}",
            f"Input       : {self.input_path}",
            f"Status      : {self.status.upper()}",
            f"Attempts    : {self.attempts}",
        ]
        if self.plan is not None:
            lines.append("Plan:")
            for op in self.plan.edit_plan.operations:
                lines.append(f"  - {op.model_dump_json()}")
            if self.plan.edit_plan.notes:
                lines.append(f"  notes: {self.plan.edit_plan.notes}")
        if self.exec_result is not None:
            lines.append("Execution:")
            lines.extend("  " + l for l in self.exec_result.report().splitlines())
        if self.verification is not None:
            lines.extend("  " + l for l in self.verification.report().splitlines())
        lines.append("Trace:")
        lines.extend("  " + l for l in self.trace)
        if self.error:
            lines.append(f"ERROR: {self.error}")
        return "\n".join(lines)


def run_agent(
    instruction: str,
    input_path: str,
    output_path: str,
    planner: Planner,
    max_attempts: int = MAX_ATTEMPTS_DEFAULT,
) -> AgentReport:
    """Run the full plan/execute/verify loop and return an :class:`AgentReport`."""

    app = build_agent(planner, max_attempts=max_attempts)
    initial: AgentState = {
        "instruction": instruction,
        "input_path": str(input_path),
        "output_path": str(output_path),
        "attempts": 0,
        "feedback": None,
        "max_attempts": max_attempts,
    }
    final = app.invoke(initial)
    return AgentReport(
        instruction=instruction,
        input_path=str(input_path),
        output_path=str(output_path),
        status=final.get("status", "failed"),
        attempts=final.get("attempts", 0),
        plan=final.get("plan"),
        exec_result=final.get("exec_result"),
        verification=final.get("verification"),
        trace=final.get("trace", []),
    )
