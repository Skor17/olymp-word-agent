"""Planner base: the contract any "LLM" in this system must fulfil.

A planner receives the user instruction plus a :class:`PlannerContext` (the
open document plus a compact text digest) and returns an :class:`EditPlan`.
The graph does not care *how* the plan is produced:

* production uses :class:`word_agent.plan.llm.LLMPlanner` (a real model),
* tests and offline runs use :class:`word_agent.plan.mock.MockPlanner`
  (a deterministic pattern interpreter with the same contract).

This separation is deliberate: the planner is the only non-deterministic
component in the system, and isolating it makes the whole pipeline testable
without network access.
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass

from docx import Document

from word_agent.plan.schema import EditPlan, Plan


class PlannerError(Exception):
    """Raised when a planner cannot produce a valid plan."""


@dataclass
class PlannerContext:
    """Everything a planner is allowed to see about the document."""

    doc: Document
    summary: str  # compact numbered digest (see reader.document_summary)


class Planner(ABC):
    """Abstract planner: instruction + context -> Plan."""

    name: str = "abstract"

    @abstractmethod
    def plan(self, instruction: str, context: PlannerContext) -> Plan:
        """Produce an :class:`EditPlan` for ``instruction``.

        The plan must be grounded in the actual document: every paragraph
        reference in an operation has to exist verbatim.
        """
        raise NotImplementedError

    def replan(self, instruction: str, context: PlannerContext, feedback: str) -> Plan:
        """Re-plan after a validation/verification failure.

        The default implementation just calls :meth:`plan` again; planners
        with memory (the LLM planner) override this to feed the error back.
        """
        del feedback
        return self.plan(instruction, context)
