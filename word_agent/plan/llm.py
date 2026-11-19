"""Real LLM planner: OpenAI-compatible chat completions, JSON tool call.

Configured entirely through environment variables:

* ``OPENAI_API_KEY``       – API key (required)
* ``OPENAI_BASE_URL``      – endpoint (default: https://api.openai.com/v1)
* ``WORD_AGENT_MODEL``     – model name (default: gpt-4o-mini)

The model is asked to emit a single tool call ``propose_edit_plan`` whose
arguments are the :class:`EditPlan` JSON schema. The response is validated
with pydantic; invalid plans raise :class:`PlannerError`, which the graph
turns into a bounded re-planning attempt (with the validation error fed back).

No LangChain dependency is required — the transport is a plain httpx call.
"""

from __future__ import annotations

import json
import os

import httpx

from word_agent.plan.base import Planner, PlannerContext, PlannerError
from word_agent.plan.schema import EditPlan, Plan

SYSTEM_PROMPT = """You are a precise document-editing planner. You receive a user's \
natural-language instruction and a numbered digest of the target .docx document. \
Produce a minimal, exact plan of editing operations that fulfils the instruction \
and changes nothing else.

Rules:
- Every `paragraph` / `text` / `anchor` field must be the EXACT full text of an \
existing paragraph, copied verbatim from the digest (including spacing). \
The executor matches by exact string equality.
- Use the smallest operation that does the job: `replace_text` for changing a \
phrase inside a paragraph; `replace_paragraph` to change a whole paragraph's text; \
`insert_paragraph` / `remove_paragraph` for adding/removing paragraphs; \
`add_table_row` / `remove_table_row` for tables.
- Do NOT invent paragraph text. If the instruction refers to something not \
present in the document, return an empty operation list and explain in `notes`.
- `remove_paragraph` with `text: null` removes ALL paragraphs; use it only when \
the instruction explicitly asks for that.
- Prefer fewer operations. The plan must be deterministic: the same document + \
instruction must always yield the same plan.
"""

PLAN_TOOL = {
    "type": "function",
    "function": {
        "name": "propose_edit_plan",
        "description": "Propose the complete edit plan for the instruction.",
        "parameters": EditPlan.model_json_schema(),
    },
}


class LLMPlanner(Planner):
    """Production planner backed by an OpenAI-compatible chat endpoint."""

    name = "llm"

    def __init__(
        self,
        model: str | None = None,
        base_url: str | None = None,
        api_key: str | None = None,
        temperature: float = 0.0,
        timeout: float = 60.0,
        max_output_tokens: int = 2048,
    ) -> None:
        self.model = model or os.environ.get("WORD_AGENT_MODEL", "gpt-4o-mini")
        self.base_url = (base_url or os.environ.get("OPENAI_BASE_URL", "https://api.openai.com/v1")).rstrip("/")
        self.api_key = api_key or os.environ.get("OPENAI_API_KEY", "")
        self.temperature = temperature
        self.timeout = timeout
        self.max_output_tokens = max_output_tokens
        if not self.api_key:
            raise PlannerError("OPENAI_API_KEY is not set")

    # ------------------------------------------------------------------ #
    def plan(self, instruction: str, context: PlannerContext) -> Plan:
        user = self._user_message(instruction, context.summary, None)
        return Plan(edit_plan=self._raw_chat(user), source=self.name)

    def replan(self, instruction: str, context: PlannerContext, feedback: str) -> Plan:
        """Re-plan after a failure, feeding the rejection reason back to the model."""
        user = self._user_message(instruction, context.summary, feedback)
        return Plan(edit_plan=self._raw_chat(user), source=self.name)

    # ------------------------------------------------------------------ #
    @staticmethod
    def _user_message(instruction: str, summary: str, feedback: str | None) -> str:
        parts = [f"INSTRUCTION:\n{instruction}", f"DOCUMENT DIGEST:\n{summary}"]
        if feedback:
            parts.append(f"Your previous plan was rejected:\n{feedback}\nProduce a corrected plan.")
        return "\n\n".join(parts)

    def _raw_chat(self, user_message: str) -> EditPlan:
        body = {
            "model": self.model,
            "temperature": self.temperature,
            "max_tokens": self.max_output_tokens,
            "tools": [PLAN_TOOL],
            "tool_choice": {"type": "function", "function": {"name": "propose_edit_plan"}},
            "messages": [
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": user_message},
            ],
        }
        try:
            resp = httpx.post(
                f"{self.base_url}/chat/completions",
                headers={"Authorization": f"Bearer {self.api_key}"},
                json=body,
                timeout=self.timeout,
            )
            resp.raise_for_status()
        except httpx.HTTPError as e:
            raise PlannerError(f"LLM request failed: {e}") from e

        data = resp.json()
        try:
            choice = data["choices"][0]["message"]
            tool_calls = choice.get("tool_calls") or []
            if tool_calls:
                args = json.loads(tool_calls[0]["function"]["arguments"])
            else:
                # Some providers return raw JSON in content despite tool_choice.
                args = json.loads(choice["content"])
        except (KeyError, IndexError, json.JSONDecodeError) as e:
            raise PlannerError(f"Could not parse model response: {e}\n{data}") from e

        try:
            return EditPlan.model_validate(args)
        except Exception as e:
            raise PlannerError(f"Model returned an invalid plan: {e}") from e
