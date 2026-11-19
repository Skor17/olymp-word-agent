"""Command-line interface.

Examples
--------
Deterministic offline mode (no API key needed)::

    python -m word_agent.cli --planner mock examples/input.docx -o out.docx \\
        "Change \"Acme Corporation\" to \"Globex Corporation\" and add a row to table 0: \"2026-01-15\", \"Dr. Lee\", \"Follow-up\", \"Completed\""

Live LLM mode (requires OPENAI_API_KEY; base URL/model configurable)::

    python -m word_agent.cli --planner llm -m gpt-4o-mini examples/input.docx -o out.docx \\
        "Replace the client name with Globex Corporation, set the next visit date to 2026-02-01, and remove the 'Draft' line."

Useful flags: ``--summary`` (show the document digest without editing),
``--max-attempts`` (default 3), ``--quiet`` (report only on failure).
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from docx import Document

from word_agent.agent import run_agent
from word_agent.docx_toolkit.reader import document_summary
from word_agent.plan.base import Planner
from word_agent.plan.mock import MockPlanner


def _build_planner(args: argparse.Namespace) -> Planner:
    if args.planner == "mock":
        return MockPlanner()
    if args.planner == "llm":
        from word_agent.plan.llm import LLMPlanner, PlannerError

        try:
            return LLMPlanner(model=args.model)
        except PlannerError as e:
            sys.exit(f"error: {e}")
    raise ValueError(args.planner)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="word-agent",
        description="Edit .docx files with natural-language instructions (LangGraph agent).",
    )
    parser.add_argument("input", help="Path to the input .docx file")
    parser.add_argument("-o", "--output", help="Path for the edited .docx (default: <input>_edited.docx)")
    parser.add_argument("instruction", nargs="?", help="Natural-language editing instruction")
    parser.add_argument(
        "--instruction-file",
        default=None,
        help="Read the instruction from a UTF-8 file instead of the command line (avoids shell quoting issues).",
    )
    parser.add_argument(
        "--planner",
        choices=["mock", "llm"],
        default="mock",
        help="Planner to use: 'mock' (deterministic, offline) or 'llm' (OpenAI-compatible). Default: mock.",
    )
    parser.add_argument("--model", default=None, help="Model name for --planner llm (env WORD_AGENT_MODEL or gpt-4o-mini)")
    parser.add_argument("--max-attempts", type=int, default=3, help="Maximum plan/execute/verify attempts (default 3)")
    parser.add_argument("--summary", action="store_true", help="Print the document digest and exit")
    parser.add_argument("--quiet", action="store_true", help="Only print output on failure")
    args = parser.parse_args(argv)

    if not args.input.endswith(".docx"):
        sys.exit("error: input must be a .docx file")

    doc = Document(args.input)
    if args.summary:
        print(document_summary(doc))
        return 0

    instruction = args.instruction
    if args.instruction_file:
        try:
            instruction = Path(args.instruction_file).read_text(encoding="utf-8").strip()
        except OSError as e:
            sys.exit(f"error: cannot read instruction file: {e}")
    if not instruction:
        parser.error(
            "an editing instruction is required (as an argument or via --instruction-file), "
            "or use --summary"
        )

    output = args.output or args.input.rsplit(".", 1)[0] + "_edited.docx"
    planner = _build_planner(args)

    report = run_agent(
        instruction=instruction,
        input_path=args.input,
        output_path=output,
        planner=planner,
        max_attempts=args.max_attempts,
    )

    if report.success:
        if not args.quiet:
            print(report.pretty())
            print(f"\nDone. Edited document written to: {output}")
        return 0

    print(report.pretty(), file=sys.stderr)
    print(f"\nFAILED. Output was NOT written: {output}", file=sys.stderr)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
