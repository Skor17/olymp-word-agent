# word-agent — AI-powered Word document editor

A **LangGraph-based AI agent** that edits existing Microsoft Word (`.docx`)
documents from **natural-language instructions** without damaging their
structure or formatting.

Built for the Olymp engineering case study (Nov 2026).

---

## Quick start

**Easiest path (Windows):** double-click [`run_demo.bat`](run_demo.bat). It
creates the venv if missing, regenerates the sample document, runs the agent
on the sample instruction (no API key needed), and runs the test suite —
with a clear pass/fail report at the end.

Manual setup:

```bash
# 1. Create a virtualenv and install dependencies (Python 3.10+)
python -m venv .venv
.venv\Scripts\activate            # Windows   (Linux/macOS: source .venv/bin/activate)
pip install -r requirements.txt

# 2. (Re)generate the synthetic sample document — all content is fictional
python examples/make_sample_doc.py

# 3. Run the agent on the sample instruction (deterministic, no API key needed)
python -m word_agent.cli --planner mock examples/input.docx -o examples/output.docx ^
  --instruction-file examples/instruction.txt
```

`examples/output.docx` is the committed result of exactly that command.
(On any shell you can also pass the instruction as a quoted argument instead
of `--instruction-file`; the file form simply avoids quote-escaping issues.)

### Using a real LLM

```bash
set OPENAI_API_KEY=sk-...            # any OpenAI-compatible endpoint
set OPENAI_BASE_URL=https://api.openai.com/v1   # optional
set WORD_AGENT_MODEL=gpt-4o-mini                     # optional

python -m word_agent.cli --planner llm examples/input.docx -o out.docx \
  "Replace the client name with Globex Corporation and mark the report as Final."
```

The LLM only *plans*; see "Architecture" below for why that matters.

### Tests

```bash
python -m pytest tests -q        # 40 tests, all offline
```

---

## What it can do

| Operation | Example instruction (mock planner grammar) |
|---|---|
| Replace text inside a paragraph | `Replace "Acme Corporation" with "Globex Corporation"` |
| Replace a whole paragraph | `Replace the paragraph "Old text" with "New text"` |
| Insert a paragraph | `Insert "New line" after "Existing line"` |
| Remove a paragraph | `Remove the paragraph "Old line"` |
| Add a table row | `Add a row to table 0: "a", "b"` |
| Remove a table row | `Remove row 1 from table 0` |

With `--planner llm`, any natural-language phrasing that a chat model can
interpret is supported — the model is shown a numbered digest of the document
and must ground every operation in actual paragraph text.

`--summary` prints that digest without editing:

```bash
python -m word_agent.cli --planner mock --summary examples/input.docx
```

---

## Architecture

```
                    ┌─────────────────────────────────────────────┐
                    │              LangGraph (agent.py)            │
 instruction +      │                                             │
 input.docx ───────►│  plan ──► execute ──► verify ──► END        │
                    │    ▲              │         │               │
                    │    └──────────────┴─────────┘               │
                    │      (bounded repair loop, ≤ N attempts,    │
                    │       each starting from a pristine copy)   │
                    └─────────────────────────────────────────────┘
```

### The key decision: the LLM plans, deterministic code executes

The system is split into a **non-deterministic planner** and a
**deterministic execution/verification core**. The LLM never touches the
document. It emits a typed `EditPlan` (pydantic, JSON tool call) — a fixed
set of six operations with strict schemas:

```json
{"type": "replace_text", "paragraph": "<exact full paragraph text>",
 "old": "Acme Corporation", "new": "Globex Corporation"}
```

All document mutation happens in `word_agent/docx_toolkit/mutate.py`, which
implements those six operations directly on the OOXML tree via
`python-docx`. There is **no code execution, no shell access, no
prebuilt coding-agent harness** anywhere in the runtime path — the model's
entire power is to fill in the fields of a fixed schema, which is precisely
the property the case study asks for.

### Graph nodes

* **`plan`** — loads a *pristine* copy of the input, builds a numbered
  digest, and asks the planner for an `EditPlan`. Planner errors become
  `feedback` for the next attempt.
* **`execute`** — applies the plan through the `Executor`. A failing
  operation aborts the plan (no half-applied partial edits are ever saved)
  and its error message becomes feedback.
* **`verify`** — writes the edited document to a **temporary** file and runs
  the integrity verifier against the original. Only a passing document is
  promoted to the requested output path. **An unverified document is never
  written.** A failing verification becomes feedback and the whole attempt
  is redone from scratch.

The loop is bounded (`--max-attempts`, default 3), so the agent always
terminates and always fails *cleanly* (exit code 1, no output file,
human-readable reason) rather than guessing.

### Why this shape (trade-offs)

* **Verification as a graph node, not a post-hoc check.** Placing the
  integrity gate inside the graph (and making it able to *reject* the write)
  turns "is the document intact?" from a CI nicety into a first-class part
  of the control flow. It is also what enables the repair loop: the verifier's
  specific complaint is fed back to the planner.
* **Small, typed operation set instead of free-form code generation.** A
  fixed schema shrinks the attack/bug surface to six well-tested code paths
  and makes the whole pipeline auditable. The cost — less expressive edits —
  is accepted deliberately; unknown instructions fail cleanly instead of
  being attempted.
* **Two planners behind one interface.** `LLMPlanner` (httpx → any
  OpenAI-compatible endpoint, JSON tool call, no LangChain dependency) and
  `MockPlanner` (deterministic rule interpreter). The graph is identical for
  both, which is what makes the tests network-free and reproducible.

### Reference resolution: the subtle part

A plan is written against the document *as the planner saw it*. When two
operations touch the same paragraph ("replace Acme with Globex, then
replace Draft with Final"), the second operation still quotes the **original**
text. The `Executor` therefore resolves every paragraph reference against
**two frames**: the current live document, and the original reference frame
captured at the start of execution (kept index-aligned through inserts and
removes). Ambiguous or unresolvable references are hard errors, never
guesses. The verifier independently simulates this same two-frame logic, so
"the plan's expected result" is computed without trusting the executor.

### Integrity verification (`docx_toolkit/verify.py`)

The verifier does not diff-and-guess. It **re-simulates the plan** to produce
the expected end state, then checks the actual document against it:

1. **Paragraphs** — the simulated result is a *labeled* sequence: each
   paragraph is `kept` (with its original index), `replaced`, or
   `inserted`. The actual document must match this sequence in text, order
   and count; every `kept` paragraph must additionally be
   **formatting-identical** to the original (style name plus per-run
   bold/italic/underline/strike/font/size/color/alignment digests).
2. **Tables** — untouched tables must be formatting-identical (full cell
   digest, including per-cell paragraph formatting); planned row
   add/remove must match the simulated grid exactly.
3. **Images** — every `a:blip` drawing reference present before must still
   be rendered after, and none may be added (the op set cannot add images,
   so any appearance is suspicious).

---

## How we tested that content & formatting outside the edit remain intact

Three complementary layers, all in `tests/` (40 tests, fully offline):

1. **Executor unit tests** (`test_execute.py`) — each operation in
   isolation, including edge cases: replacements spanning multiple runs,
   in-run suffix preservation, two replacements on the same paragraph via
   the original-text frame, ambiguity detection, stop-at-first-failure
   semantics, and a formatting-preservation test that checks the bold
   labels and the *red, bold "Draft"* run are byte-for-byte intact after a
   sibling replacement.

2. **Verifier safety-net tests** (`test_verify.py`) — the core of the
   integrity claim. Each test builds a document, applies a *legitimate*
   plan, then **sabotages** the result in the exact way a broken tool would:
   * silently edit an unrelated paragraph → rejected
   * un-bold / recolor an unrelated run → rejected
   * delete an unrelated paragraph → rejected
   * inject a new paragraph → rejected
   * drop the embedded image → rejected
   * mutate a table cell the plan didn't touch → rejected

   And the mirror set: every *planned* edit (text replace, insert, remove,
   row add, mixed plans, image-bearing documents) is **accepted**. This
   proves the gate both works and doesn't produce false positives on
   legitimate edits.

3. **End-to-end agent tests** (`test_agent.py`) — the full LangGraph loop:
   multi-op instructions succeed and the input file is bit-identical
   afterwards; unparseable instructions fail cleanly after exactly
   `max_attempts` with **no output file**; a broken planner (ops that can't
   execute) proves the re-plan routing actually loops back to `plan` three
   times; and a sabotaged executor (injected text after a successful edit)
   proves the deepest guarantee — **a document that fails verification is
   never written to the output path**.

The committed `examples/input.docx` → `examples/output.docx` pair is the
human-auditable sample: open both in Word and confirm (a) the three planned
changes are present and (b) the title, headings, styled/red text, the image,
and the pre-existing table rows are visually identical.

---

## Project layout

```
word-agent/
├── word_agent/
│   ├── agent.py               # LangGraph: plan → execute → verify (repair loop)
│   ├── cli.py                 # command-line interface
│   ├── plan/
│   │   ├── schema.py          # typed EditPlan operations (pydantic)
│   │   ├── base.py            # Planner contract + PlannerContext
│   │   ├── llm.py             # OpenAI-compatible planner (JSON tool call)
│   │   └── mock.py            # deterministic rule-based planner (offline)
│   └── docx_toolkit/
│       ├── snapshot.py        # content + formatting fingerprints
│       ├── reader.py          # document summary & lookups
│       ├── mutate.py          # the ONLY place that mutates a document
│       └── verify.py          # integrity verification (plan re-simulation)
├── tests/                     # 40 offline tests
├── examples/
│   ├── make_sample_doc.py     # builds the synthetic sample document
│   ├── input.docx             # sample input (synthetic data)
│   ├── instruction.txt        # the documented sample instruction
│   └── output.docx            # result of the documented instruction
├── run_demo.bat               # double-clickable end-to-end demo (Windows)
└── requirements.txt
```

---

## Known limitations (deliberate, transparent)

* **Body-level content only.** The executor/verifier work on the main
  document body: paragraphs, top-level tables, and drawings rendered there.
  Text inside text boxes, headers/footers, footnotes, comments, and tracked
  changes is *not* editable and is *not* covered by verification (the
  underlying XML is untouched, so it passes through unchanged — but we don't
  claim to protect it at the semantic level).
* **Exact-text grounding.** Operations reference paragraphs by exact full
  text. This is a feature (no fuzzy mis-edits) but means the planner — human
  or model — must copy paragraph text verbatim from the document digest.
  Duplicate paragraphs are refused as ambiguous rather than guessed at.
* **`replace_text` match location.** The replacement is written into the run
  where the match starts; if that run's formatting differs from the rest of
  the match span, the new text takes the starting run's formatting. Matches
  inside hyperlinks/smartTags are refused (not supported), not attempted.
* **`replace_paragraph` formatting.** The new text keeps the paragraph style
  plus the *first run's* direct formatting; multi-run mixed formatting within
  a replaced paragraph is collapsed by design (the whole text is new).
* **Table addressing is positional** (`table 0`, `row 2`), with cell counts
  validated against the actual column count.
* **Single document, single instruction per run.** Batch/queue processing
  would be a thin wrapper over `run_agent` but is out of scope.
* **Verifier scope = executor scope.** If an operation type is added, both
  the executor and the verifier simulation must be extended together; a
  plan using an op the verifier doesn't understand would be rejected as an
  unexpected change, which is the safe direction to fail.
