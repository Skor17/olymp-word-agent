"""docx_toolkit — self-implemented document tools.

* :mod:`snapshot` – content/formatting digests (fingerprinting)
* :mod:`reader`   – read-only lookups and summaries
* :mod:`mutate`   – deterministic execution of typed edit operations
* :mod:`verify`   – before/after integrity verification
"""

from word_agent.docx_toolkit.mutate import (  # noqa: F401
    Executor,
    ExecutionResult,
    OpResult,
    execute_plan,
)
from word_agent.docx_toolkit.reader import (  # noqa: F401
    LookupError_,
    body_paragraphs,
    body_paragraph_texts,
    document_summary,
    find_all_paragraph_indices,
    find_paragraph_index,
    get_table,
)
from word_agent.docx_toolkit.snapshot import Snapshot, snapshot_document  # noqa: F401
from word_agent.docx_toolkit.verify import (  # noqa: F401
    VerificationResult,
    verify,
)

__all__ = [
    "Executor",
    "ExecutionResult",
    "OpResult",
    "execute_plan",
    "LookupError_",
    "body_paragraphs",
    "body_paragraph_texts",
    "document_summary",
    "find_all_paragraph_indices",
    "find_paragraph_index",
    "get_table",
    "Snapshot",
    "snapshot_document",
    "VerificationResult",
    "verify",
]
