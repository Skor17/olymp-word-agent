"""word_agent — a LangGraph agent that edits .docx files from natural language.

The LLM only *plans*; all document mutation happens through the deterministic,
typed operations in :mod:`word_agent.docx_toolkit`.
"""

__version__ = "0.1.0"
