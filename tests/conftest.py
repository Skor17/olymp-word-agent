"""Shared fixtures.

``rich_doc`` builds, in memory, a document that exercises every integrity
dimension at once: a styled heading, a paragraph with *mixed inline run
formatting* (bold + colored), a normal paragraph, an embedded image, and a
table with a bolded header row. It is written to a temp file so the agent can
treat it like any other input.
"""

from __future__ import annotations

import io
import struct
import zlib
from pathlib import Path

import pytest
from docx import Document
from docx.shared import Pt, RGBColor

import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))


def _solid_png(width: int = 12, height: int = 12, rgb: tuple[int, int, int] = (200, 60, 60)) -> bytes:
    def chunk(tag: bytes, data: bytes) -> bytes:
        return (
            struct.pack(">I", len(data))
            + tag
            + data
            + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)
        )

    row = b"\x00" + bytes(rgb) * width
    ihdr = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", ihdr)
        + chunk(b"IDAT", zlib.compress(row * height))
        + chunk(b"IEND", b"")
    )


def make_rich_doc(path: Path) -> None:
    doc = Document()
    doc.add_paragraph("Quarterly Report", style="Heading 1")

    # Mixed inline formatting: label bold, value colored.
    p = doc.add_paragraph()
    r = p.add_run("Prepared by: ")
    r.bold = True
    p.add_run("Jane Analyst")
    r2 = p.add_run("  (internal)")
    r2.italic = True
    r2.font.color.rgb = RGBColor(0x00, 0x70, 0xC0)

    doc.add_paragraph("This is a body paragraph that should never be touched.")

    doc.add_picture(io.BytesIO(_solid_png()), width=Pt(60))

    table = doc.add_table(rows=2, cols=3)
    hdr = table.rows[0].cells
    for c, t in zip(hdr, ["Region", "Q1", "Q2"]):
        c.text = t
        for run in c.paragraphs[0].runs:
            run.bold = True
    row = table.rows[1].cells
    for c, t in zip(row, ["North", "10", "12"]):
        c.text = t

    doc.save(str(path))


@pytest.fixture
def rich_doc(tmp_path: Path) -> Path:
    p = tmp_path / "rich.docx"
    make_rich_doc(p)
    return p
