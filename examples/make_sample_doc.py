"""Generate the synthetic sample document used in the demo and tests.

All content is fictional (per the case-study practical note). Run:

    python examples/make_sample_doc.py
"""

from __future__ import annotations

import io
import struct
import zlib
from pathlib import Path

from docx import Document
from docx.shared import Pt, RGBColor

HERE = Path(__file__).parent
OUT = HERE / "input.docx"


def _solid_png(width: int = 16, height: int = 16, rgb: tuple[int, int, int] = (40, 90, 200)) -> bytes:
    """Build a small solid-colour PNG using only the standard library."""

    def chunk(tag: bytes, data: bytes) -> bytes:
        return (
            struct.pack(">I", len(data))
            + tag
            + data
            + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)
        )

    row = b"\x00" + bytes(rgb) * width  # filter byte 0 + RGB pixels
    ihdr = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)  # 8-bit RGB
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", ihdr)
        + chunk(b"IDAT", zlib.compress(row * height))
        + chunk(b"IEND", b"")
    )


def build() -> Document:
    doc = Document()

    # --- Title block ----------------------------------------------------
    doc.add_paragraph("Therapy Session Report", style="Title")
    subtitle = doc.add_paragraph("Confidential — Clinical Documentation")
    subtitle.runs[0].italic = True

    # --- Header info paragraph with mixed inline formatting -------------
    p = doc.add_paragraph()
    r1 = p.add_run("Client: ")
    r1.bold = True
    p.add_run("Acme Corporation")
    r2 = p.add_run("  |  Therapist: ")
    r2.bold = True
    p.add_run("Dr. Smith")
    r3 = p.add_run("  |  Status: ")
    r3.bold = True
    r4 = p.add_run("Draft")
    r4.bold = True
    r4.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

    # --- Body -----------------------------------------------------------
    doc.add_heading("Session Summary", level=1)
    doc.add_paragraph(
        "The session focused on stress management techniques. The client "
        "reported improved sleep quality since the last visit."
    )
    doc.add_paragraph(
        "Homework: practice the breathing exercise twice daily for the next week."
    )

    # --- Image (embedded media that integrity checks must preserve) ------
    doc.add_picture(io.BytesIO(_solid_png()), width=Pt(72))

    # --- Table: upcoming appointments ------------------------------------
    doc.add_heading("Upcoming Appointments", level=1)
    table = doc.add_table(rows=1, cols=4)
    table.style = "Light Grid Accent 1"
    hdr = table.rows[0].cells
    for cell, text in zip(hdr, ["Date", "Therapist", "Type", "Notes"]):
        cell.text = text
        for run in cell.paragraphs[0].runs:
            run.bold = True
    row1 = table.add_row().cells
    for cell, text in zip(row1, ["2025-11-20", "Dr. Smith", "Follow-up", "Review homework"]):
        cell.text = text
    row2 = table.add_row().cells
    for cell, text in zip(row2, ["2025-12-04", "Dr. Jones", "Assessment", "Quarterly review"]):
        cell.text = text

    # --- Footer note ------------------------------------------------------
    doc.add_paragraph("End of report.")

    doc.save(str(OUT))
    return doc


if __name__ == "__main__":
    build()
    print(f"Wrote {OUT}")
