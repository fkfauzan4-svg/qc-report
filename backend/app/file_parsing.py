"""Turns an uploaded spreadsheet into a plain-text dump the Claude client can
reason over. Kept separate from claude_client.py so the parsing logic is easy
to test on its own."""

import csv
import io
import os

import openpyxl

MAX_CHARS = 60_000  # keep the prompt payload bounded for very large workbooks


def extract_text(filename: str, data: bytes) -> str:
    ext = os.path.splitext(filename)[1].lower()

    if ext in (".csv", ".tsv"):
        text = data.decode("utf-8", errors="replace")
        return _truncate(text)

    if ext in (".xlsx", ".xlsm", ".xls"):
        return _truncate(_extract_excel_text(data))

    raise ValueError(f"Don't know how to extract text from '{filename}' ({ext})")


def _extract_excel_text(data: bytes) -> str:
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    lines = []
    for ws in wb.worksheets:
        lines.append(f"=== SHEET: {ws.title} ===")
        for row in ws.iter_rows():
            values = [c.value for c in row]
            if any(v is not None and str(v).strip() != "" for v in values):
                cells = [str(v) if v is not None else "" for v in values]
                lines.append("\t".join(cells))
        lines.append("")
    return "\n".join(lines)


def _truncate(text: str) -> str:
    if len(text) <= MAX_CHARS:
        return text
    return text[:MAX_CHARS] + "\n...[truncated]..."
