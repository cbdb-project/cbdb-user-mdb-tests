"""Shared helpers for the static-audit scripts.

Why this exists: dump files in `analysis/dump/vba/` use `\\r\\r\\n` line
endings (Windows double-CR quirk).  Python's universal-newlines text
mode treats each `\\r` as a line separator, so `read_text(...).splitlines()`
inflates the line count 2x and reports diagnostics with line numbers
that don't match grep / the VBE editor.

`read_vba_lines(path)` returns lines indexed the same way grep does.
"""
from __future__ import annotations
from pathlib import Path


def read_vba_lines(path: Path) -> list[str]:
    """Read a VBA dump file and return one entry per logical line,
    matching grep / VBE numbering.  Handles \\r\\r\\n, \\r\\n, and \\n
    line endings transparently."""
    text = path.read_bytes().decode("utf-8")
    return [ln.rstrip("\r") for ln in text.split("\n")]
