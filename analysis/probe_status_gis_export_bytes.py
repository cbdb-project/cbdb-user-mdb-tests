"""Run LookAtStatus.CmdGIS once and dump byte-level offending rows.

Companion to probe_status_gis_embedded_delim.py.  Where that one
scans the source data, this one runs the full export pipeline so we
can see what actually gets written to disk.

Why this exists: pyodbc reads ADDR_CODES.c_name_chn as `﻿尉氏`
(3 chars, no embedded TAB) for the row that triggers the test
failure.  But the exported `.tab` row has 10 tab-separated cells
against a 9-col header — i.e. the export *did* write a tab
somewhere.  Goal here is to read the actual bytes and figure out
which step in the JET → DAO → ADODB.Stream pipeline introduces it.

Outputs:

  - reports/gis_status_export_bytes_dump.json with the offending
    rows' raw cells, each as repr + hex.
  - Also re-saves the export at analysis/_status_gis_dump.tab so it
    can be inspected manually if needed.

Run:

  python analysis/probe_status_gis_export_bytes.py
"""
from __future__ import annotations

import json
import shutil
import sys
from pathlib import Path

# Make tests/ importable so we can reuse VbaSession + fixtures.
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))

from cbdb_driver.vba_session import VbaSession, make_fixture
from test_vba_matrix_all_forms import _all_fixtures, SRC

OUT_TAB = ROOT / "analysis" / "_status_gis_dump.tab"
OUT_JSON = ROOT / "reports" / "gis_status_export_bytes_dump.json"
WORK = ROOT / "analysis" / "_cmdgis_probe_copy.mdb"


def _seed(vba: VbaSession, fx) -> None:
    spec = fx.spec
    vba.open_form(spec.name)
    for ctl, val in fx.controls.items():
        try:
            vba.set_control(spec.name, ctl, val)
        except Exception as e:
            print(f"  warn setting {ctl}={val!r}: {e}")
    if fx.picker_ids and spec.picker_table:
        vba.set_picker_codes(spec.picker_table, fx.picker_ids,
                             column=spec.picker_column)
    try:
        vba.set_control(spec.name, "GISFrame", 1)
    except Exception:
        pass


def main() -> int:
    fx = next((f for f in _all_fixtures()
                if f.name == "status_40_unfiltered"), None)
    if fx is None:
        print("could not locate status_40_unfiltered fixture")
        return 1

    print(f"running CmdGIS for {fx.name} → {OUT_TAB}")
    sess_iter = make_fixture(SRC, WORK)
    vba = next(sess_iter)
    try:
        spec = fx.spec
        vba.patch_filedialog(spec.name)
        _seed(vba, fx)
        if OUT_TAB.exists():
            OUT_TAB.unlink()
        vba.set_form_tag(spec.name, f"{spec.cmd_name},CmdGIS",
                          str(OUT_TAB))
        n = vba.click_via_timer(
            spec.name, ctl=spec.cmd_name,
            result_table=spec.result_table, timeout=180,
        )
        print(f"  {spec.cmd_name} produced {n} scratch rows")
    finally:
        try:
            next(sess_iter)
        except StopIteration:
            pass

    if not OUT_TAB.exists():
        print(f"export did not appear at {OUT_TAB}")
        return 1
    raw = OUT_TAB.read_bytes()
    print(f"  wrote {len(raw)} bytes to {OUT_TAB}")
    text = raw.decode("utf-8", errors="replace").lstrip("﻿")
    lines = [ln for ln in text.replace("\r\n", "\n").split("\n")
             if ln.strip()]
    header = lines[0]
    sep = "\t"
    n_header = len(header.split(sep))
    print(f"  header has {n_header} cells: "
          f"{header.split(sep)!r}")

    # Find every row whose cell count != header.
    bad = []
    for i, ln in enumerate(lines[1:], start=1):
        cells = ln.split(sep)
        if len(cells) != n_header:
            bad.append((i, cells, ln))
    print(f"  bad rows: {len(bad)} / {len(lines) - 1}")

    # For each bad row: cell-by-cell raw + hex.
    bad_dump = []
    for i, cells, ln in bad[:25]:
        cell_dumps = []
        for j, c in enumerate(cells):
            cell_dumps.append({
                "cell_index": j,
                "repr": repr(c)[:200],
                "len_chars": len(c),
                "utf16le_hex": c.encode("utf-16-le").hex(" "),
                "utf8_hex": c.encode("utf-8", errors="replace").hex(" "),
            })
        bad_dump.append({
            "row_index_1based": i,
            "cell_count": len(cells),
            "raw_line_repr": repr(ln)[:400],
            "cells": cell_dumps,
        })

    out = {
        "fixture": fx.name,
        "n_scratch_rows": n,
        "n_data_rows": len(lines) - 1,
        "n_header_cells": n_header,
        "n_bad_rows": len(bad),
        "bad_rows_sample": bad_dump,
        "header_repr": repr(header),
    }
    OUT_JSON.write_text(
        json.dumps(out, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    print(f"wrote {OUT_JSON}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
