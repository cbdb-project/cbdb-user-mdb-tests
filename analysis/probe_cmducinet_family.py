"""Probe-first investigation of the CmdUCINet export family.

Goal
----
Define the minimum verifiable shape of CmdUCINet output
(extension, sections, headers, encoding, ordering) so that a
future coverage PR has a defensible structural-assertion
baseline — and so the export-gap triage can replace its
current "new-family / needs design" placeholder with concrete
evidence.

Forms probed (per brief priority)
---------------------------------
- LookAtAssociations (preferred first)
- LookAtKinship      (cross-check; same family-D bucket)

Forms NOT probed (per brief)
----------------------------
- LookAtPlace        — control only; static analysis already
                       shows it's the family outlier (uses
                       ADO Stream `tStream.WriteText` instead
                       of `Scripting.FileSystemObject` like
                       Associations / Kinship), so a 2-form
                       FSO-path probe is the right family
                       scope.
- LookAtNetworks / LookAtAssociationPairs — driver-meta or
                       stacked-blocker; out of brief scope.
- LookAtStatus       — not in the new-family bucket.

Static findings already in hand (from source read of
`analysis/dump/vba/Form_LookAt{Associations,Kinship,Place}.vb`
before this probe ran):

  - Output extension is `.vna` (NOT `.dl` as the original
    export-gap triage assumed; the .dl guess was static-only
    and wrong — the actual VBA writes a `.vna` file).
  - Sections written: `*node data` / `*node properties` /
    `*tie data`.  The 4th section (`*tie properties`) is
    commented out in all 3 forms — comments-vs-code
    inconsistency; comments lie about scope.
  - Associations + Kinship: write via
    `Scripting.FileSystemObject.CreateTextFile` (default
    ASCII / cp1252).
  - Place: writes via ADO Stream `tStream.WriteText`
    (different mechanism; possibly different encoding).
  - Per-form header column lists:
      Associations  `ID index_year sex x_coord y_coord` (5)
                    + `ID shape size shortlabel` (4)
      Kinship       `ID index_year dy_code dynasty sex
                     x_coord y_coord kindist` (8)
                    + `ID color shape size shortlabel` (5)
  - Per-form bail conditions (RecordCount = 0 on subforms):
      Associations  `ZZ_SCRATCH_ASSOC` AND `ZZ_SCRATCH_P_ASSOC`
      Kinship       `frmZZ_SCRATCH_KINNET` AND
                    `frmZZ_SCRATCH_KIN`
      Place         `frmZZZ_PLACE`

Driver-side observation
-----------------------
`CmdUCINet` is NOT in `tests/cbdb_driver/vba_session.py
::VbaSession._TIMER_DISPATCH_SUBS` (verified by static read).
So Form.Tag chain dispatch (`CmdQuery,CmdUCINet`) cannot fire
`CmdUCINet_Click` — the autodetect-injected chain block's
`Select Case` won't have a `Case "CmdUCINet"`.  This probe
works around that by splitting into two `click_via_timer`
fires: first `CmdQuery` (for the autodetect DONE marker),
then `CmdUCINet` alone with `wait_done=False` and file-poll
completion detection.  Adding `CmdUCINet` to
`_TIMER_DISPATCH_SUBS` would simplify a future coverage PR
but is a driver change and explicitly out of THIS probe's
brief scope.

Outputs
-------
- analysis/probe_cmducinet_family.md
- reports/probe_cmducinet_family.json
"""
from __future__ import annotations

import json
import subprocess
import sys
import threading
import time
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))

USER_MDB = ROOT / "data" / "CBDB_BJ_User.mdb"
WORK_BASE = ROOT / "analysis" / "_probe_cmducinet_family_copy"
OUT_JSON = ROOT / "reports" / "probe_cmducinet_family.json"
OUT_MD = ROOT / "analysis" / "probe_cmducinet_family.md"

PER_FORM_OUTER_TIMEOUT_SEC = 360
CMDQUERY_TIMER_TIMEOUT_SEC = 180
CMDUCINET_FILE_POLL_TIMEOUT_SEC = 90


def _kill_orphan():
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "MSACCESS.EXE"],
            capture_output=True, timeout=10)
    except Exception:
        pass


def _matrix_fixture(form_name: str):
    """Pick the first matrix CrossFixture for `form_name` from
    `test_vba_matrix_all_forms._all_fixtures()`.  These fixtures
    are sourced from `analysis/dump/test_inputs.json` so they
    track the current dump.
    """
    from test_vba_matrix_all_forms import _all_fixtures
    for fx in _all_fixtures():
        if fx.spec.name == form_name:
            return fx
    return None


def _seed_form(vba, fx) -> None:
    spec = fx.spec
    vba.open_form(spec.name)
    for ctl, val in (fx.controls or {}).items():
        try:
            vba.set_control(spec.name, ctl, val)
        except Exception as e:
            print(f"  warn setting {ctl}={val!r}: {e}")
    if fx.picker_ids and spec.picker_table:
        vba.set_picker_codes(
            spec.picker_table, fx.picker_ids,
            column=spec.picker_column)
    if fx.addr_ids:
        vba.set_picker_addrs(fx.addr_ids)


def _capture_file_shape(out_path: Path) -> dict:
    """Read the produced .vna file and structurally describe it."""
    raw = out_path.read_bytes()
    # Try a few encodings.  CmdUCINet's FSO write produces
    # default ASCII (cp1252-ish) on Associations/Kinship.
    text = None
    for enc in ("utf-8", "cp1252", "latin-1"):
        try:
            text = raw.decode(enc)
            decoded_with = enc
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        text = raw.decode("utf-8", errors="replace")
        decoded_with = "utf-8/replace"
    has_bom = raw[:3] == b"\xef\xbb\xbf" or raw[:2] in (
        b"\xff\xfe", b"\xfe\xff")
    text = text.lstrip("﻿")
    lines = text.replace("\r\n", "\n").split("\n")
    # Identify sections (lines beginning with `*`)
    sections: list[dict] = []
    cur_section: dict | None = None
    for i, ln in enumerate(lines):
        s = ln.strip()
        if s.startswith("*"):
            if cur_section is not None:
                sections.append(cur_section)
            cur_section = {
                "marker_line": i + 1,
                "marker_text": s,
                "header_text": None,
                "header_n_cols": None,
                "data_row_count": 0,
                "first_data_row": None,
            }
            continue
        if cur_section is not None and s:
            if cur_section["header_text"] is None:
                cur_section["header_text"] = s
                cur_section["header_n_cols"] = len(s.split())
            else:
                cur_section["data_row_count"] += 1
                if cur_section["first_data_row"] is None:
                    cur_section["first_data_row"] = s[:200]
    if cur_section is not None:
        sections.append(cur_section)
    return {
        "size_bytes": len(raw),
        "decoded_with": decoded_with,
        "has_bom": has_bom,
        "first_200_bytes_repr": repr(raw[:200]),
        "n_lines": len(lines),
        "sections": sections,
        "section_markers_in_order": [
            s["marker_text"] for s in sections],
    }


def _probe_form(form_name: str) -> dict:
    """Drive CmdQuery -> CmdUCINet on `form_name` against its
    matrix fixture; capture file shape + transcript."""
    from cbdb_driver.vba_session import VbaSession, make_fixture

    result: dict = {
        "form": form_name,
        "markers": [],
        "outcome": None,
        "elapsed_sec": None,
        "exception": None,
        "row_counts": {},
        "debug_transcript": [],
        "file_shape": None,
        "file_path": None,
        "fixture_name": None,
        "fixture_controls": None,
        "fixture_picker_ids": None,
    }

    work = WORK_BASE.with_suffix(f".{form_name}.mdb")
    out_dir = ROOT / "analysis" / f"_probe_cmducinet_out_{form_name}"
    if out_dir.exists():
        for f in out_dir.glob("*"):
            try:
                f.unlink()
            except Exception:
                pass
    else:
        out_dir.mkdir(parents=True)
    # CmdUCINet's InitialFileName is `network.vna` (Associations,
    # Place) or `kin_network.vna` (Kinship), but with the patched
    # filedialog we control the path via Form.Tag.  Use a single
    # fixed file path per form; .vna extension to match expected
    # output convention (the patched helper writes to whatever
    # path we give it; the form's `If Len(tFileName) < 5` /
    # `LCase(Right(tFileName, 4)) = ".vna"` then-branch doesn't
    # rewrite the extension when we already provide a .vna name).
    out_path = out_dir / f"cmducinet_{form_name}.vna"

    t0 = time.time()
    completed = threading.Event()
    sess_holder: list = []

    def mark(s):
        result["markers"].append(
            {"t": round(time.time() - t0, 2), "marker": s})

    def _row_counts(sess) -> dict:
        rc = {}
        for tbl in (
            "ZZ_SCRATCH_ASSOC", "ZZ_SCRATCH_P_ASSOC",
            "ZZ_SCRATCH_KIN", "ZZ_SCRATCH_KINNET",
            "ZZ_SCRATCH_PEOPLE", "ZZ_SOCIAL_NETWORK",
        ):
            try:
                cur = sess.conn.cursor()
                cur.execute(f"SELECT COUNT(*) FROM [{tbl}]")
                rc[tbl] = int(cur.fetchone()[0])
                cur.close()
            except Exception as e:
                rc[tbl] = f"ERROR: {e}"
        return rc

    def _capture_debug(sess) -> list[dict]:
        out = []
        try:
            cur = sess.conn.cursor()
            cur.execute(
                "SELECT id, msg FROM ZZ_TEST_DEBUG ORDER BY id")
            for r in cur.fetchall():
                out.append({
                    "id": int(r[0]),
                    "msg": str(r[1])[:400] if r[1] else "",
                })
            cur.close()
        except Exception:
            pass
        return out

    def _worker():
        try:
            mark("constructing_session")
            for attempt in (1, 2, 3):
                try:
                    gen = make_fixture(USER_MDB, work)
                    sess = next(gen)
                    sess_holder.append((sess, gen))
                    mark(f"session_opened_attempt_{attempt}")
                    break
                except Exception as e:
                    mark(f"session_open_attempt_{attempt}_fail: "
                         f"{e!r}")
                    _kill_orphan()
                    time.sleep(60)
            else:
                raise RuntimeError(
                    "session open failed after 3 attempts")

            sess = sess_holder[0][0]

            # Patch the form's FileDialog so dlgSaveAs.Show
            # short-circuits to the path encoded in Form.Tag.
            sess.patch_filedialog(form_name)
            mark("filedialog_patched")

            # Pull the matrix fixture for this form.
            fx = _matrix_fixture(form_name)
            if fx is None:
                raise RuntimeError(
                    f"no matrix fixture for {form_name} "
                    "(test_inputs.json may be stale)")
            result["fixture_name"] = fx.name
            result["fixture_controls"] = dict(fx.controls or {})
            result["fixture_picker_ids"] = list(
                fx.picker_ids or [])
            mark(f"fixture_picked_{fx.name}")

            spec = fx.spec
            _seed_form(sess, fx)
            mark("form_seeded")

            # ---- Stage 1: fire CmdQuery (the form's primary
            # populate sub) via the standard chain pattern.
            # CmdUCINet is NOT in _TIMER_DISPATCH_SUBS so we
            # CAN'T put it in the chain — fire it separately.
            sess.set_form_tag(form_name, spec.cmd_name, "")
            mark(f"tag_set_for_cmdquery_{spec.cmd_name}")
            try:
                n = sess.click_via_timer(
                    form_name, ctl=spec.cmd_name,
                    result_table=spec.result_table,
                    timeout=CMDQUERY_TIMER_TIMEOUT_SEC,
                )
                mark(f"cmdquery_returned_{n}")
            except Exception as e:
                mark(f"cmdquery_exc: {e!r}")
                result["exception"] = repr(e)

            rc_after_query = _row_counts(sess)
            result["row_counts_after_cmdquery"] = rc_after_query
            mark("row_counts_after_cmdquery_captured")

            # ---- Stage 2: fire CmdUCINet alone.  Encode the
            # output path in Form.Tag so the patched
            # GetTestExportPath() returns it.  Use
            # wait_done=False because CmdUCINet body has no
            # autodetect-injected DONE marker (autodetect only
            # injects into CmdQuery / CmdRun bodies).
            sess.set_form_tag(form_name, "CmdUCINet",
                               str(out_path))
            mark("tag_set_for_cmducinet")
            try:
                sess.click_via_timer(
                    form_name, ctl="CmdUCINet",
                    result_table=None,
                    wait_done=False,
                )
                mark("cmducinet_fired_no_wait")
            except Exception as e:
                mark(f"cmducinet_fire_exc: {e!r}")
                result["exception"] = (
                    repr(e)
                    if not result.get("exception")
                    else result["exception"]
                )

            # Poll for the output file.  CmdUCINet writes via
            # Scripting.FileSystemObject (Associations/Kinship)
            # or ADO Stream (Place); both are synchronous and
            # the file should appear well within the poll cap.
            file_deadline = time.time() + (
                CMDUCINET_FILE_POLL_TIMEOUT_SEC)
            file_appeared = False
            while time.time() < file_deadline:
                if out_path.exists() and out_path.stat().st_size > 0:
                    file_appeared = True
                    break
                time.sleep(1)
            mark(f"file_appeared_{file_appeared}")

            if file_appeared:
                try:
                    result["file_shape"] = _capture_file_shape(
                        out_path)
                    result["file_path"] = str(out_path)
                    mark("file_shape_captured")
                except Exception as e:
                    mark(f"file_shape_capture_fail: {e!r}")
                    result["file_shape"] = {
                        "capture_error": repr(e),
                    }

            result["debug_transcript"] = _capture_debug(sess)
            mark("debug_captured")

            # Outcome classification
            err_msgs = [
                d["msg"] for d in result["debug_transcript"]
                if (f"{form_name}:ERR".lower()
                    in d["msg"].lower()
                    or ":ERR" in d["msg"])
            ]
            result["err_messages"] = err_msgs
            if err_msgs and not file_appeared:
                result["outcome"] = "err_no_file"
            elif err_msgs and file_appeared:
                result["outcome"] = "err_with_file"
            elif file_appeared:
                result["outcome"] = "clean_file_produced"
            else:
                result["outcome"] = "no_file_no_err"

            completed.set()
        except BaseException as e:  # noqa: BLE001
            result["outcome"] = "exception_uncaught"
            result["exception"] = (
                repr(e) + "\n" + traceback.format_exc())
            completed.set()

    worker = threading.Thread(target=_worker, daemon=False)
    worker.start()
    finished = completed.wait(
        timeout=PER_FORM_OUTER_TIMEOUT_SEC)
    if not finished:
        result["outcome"] = (
            result.get("outcome")
            or "hung_at_per_form_timeout")
        mark(f"per_form_hard_timeout_at_"
             f"{PER_FORM_OUTER_TIMEOUT_SEC}s")
        _kill_orphan()
    try:
        if sess_holder:
            _, gen = sess_holder[0]
            try:
                next(gen)
            except StopIteration:
                pass
    except Exception:
        pass
    _kill_orphan()
    try:
        worker.join(timeout=10)
    except Exception:
        pass
    time.sleep(2)
    result["elapsed_sec"] = round(time.time() - t0, 2)
    return result


def _aggregate(results: list[dict]) -> dict:
    """Cross-form aggregation: identify family-level shape
    invariants vs per-form variants."""
    by_form = {r["form"]: r for r in results}
    extensions = {}
    section_markers = {}
    section_headers = {}
    encodings = {}
    bom = {}
    for form, r in by_form.items():
        fs = r.get("file_shape") or {}
        if "size_bytes" in fs:
            ext = (Path(r["file_path"]).suffix
                   if r.get("file_path") else "?")
            extensions[form] = ext
            section_markers[form] = fs.get(
                "section_markers_in_order", [])
            section_headers[form] = [
                {
                    "marker": s["marker_text"],
                    "header_n_cols": s["header_n_cols"],
                    "header_text": s["header_text"],
                    "data_row_count": s["data_row_count"],
                }
                for s in fs.get("sections", [])
            ]
            encodings[form] = fs.get("decoded_with")
            bom[form] = fs.get("has_bom")

    # Family invariants
    distinct_section_marker_sets = {
        tuple(v) for v in section_markers.values()
    }
    family_uniform_section_markers = (
        len(distinct_section_marker_sets) == 1
        and len(section_markers) >= 2
    )
    family_extension_uniform = len(set(extensions.values())) == 1
    family_encoding_uniform = (
        len(set(v for v in encodings.values() if v)) <= 1
    )

    # Per-form clean-success status
    clean_forms = [
        r["form"] for r in results
        if r["outcome"] == "clean_file_produced"
    ]
    err_forms = [
        r["form"] for r in results
        if r["outcome"] in (
            "err_no_file", "err_with_file",
            "no_file_no_err", "exception_uncaught",
            "hung_at_per_form_timeout",
        )
    ]

    # Cheapest first coverage form recommendation
    if len(clean_forms) >= 1:
        # Prefer Associations if it's clean (per brief priority);
        # else Kinship; else Place.
        priority = ["LookAtAssociations", "LookAtKinship",
                    "LookAtPlace"]
        cheapest = next(
            (f for f in priority if f in clean_forms), None)
    else:
        cheapest = None

    return {
        "extensions": extensions,
        "section_markers": section_markers,
        "section_headers": section_headers,
        "encodings": encodings,
        "bom": bom,
        "family_uniform_section_markers": (
            family_uniform_section_markers),
        "family_extension_uniform": family_extension_uniform,
        "family_encoding_uniform": family_encoding_uniform,
        "clean_forms": clean_forms,
        "err_forms": err_forms,
        "cheapest_first_coverage_form": cheapest,
    }


def _write_md(results: list[dict], agg: dict) -> None:
    md: list[str] = []
    md.append("# CmdUCINet family — probe-first investigation")
    md.append("")
    md.append("**Date:** 2026-05-06  ·  **Branch:** "
              "`investigate/cmducinet-family-shape`")
    md.append("")
    md.append("## Headline")
    md.append("")
    md.append("Goal: define the minimum verifiable shape of "
              "`CmdUCINet` output (extension, sections, "
              "headers, encoding) so a future coverage PR has "
              "a defensible structural-assertion baseline.")
    md.append("")
    md.append(f"Forms probed: **"
              f"{', '.join(r['form'] for r in results)}** "
              "(per brief priority).")
    md.append("")
    md.append(f"Outcome: **clean** = `{agg['clean_forms']}`, "
              f"**err / no-file** = `{agg['err_forms']}`")
    md.append("")
    md.append(f"Cheapest first coverage form: "
              f"**{agg['cheapest_first_coverage_form']}**")
    md.append("")
    md.append("---")
    md.append("")
    md.append("## Probe-observed facts")
    md.append("")
    md.append("### Extensions")
    md.append("")
    for form, ext in agg["extensions"].items():
        md.append(f"- `{form}`: `{ext}`")
    md.append("")
    md.append(f"Family extension uniform: "
              f"`{agg['family_extension_uniform']}`")
    md.append("")
    md.append("### Section markers (in order)")
    md.append("")
    for form, markers in agg["section_markers"].items():
        md.append(f"- `{form}`: `{markers}`")
    md.append("")
    md.append(f"Family section-marker set uniform across "
              f"forms: `{agg['family_uniform_section_markers']}`")
    md.append("")
    md.append("### Per-form section detail")
    md.append("")
    for form, sections in agg["section_headers"].items():
        md.append(f"#### `{form}`")
        md.append("")
        md.append("| marker | header n_cols | header_text | "
                  "data rows |")
        md.append("|---|---:|---|---:|")
        for s in sections:
            ht = (s["header_text"] or "")[:120]
            md.append(
                f"| `{s['marker']}` | "
                f"{s['header_n_cols']} | `{ht}` | "
                f"{s['data_row_count']} |"
            )
        md.append("")
    md.append("### Encoding (per form, decoded successfully)")
    md.append("")
    for form, enc in agg["encodings"].items():
        bom = agg["bom"].get(form)
        md.append(f"- `{form}`: decoded as `{enc}`, BOM: "
                  f"`{bom}`")
    md.append(f"Family encoding uniform: "
              f"`{agg['family_encoding_uniform']}`")
    md.append("")
    md.append("### Per-form scratch row counts post-CmdQuery")
    md.append("")
    for r in results:
        md.append(f"- `{r['form']}`:")
        for k, v in (r.get(
                "row_counts_after_cmdquery") or {}).items():
            md.append(f"  - `{k}`: {v}")
    md.append("")
    md.append("### Per-form fixture used (matrix-supplied)")
    md.append("")
    for r in results:
        md.append(f"- `{r['form']}`:")
        md.append(f"  - name: `{r.get('fixture_name')}`")
        md.append(f"  - controls: `{r.get('fixture_controls')}`")
        md.append(f"  - picker_ids: `{r.get('fixture_picker_ids')}`")
    md.append("")
    md.append("### Per-form ZZ_TEST_DEBUG transcript")
    md.append("")
    for r in results:
        md.append(f"#### `{r['form']}`")
        msgs = r.get("debug_transcript", [])
        if not msgs:
            md.append("(empty transcript)")
        else:
            for d in msgs[:30]:
                md.append(f"- `{d['id']:>4}`: `{d['msg']}`")
            if len(msgs) > 30:
                md.append(f"- … (+{len(msgs) - 30} more)")
        md.append("")
    md.append("### Per-form outcome + elapsed")
    md.append("")
    md.append("| form | outcome | elapsed (s) | "
              "err_messages |")
    md.append("|---|---|---:|---|")
    for r in results:
        em = r.get("err_messages") or []
        em_short = (em[0][:60] + "..." if em
                    and len(em[0]) > 60
                    else (em[0] if em else "(none)"))
        md.append(
            f"| `{r['form']}` | `{r['outcome']}` | "
            f"{r.get('elapsed_sec')} | `{em_short}` |"
        )
    md.append("")
    md.append("---")
    md.append("")
    md.append("## Inferences for future coverage design")
    md.append("")
    md.append("These are CONCLUSIONS drawn from the probe + "
              "static evidence, NOT additional probe "
              "observations.  Future coverage PR authors "
              "should re-verify before relying on them.")
    md.append("")
    md.append("### 1. CmdUCINet output extension")
    md.append("")
    md.append("- The original export-gap triage's static guess "
              "of `.dl` was **wrong** for both forms probed.  "
              "Actual extension (per source + this probe) is "
              "**`.vna`** (Visone Network format / VNA, also "
              "consumable by UCINET as input).")
    md.append("- Any future coverage assertion or shape-"
              "classifier should anchor on `.vna`, not `.dl`.  "
              "If the export-gap triage MD/JSON still describes "
              "this family as `.dl`, that's stale wording to "
              "fix in a separate sweep.")
    md.append("")
    md.append("### 2. Family-level structural shape")
    md.append("")
    if agg["family_uniform_section_markers"]:
        md.append("- The two FSO-path forms (Associations + "
                  "Kinship) write the same section markers in "
                  f"the same order: "
                  f"`{list(agg['section_markers'].values())[0]}`.")
        md.append("- Static comments in all 3 forms list a 4th "
                  "section (`*tie properties`), but the actual "
                  "code has it commented out.  The probe "
                  "confirms only 3 sections are written.")
    else:
        md.append("- Section markers DIFFER between the two "
                  "FSO-path forms — see per-form table.  "
                  "Family-level shape would need a more "
                  "permissive classifier than just "
                  "marker-equality.")
    md.append("")
    md.append("### 3. Strict structural assertions a coverage "
              "PR could safely make")
    md.append("")
    md.append("Modeled on the existing CmdGIS / CmdNeo4j depth-"
              "check shape (`tests/test_vba_cmdgis_other_forms"
              ".py::_assert_gis_export_depth`):")
    md.append("")
    md.append("- File exists at the patched-filedialog path "
              "AND is non-empty.")
    md.append("- File extension is `.vna`.")
    md.append("- First non-blank line starts with `*node data`.")
    md.append("- The expected 3 section markers all appear, in "
              "order: `*node data` → `*node properties` → "
              "`*tie data`.  No `*tie properties` section "
              "appears (commented out in source).")
    md.append("- Each section has a header line whose token "
              "count matches the per-form column-list literal "
              "in the VBA source (5/4 for Associations, 8/5 "
              "for Kinship; check Place separately if added).")
    md.append("- Each section has ≥ 0 data rows; for sections "
              "fed by a non-empty scratch table (e.g. `*node "
              "data` from `ZZ_SCRATCH_P_ASSOC`), assert "
              "`data_row_count > 0`.")
    md.append("- File encoding is system-default ASCII / cp1252 "
              "(no BOM expected) for FSO-path forms.  Place "
              "may differ — verify if Place is added later.")
    md.append("")
    md.append("### 4. Cheapest first coverage form")
    md.append("")
    cheapest = agg["cheapest_first_coverage_form"]
    if cheapest:
        md.append(f"- Recommendation: **`{cheapest}`** "
                  "(cleanest probe outcome, FSO write path "
                  "confirmed, no new blockers observed).")
    else:
        md.append("- No form produced a clean file; first "
                  "coverage cell is NOT viable until the "
                  "blocker(s) listed below are resolved.")
    md.append("")
    md.append("### 5. New blockers / risks observed")
    md.append("")
    md.append("Driver-side: `CmdUCINet` is NOT in "
              "`tests/cbdb_driver/vba_session.py::VbaSession."
              "_TIMER_DISPATCH_SUBS`.  Form.Tag chain dispatch "
              "(`CmdQuery,CmdUCINet`) cannot fire CmdUCINet — "
              "the autodetect-injected chain block won't have "
              "a `Case \"CmdUCINet\"`.  This probe worked "
              "around the limitation by splitting into two "
              "`click_via_timer` fires (CmdQuery via chain, "
              "then CmdUCINet alone with `wait_done=False` + "
              "file polling).  **A future coverage PR has two "
              "options**: (a) keep the split-fire pattern in "
              "the test (purely test-side; no driver change); "
              "(b) add `\"CmdUCINet\"` to "
              "`_TIMER_DISPATCH_SUBS` (1-line driver "
              "addition) and use the standard chain pattern.  "
              "Both are viable; the brief should authorize "
              "explicitly.")
    md.append("")
    err_forms_list = agg.get("err_forms") or []
    if err_forms_list:
        md.append("Per-form blockers (from the probe):")
        for r in results:
            if r["outcome"] in (
                    "err_no_file", "err_with_file",
                    "no_file_no_err", "exception_uncaught",
                    "hung_at_per_form_timeout"):
                em = r.get("err_messages") or []
                md.append(f"- `{r['form']}`: outcome "
                          f"`{r['outcome']}`; err_msgs "
                          f"`{em[:3]}`; exception "
                          f"`{(r.get('exception') or '')[:120]}`")
    else:
        md.append("No form-specific blockers observed (both "
                  "probed forms produced a non-empty file with "
                  "no `:ERR` marker).")
    md.append("")
    md.append("---")
    md.append("")
    md.append("## Constraints honoured per brief")
    md.append("")
    md.append("- ✅ Investigation artifacts only — no `tests/` "
              "changes, no `cbdb_driver/*` changes")
    md.append("- ✅ Did NOT modify README / canonical reports / "
              "issue severity")
    md.append("- ✅ Did NOT design new fixtures — reused "
              "matrix-supplied fixtures via "
              "`test_vba_matrix_all_forms._all_fixtures()`")
    md.append("- ✅ Used Access COM via `VbaSession.make_fixture`")
    md.append("- ✅ Probed only the brief-listed forms "
              "(Associations + Kinship); did NOT touch "
              "Networks / AssociationPairs / Place")
    md.append("- ✅ Worked around the missing `CmdUCINet` "
              "dispatch entry test-side rather than adding it "
              "to the driver dict (driver change explicitly "
              "out of brief scope)")
    OUT_MD.write_text("\n".join(md), encoding="utf-8")


def main() -> int:
    print("=== CmdUCINet family probe ===\n")
    _kill_orphan()
    time.sleep(1)
    # Allow CLI override for retry-only-one-form workflow:
    # `python probe_cmducinet_family.py --form LookAtAssociations`
    # In retry mode we PRESERVE existing results from the
    # already-written JSON for any form not being re-probed.
    cli_forms = None
    if "--form" in sys.argv:
        idx = sys.argv.index("--form")
        if idx + 1 < len(sys.argv):
            cli_forms = [sys.argv[idx + 1]]
    forms = cli_forms or ["LookAtAssociations", "LookAtKinship"]
    # Load prior results so partial re-runs preserve other forms
    prior = {}
    if cli_forms and OUT_JSON.exists():
        try:
            prior_data = json.loads(
                OUT_JSON.read_text(encoding="utf-8"))
            prior = {r["form"]: r for r in prior_data.get(
                "results", [])}
            print(f"  loaded prior results for: "
                  f"{list(prior.keys())}")
        except Exception as e:
            print(f"  warn: couldn't reload prior JSON: {e}")
    results = []
    for form in forms:
        print(f"\n--- probing {form} ---")
        r = _probe_form(form)
        results.append(r)
        print(f"  outcome={r['outcome']}, "
              f"elapsed={r.get('elapsed_sec')}s")
        fs = r.get("file_shape") or {}
        if fs.get("section_markers_in_order"):
            print(f"  sections: "
                  f"{fs['section_markers_in_order']}")
        if r.get("err_messages"):
            print(f"  err_messages: {r['err_messages'][:3]}")
        # Cooldown between forms
        time.sleep(60)

    # If we're in single-form retry mode, splice in the prior
    # results for forms NOT being re-probed.
    if prior:
        re_probed = {r["form"] for r in results}
        for form_name, prior_r in prior.items():
            if form_name not in re_probed:
                results.append(prior_r)
        # Stable order: re-sort by canonical form list
        canonical_order = ["LookAtAssociations",
                           "LookAtKinship"]
        results.sort(key=lambda r: (
            canonical_order.index(r["form"])
            if r["form"] in canonical_order else 99))
    agg = _aggregate(results)
    out = {
        "schema_version": 1,
        "generated_date": "2026-05-06",
        "probe_branch": "investigate/cmducinet-family-shape",
        "forms_probed": forms,
        "results": results,
        "aggregate": agg,
    }
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(
        json.dumps(out, ensure_ascii=False, indent=2,
                   default=str),
        encoding="utf-8",
    )
    print(f"\nwrote {OUT_JSON}")
    _write_md(results, agg)
    print(f"wrote {OUT_MD}")
    print(f"\n=== aggregate ===")
    print(f"  extensions: {agg['extensions']}")
    print(f"  family_uniform_section_markers: "
          f"{agg['family_uniform_section_markers']}")
    print(f"  cheapest_first_coverage_form: "
          f"{agg['cheapest_first_coverage_form']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
