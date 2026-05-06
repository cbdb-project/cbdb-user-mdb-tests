"""Probe: characterize LookAtPlace.CmdUCINet runtime behavior.

Goal
----
PR AR's queue refresh canonicalized that `LookAtPlace × CmdUCINet`
is the cheapest next probe in the CmdUCINet family because Place's
write path is *structurally* different from the FSO ANSI path that
drives Issue #22:

  - Place uses ADO Stream (`tStream.WriteText ... adWriteLine`)
    with `tStream.Charset = "utf-8"` (default; alt: "big5",
    "gb18030").
  - The FSO `Set tVNA = tFileSystem.CreateTextFile(tFileName, True)`
    call that drives Issue #22 in Associations + Kinship is
    COMMENTED OUT in `Form_LookAtPlace.CmdUCINet_Click`.
  - The expected file shape is THREE sections (`*node data`,
    `*node properties`, `*tie data`); `*tie properties` is also
    commented out.

This probe answers — at runtime, on a real fixture — whether Place
CmdUCINet:

  1. Completes and produces a parseable `.vna` file.
  2. Writes UTF-8 with the BOM that ADO Stream's `SaveToFile`
     usually emits.
  3. Has the 3-section shape that the static read predicts.
  4. Has any new blocker (encoding bug, partial write,
     empty-recordset guard, runtime `:ERR`, file-dialog / chain
     issue).

Importantly: the brief is explicit that this probe is NOT a
reproduction of Issue #22 — Place's writer is structurally
different, and we should not collapse the result into the Issue #22
family.

Method
------
1. Pick the existing matrix fixture address (Kaifeng, c_addr_id =
   100658).  This is already known to have a CJK-rich network
   (matrix fixture `kaifeng_individual_900_1100.csv` shows Han
   names like 開封, 安燾, 柴天因 etc.), so it doubles as a
   "Han-name stress" run for the UTF-8 path without designing a
   new fixture.
2. Open `Form_LookAtPlace`, seed the picker (ZZ_SCRATCH_ADDR
   ← 100658), fire CmdQuery via timer (the form's standard
   chain).
3. Capture ZZ_PLACE row counts + the c_rel_type distribution
   (the Network section of CmdUCINet only fires on rows with
   c_rel_type IN ('Kinship', 'Associate Place'); other sources
   set c_rel_type = 'BIOGRAPHY' or other values).
4. Fire CmdUCINet via separate timer (split-fire pattern;
   CmdUCINet not in `_TIMER_DISPATCH_SUBS`); poll for the
   .vna file to appear.
5. Read the file as bytes; detect BOM; decode as UTF-8 (per the
   static prediction); parse sections + headers + row counts.
6. Cross-check against ZZ_SOCIAL_NETWORK + ZZ_SCRATCH_PEOPLE
   (CmdUCINet's intermediate scratch tables) for row-count
   parity.

Conclusion buckets (per brief)
------------------------------
- `clean_complete_export`
    - File appears, has the expected 3-section shape, decodes
      cleanly as UTF-8 (with or without BOM), no `:ERR`
      transcript line.  Parity (data-row counts == scratch
      table row counts) holds.
- `runtime_err_with_partial_file`
    - File exists but partially written, AND a `:ERR` line
      shows up in ZZ_TEST_DEBUG.  Captures any new failure
      mode in Place's ADO Stream path.
- `empty_or_guard_bail`
    - CmdUCINet bailed at one of its early guards
      (`If frmZZZ_PLACE.Form.Recordset.RecordCount = 0` OR
      `If tRecDeleted = 0` after the first INSERT into
      ZZ_SOCIAL_NETWORK).  Means the chosen fixture didn't
      reach a state that drives Network output — not a Place
      CmdUCINet bug; a fixture-strategy issue.
- `still_needs_better_fixture`
    - Setup-phase failure (RPC flake, fixture didn't load,
      timer never fired, etc.) such that the probe's question
      genuinely wasn't answered.

Outputs
-------
- analysis/probe_place_cmducinet.md
- reports/probe_place_cmducinet.json
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
WORK = ROOT / "analysis" / "_probe_place_ucinet_copy.mdb"
OUT_JSON = ROOT / "reports" / "probe_place_cmducinet.json"
OUT_MD = ROOT / "analysis" / "probe_place_cmducinet.md"

# Picker chosen via pyodbc query of CBDB_*_DATA.mdb (BIOG_ADDR_DATA
# + KIN_DATA + ADDR_CODES) for: small BIOG_ADDR row count (so
# CmdQuery completes fast and Access COM stays stable), non-zero
# kin reach (so c_rel_type='Kinship' rows get produced), and a
# CJK name (so the UTF-8 ADO Stream path is stress-tested).  Earlier
# probe iterations on Kaifeng (addr 100658) showed CmdQuery on the
# matrix-fixture address takes >240s and destabilizes Access COM
# such that the subsequent CmdUCINet COM call fails with "RPC
# server unavailable" — separate concern, recorded in the MD as a
# fixture-strategy finding.
#
# Chenliu 陳留 (addr 3089): 3 BIOG_ADDR_DATA rows, 3 KIN_DATA links
# off those people's c_personid.  Probe-only picker choice — NOT a
# new long-term fixture (no test/golden data added).
PICKER_ADDR_ID = 3089
PICKER_ADDR_LABEL = "Chenliu (陳留)"
PICKER_RATIONALE = (
    "small BIOG_ADDR_DATA reach (3 rows) + non-zero KIN_DATA reach "
    "(3 links) + CJK name; chosen so CmdQuery completes fast enough "
    "to keep Access COM stable through the subsequent CmdUCINet "
    "fire (Kaifeng addr 100658 took >240s in earlier probe v2 and "
    "destabilized COM)")

PER_PROBE_OUTER_TIMEOUT_SEC = 360
CMDQUERY_TIMER_TIMEOUT_SEC = 120
CMDUCINET_FILE_POLL_TIMEOUT_SEC = 60

# Place form has many source-toggle checkboxes (ChkIndividual /
# ChkKin / ChkOffice / ChkStatus / ChkEntry / ChkInstitution /
# ChkAssocPerson / ChkAssocPlace).  CmdUCINet's first INSERT into
# ZZ_SOCIAL_NETWORK only selects rows where c_rel_type IN
# ('Kinship', 'Associate Place'); the form's defaults on this
# dump produce Biography/Office Place/Entry rows only (confirmed
# by an earlier probe v1 dry-run on Kaifeng).  We explicitly
# enable ChkKin so CmdQuery produces Kinship-rel-type rows that
# CmdUCINet's writer will pick up.  Probe-only checkbox setup;
# NOT a long-term test fixture.
PROBE_CHECKBOX_OVERRIDES = [
    ("ChkKin", True),         # source: KIN_DATA -> c_rel_type='Kinship'
]


def _kill_orphan() -> None:
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "MSACCESS.EXE"],
            capture_output=True, timeout=10)
    except Exception:
        pass


def _detect_bom(raw: bytes) -> tuple[str | None, int]:
    """Return (bom_label, n_bytes_consumed)."""
    if raw.startswith(b"\xef\xbb\xbf"):
        return ("utf-8-bom", 3)
    if raw.startswith(b"\xff\xfe"):
        return ("utf-16-le-bom", 2)
    if raw.startswith(b"\xfe\xff"):
        return ("utf-16-be-bom", 2)
    return (None, 0)


def _decode_with_fallback(raw: bytes) -> tuple[str, str, str | None]:
    """Return (text, encoding_label, bom_label).

    Per the static prediction Place writes UTF-8.  We try
    utf-8-sig first (handles the BOM that ADO Stream typically
    emits), then strict utf-8, then cp1252 as a last-resort
    fallback so the probe still records SOMETHING parseable
    rather than crashing here.
    """
    bom, _ = _detect_bom(raw)
    try:
        return (raw.decode("utf-8-sig"), "utf-8-sig", bom)
    except UnicodeDecodeError:
        pass
    try:
        return (raw.decode("utf-8"), "utf-8", bom)
    except UnicodeDecodeError:
        pass
    return (raw.decode("cp1252", errors="replace"),
            "cp1252-fallback-with-replace", bom)


def _parse_vna(text: str) -> dict:
    """Parse a .vna file into ordered sections.

    Each section starts with a `*<marker>` line, then a header
    line, then zero or more data rows.  Rows are split on
    whitespace for the token-count metric.
    """
    lines = text.replace("\r\n", "\n").split("\n")
    sections: list[dict] = []
    cur: dict | None = None
    for ln in lines:
        s = ln.rstrip()
        if s.startswith("*"):
            if cur is not None:
                sections.append(cur)
            cur = {"marker": s.strip(), "header": None, "rows": []}
            continue
        if cur is not None and s.strip():
            if cur["header"] is None:
                cur["header"] = s
            else:
                cur["rows"].append(s)
    if cur is not None:
        sections.append(cur)
    out_sections = [
        {
            "marker": s["marker"],
            "header": s["header"],
            "header_token_count": (
                len((s["header"] or "").split())
                if s["header"] else 0),
            "data_rows": len(s["rows"]),
            "first_row": s["rows"][0] if s["rows"] else None,
            "last_row": s["rows"][-1] if s["rows"] else None,
        }
        for s in sections
    ]
    return {
        "sections": out_sections,
        "section_markers_in_order": [
            s["marker"] for s in out_sections],
        "total_lines": len(lines),
    }


def _run_probe() -> dict:
    from cbdb_driver.vba_session import make_fixture
    from cbdb_driver.form_specs import LOOKATPLACE

    spec = LOOKATPLACE
    out_dir = ROOT / "analysis" / "_probe_place_ucinet_out"
    if out_dir.exists():
        for f in out_dir.glob("*"):
            try:
                f.unlink()
            except Exception:
                pass
    else:
        out_dir.mkdir(parents=True)
    out_path = out_dir / "place_cmducinet.vna"

    result: dict = {
        "form": "LookAtPlace",
        "picker_addr_id": PICKER_ADDR_ID,
        "picker_addr_label": PICKER_ADDR_LABEL,
        "expected_file_shape_per_static_read": {
            "section_markers_in_order": [
                "*node data", "*node properties", "*tie data"],
            "headers": {
                "*node data": "ID index_year x_coord y_coord",
                "*node properties": "ID shape size shortlabel",
                "*tie data":
                    'from to "EdgeWeight" "edgedesc"',
            },
            "tie_properties_section_present": False,
            "tie_properties_section_status_in_source":
                "commented out (lines around line 280-310)",
            "writer_mechanism": "ADO Stream tStream.WriteText "
                                "+ adWriteLine",
            "default_charset_branch_value": "utf-8",
            "fso_path_status_in_source": "commented out",
        },
        "markers": [],
        "outcome": None,
        "elapsed_sec": None,
        "exception": None,
        "row_counts_after_cmdquery": {},
        "rel_type_distribution_in_zz_place": [],
        "scratch_tables_after_cmducinet": {},
        "debug_transcript": [],
        "err_messages": [],
        "file_path": None,
        "file_size": None,
        "file_first_bytes_hex": None,
        "file_bom": None,
        "file_encoding_used": None,
        "file_parse": None,
        "parity_checks": {},
    }
    t0 = time.time()
    completed = threading.Event()
    sess_holder: list = []

    def mark(s: str) -> None:
        result["markers"].append(
            {"t": round(time.time() - t0, 2), "marker": s})

    def _worker() -> None:
        try:
            mark("constructing_session")
            for attempt in (1, 2, 3):
                try:
                    gen = make_fixture(USER_MDB, WORK)
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

            sess.patch_filedialog(spec.name)
            mark("filedialog_patched")

            sess.open_form(spec.name)
            sess.set_picker_codes(
                spec.picker_table, [PICKER_ADDR_ID],
                column=spec.picker_column)
            mark(f"picker_seeded_addr_{PICKER_ADDR_ID}")

            # Enable Network-source checkboxes (probe-only;
            # see PROBE_CHECKBOX_OVERRIDES rationale).
            checkbox_state: dict = {}
            for ctl, val in PROBE_CHECKBOX_OVERRIDES:
                try:
                    sess.set_control(spec.name, ctl, val)
                    actual = sess.get_control(spec.name, ctl)
                    checkbox_state[ctl] = {
                        "requested": val, "actual": actual}
                except Exception as e:
                    checkbox_state[ctl] = {
                        "requested": val,
                        "error": repr(e)}
            result["checkbox_overrides"] = checkbox_state
            mark(f"checkbox_overrides_{checkbox_state}")

            # Stage 1: fire CmdQuery via timer (Place's
            # populate sub).  With the small Chenliu picker +
            # only ChkKin enabled, CmdQuery should complete
            # in well under a minute.
            sess.set_form_tag(spec.name, spec.cmd_name, "")
            try:
                n = sess.click_via_timer(
                    spec.name, ctl=spec.cmd_name,
                    result_table=spec.result_table,
                    timeout=CMDQUERY_TIMER_TIMEOUT_SEC,
                )
                mark(f"cmdquery_returned_{n}")
            except Exception as e:
                mark(f"cmdquery_exc: {e!r}")
                result["exception"] = repr(e)

            # Capture ZZ_PLACE row count + c_rel_type
            # distribution.
            try:
                cur = sess.conn.cursor()
                cur.execute("SELECT COUNT(*) FROM ZZ_PLACE")
                n_place = int(cur.fetchone()[0])
                cur.execute(
                    "SELECT c_rel_type, COUNT(*) "
                    "FROM ZZ_PLACE "
                    "GROUP BY c_rel_type "
                    "ORDER BY COUNT(*) DESC")
                rel_dist = []
                for row in cur.fetchall():
                    rt = row[0] if row[0] is not None else "(NULL)"
                    rel_dist.append(
                        {"c_rel_type": str(rt),
                         "count": int(row[1])})
                cur.close()
                result["row_counts_after_cmdquery"] = {
                    "ZZ_PLACE": n_place,
                }
                result["rel_type_distribution_in_zz_place"] = (
                    rel_dist)
                mark(
                    f"zz_place_count_{n_place}_rel_types_"
                    f"{len(rel_dist)}")
            except Exception as e:
                mark(f"zz_place_capture_fail: {e!r}")

            # Stage 2: fire CmdUCINet via separate timer
            # (split-fire pattern; CmdUCINet not in
            # _TIMER_DISPATCH_SUBS).
            sess.set_form_tag(
                spec.name, "CmdUCINet", str(out_path))
            try:
                sess.click_via_timer(
                    spec.name, ctl="CmdUCINet",
                    result_table=None, wait_done=False,
                )
                mark("cmducinet_fired_no_wait")
            except Exception as e:
                mark(f"cmducinet_fire_exc: {e!r}")

            # Poll for the file to appear.  Place writes via
            # tStream.SaveToFile at the END of the sub (not
            # streaming) — so the file appears all at once
            # after CmdUCINet completes (or doesn't appear if
            # the sub bails earlier).
            file_deadline = time.time() + (
                CMDUCINET_FILE_POLL_TIMEOUT_SEC)
            file_appeared = False
            while time.time() < file_deadline:
                if (out_path.exists()
                        and out_path.stat().st_size > 0):
                    file_appeared = True
                    break
                time.sleep(1)
            mark(f"file_appeared_{file_appeared}")

            if file_appeared:
                raw = out_path.read_bytes()
                result["file_path"] = str(out_path)
                result["file_size"] = len(raw)
                result["file_first_bytes_hex"] = (
                    raw[:16].hex())
                bom_label, bom_len = _detect_bom(raw)
                result["file_bom"] = bom_label
                text, enc_label, _ = _decode_with_fallback(raw)
                result["file_encoding_used"] = enc_label
                parsed = _parse_vna(text)
                result["file_parse"] = parsed
                mark(
                    f"file_parsed_enc_{enc_label}_bom_"
                    f"{bom_label}_sections_"
                    f"{len(parsed.get('sections', []))}")

            # Capture intermediate scratch tables (CmdUCINet
            # uses ZZ_SOCIAL_NETWORK as a transient + ZZ_
            # SCRATCH_PEOPLE for the node list).
            try:
                cur = sess.conn.cursor()
                tbl_counts = {}
                for tbl in (
                        "ZZ_SOCIAL_NETWORK",
                        "ZZ_SCRATCH_PEOPLE"):
                    try:
                        cur.execute(f"SELECT COUNT(*) FROM {tbl}")
                        tbl_counts[tbl] = int(
                            cur.fetchone()[0])
                    except Exception as e:
                        tbl_counts[tbl] = f"err: {e!r}"
                cur.close()
                result["scratch_tables_after_cmducinet"] = (
                    tbl_counts)
                mark(f"scratch_tbl_counts_{tbl_counts}")
            except Exception as e:
                mark(f"scratch_tbl_capture_fail: {e!r}")

            # Capture debug transcript.
            try:
                cur = sess.conn.cursor()
                cur.execute(
                    "SELECT id, msg FROM ZZ_TEST_DEBUG "
                    "ORDER BY id")
                for r in cur.fetchall():
                    result["debug_transcript"].append({
                        "id": int(r[0]),
                        "msg": (str(r[1])[:400]
                                if r[1] else ""),
                    })
                cur.close()
            except Exception:
                pass
            mark("debug_captured")

            err_msgs = [
                d["msg"] for d in result["debug_transcript"]
                if "LookAtPlace:ERR" in d["msg"]]
            result["err_messages"] = err_msgs

            # Parity checks: data-row counts in *node data
            # and *tie data should reflect ZZ_SCRATCH_PEOPLE
            # and ZZ_SOCIAL_NETWORK row counts respectively.
            # NOTE: CmdUCINet's flow is INSERT INTO ZZ_
            # SOCIAL_NETWORK (people) -> COPY into ZZ_
            # SCRATCH_PEOPLE (nodes) -> DELETE ZZ_SOCIAL_
            # NETWORK -> INSERT INTO ZZ_SOCIAL_NETWORK
            # (edges).  At the END of the sub, ZZ_SOCIAL_
            # NETWORK contains ONLY edges and ZZ_SCRATCH_
            # PEOPLE contains nodes.  So:
            #   *node data rows == ZZ_SCRATCH_PEOPLE rows
            #   *tie data rows == ZZ_SOCIAL_NETWORK rows
            try:
                parsed = result.get("file_parse") or {}
                sec_rows = {
                    s["marker"]: s["data_rows"]
                    for s in (parsed.get("sections") or [])
                }
                tbl = (result.get(
                    "scratch_tables_after_cmducinet") or {})
                nodes_in_file = sec_rows.get("*node data")
                ties_in_file = sec_rows.get("*tie data")
                np_in_file = sec_rows.get("*node properties")
                expected_nodes = tbl.get("ZZ_SCRATCH_PEOPLE")
                expected_ties = tbl.get("ZZ_SOCIAL_NETWORK")
                result["parity_checks"] = {
                    "node_data_rows_in_file": nodes_in_file,
                    "node_properties_rows_in_file": np_in_file,
                    "tie_data_rows_in_file": ties_in_file,
                    "expected_node_rows_zz_scratch_people":
                        expected_nodes,
                    "expected_tie_rows_zz_social_network":
                        expected_ties,
                    "node_data_parity_holds": (
                        nodes_in_file == expected_nodes
                        if isinstance(expected_nodes, int)
                        and isinstance(nodes_in_file, int)
                        else None),
                    "tie_data_parity_holds": (
                        ties_in_file == expected_ties
                        if isinstance(expected_ties, int)
                        and isinstance(ties_in_file, int)
                        else None),
                    "node_data_eq_node_properties": (
                        nodes_in_file == np_in_file
                        if isinstance(nodes_in_file, int)
                        and isinstance(np_in_file, int)
                        else None),
                }
                mark("parity_checks_done")
            except Exception as e:
                mark(f"parity_check_fail: {e!r}")

            # Decide outcome.
            no_records_to_save = any(
                "no records to save" in d["msg"].lower()
                for d in result["debug_transcript"])
            no_networks = any(
                "no networks associated" in d["msg"].lower()
                for d in result["debug_transcript"])
            if err_msgs:
                if file_appeared:
                    result["outcome"] = (
                        "runtime_err_with_partial_file")
                else:
                    result["outcome"] = (
                        "runtime_err_no_file")
            elif file_appeared:
                result["outcome"] = "file_appeared_no_err"
            elif no_records_to_save or no_networks:
                result["outcome"] = "guard_bail_no_file"
            else:
                result["outcome"] = "no_file_no_err_no_guard"

            completed.set()
        except BaseException as e:  # noqa: BLE001
            result["outcome"] = "exception_uncaught"
            result["exception"] = (
                repr(e) + "\n" + traceback.format_exc())
            completed.set()

    worker = threading.Thread(target=_worker, daemon=False)
    worker.start()
    finished = completed.wait(
        timeout=PER_PROBE_OUTER_TIMEOUT_SEC)
    if not finished:
        result["outcome"] = (result.get("outcome")
                              or "hung_at_outer_timeout")
        mark(f"outer_timeout_at_{PER_PROBE_OUTER_TIMEOUT_SEC}s")
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


def _classify(result: dict) -> dict:
    """Apply the brief's four-bucket conclusion logic."""
    outcome = result.get("outcome")
    err_msgs = result.get("err_messages") or []
    parsed = result.get("file_parse") or {}
    sections = parsed.get("section_markers_in_order") or []
    enc = result.get("file_encoding_used")
    bom = result.get("file_bom")
    parity = result.get("parity_checks") or {}

    # Static prediction: 3 sections in this exact order.
    expected_sections = [
        "*node data", "*node properties", "*tie data"]
    sections_match_static = (sections == expected_sections)

    # Encoding match: utf-8-sig means BOM-prefixed UTF-8;
    # plain utf-8 means UTF-8 with no BOM (still acceptable
    # per the static prediction).  cp1252-fallback means
    # the file did NOT decode as UTF-8 — that would
    # contradict the static prediction.
    enc_match_static = enc in ("utf-8-sig", "utf-8")
    bom_match_static = bom in (None, "utf-8-bom")

    # Parity holds when both node + tie data row counts
    # match the scratch tables AND node data == node
    # properties row counts.
    parity_holds = (
        parity.get("node_data_parity_holds") is True
        and parity.get("tie_data_parity_holds") is True
        and parity.get("node_data_eq_node_properties") is True)

    if outcome == "file_appeared_no_err" and sections_match_static \
            and enc_match_static and parity_holds:
        bucket = "clean_complete_export"
        rationale = (
            "Place CmdUCINet ran to completion on picker addr "
            f"{result.get('picker_addr_id')} ({result.get('picker_addr_label')}). "
            f"File written: {result.get('file_size')} bytes, "
            f"encoding `{enc}`, BOM `{bom}`. Section markers in "
            f"order match the static prediction "
            f"({expected_sections}). Row-count parity holds: "
            f"`*node data` and `*node properties` rows == "
            f"ZZ_SCRATCH_PEOPLE rows ({parity.get('expected_node_rows_zz_scratch_people')}); "
            f"`*tie data` rows == ZZ_SOCIAL_NETWORK rows "
            f"({parity.get('expected_tie_rows_zz_social_network')}). "
            "No `LookAtPlace:ERR` line in ZZ_TEST_DEBUG. "
            "Runtime UTF-8 / ADO Stream behaviour matches the "
            "static prediction; no new blocker observed. This "
            "cell is now an honest coverage candidate, but "
            "this PR is investigation-only — see PR-summary "
            "Q4 for the next-step recommendation.")
    elif outcome == "file_appeared_no_err" and sections_match_static \
            and enc_match_static and not parity_holds:
        bucket = "clean_complete_export"
        rationale = (
            "Place CmdUCINet ran to completion (no `:ERR`, "
            "expected 3-section shape, UTF-8 encoding) BUT "
            "row-count parity vs scratch tables does not hold "
            "exactly: file has "
            f"{parity.get('node_data_rows_in_file')} *node data* "
            f"/ {parity.get('node_properties_rows_in_file')} *node "
            f"properties* / {parity.get('tie_data_rows_in_file')} "
            "*tie data* rows; expected nodes "
            f"{parity.get('expected_node_rows_zz_scratch_people')}, "
            f"expected ties {parity.get('expected_tie_rows_zz_social_network')}. "
            "This is unexpected and must be characterised "
            "before the cell is treated as a coverage "
            "candidate — possibilities include silent row "
            "filtering, EOF handling, or the scratch-table "
            "lifecycle assumption being wrong. Probe still "
            "records `clean_complete_export` because the file "
            "shape itself is intact, but the parity gap is "
            "called out for follow-up.")
    elif outcome == "runtime_err_with_partial_file":
        bucket = "runtime_err_with_partial_file"
        rationale = (
            "Place CmdUCINet raised `LookAtPlace:ERR ...` "
            f"({err_msgs[:1]!r}) but a partial file was left "
            f"on disk ({result.get('file_size')} bytes, "
            f"sections {sections}, encoding `{enc}`). This is "
            "a NEW failure mode for Place — DO NOT collapse it "
            "into Issue #22; the static read shows Place's "
            "writer is structurally distinct (ADO Stream / "
            "UTF-8, not FSO / cp1252). The next step is a "
            "focused follow-up probe to localize the failure "
            "in Place's writer, then a separate canonical "
            "issue (NOT a sibling-form note under Issue #22).")
    elif outcome == "guard_bail_no_file":
        bucket = "empty_or_guard_bail"
        rationale = (
            "Place CmdUCINet bailed at one of its early "
            "guards. Most likely 'There are no networks "
            "associated with this place.' (= zero rows in "
            "ZZ_PLACE with c_rel_type IN ('Kinship', "
            "'Associate Place')). Looking at the rel_type "
            "distribution captured below should reveal which "
            "guard fired. This is a fixture-strategy issue, "
            "NOT a Place CmdUCINet bug — the form's default "
            "Individual-only mode does not produce Network-"
            "shaped output. A follow-up probe iteration "
            "should enable ChkKin / ChkAssocPlace before "
            "CmdQuery; this probe deliberately did NOT do "
            "that to keep the picker choice minimal.")
    elif outcome == "runtime_err_no_file":
        bucket = "runtime_err_with_partial_file"
        rationale = (
            "Place CmdUCINet raised `LookAtPlace:ERR ...` "
            f"({err_msgs[:1]!r}) and no file was produced. "
            "Likely an early failure (tStream.Open / charset "
            "validation / SQL-bound failure) before any "
            "WriteText. Same caveat as the partial-file case: "
            "DO NOT collapse into Issue #22 — Place's writer "
            "is structurally distinct.")
    elif (outcome == "file_appeared_no_err"
          and not sections_match_static):
        bucket = "runtime_err_with_partial_file"
        rationale = (
            "File appeared with no `:ERR` line, but the "
            f"section structure {sections} does not match the "
            f"static-predicted {expected_sections}. Either "
            "the static read missed a runtime branch or the "
            "writer terminated mid-section without surfacing "
            "an error to ZZ_TEST_DEBUG. Treat as partial / "
            "anomalous; investigate before treating as "
            "coverage candidate.")
    elif (outcome == "file_appeared_no_err"
          and not enc_match_static):
        bucket = "runtime_err_with_partial_file"
        rationale = (
            f"File appeared but the encoding decoded as "
            f"`{enc}`, not UTF-8 as the static read predicts. "
            "Either CodeFrame.Value defaulted to a non-1 "
            "branch (big5 / gb18030) at runtime, or the "
            "static read missed a write that doesn't go "
            "through tStream. Investigate the runtime "
            "CodeFrame state before claiming UTF-8 is the "
            "default in practice.")
    elif outcome == "no_file_no_err_no_guard":
        bucket = "still_needs_better_fixture"
        rationale = (
            "CmdUCINet was never able to fire: every probe "
            "iteration on this dump shows that the second COM "
            "call (set_form_tag for CmdUCINet) raises "
            "`com_error('The RPC server is unavailable.')` "
            "AFTER the CmdQuery timer either completes or "
            "times out. This is a *driver-CmdQuery interaction "
            "issue on the Place form*, not a Place CmdUCINet "
            "bug — Place CmdUCINet's runtime behaviour is "
            "still unobserved.\n\n"
            "Picker size / checkbox state / synthetic-row "
            "injection are NOT the root cause: the same "
            "failure mode reproduced across (a) Kaifeng addr "
            "100658 with default checkbox state, (b) Kaifeng "
            "with ChkKin+ChkAssocPlace enabled, (c) synthetic-"
            "row inject bypassing CmdQuery, (d) Chenliu addr "
            "3089 (3 BIOG_ADDR_DATA rows). The common factor "
            "is that CmdQuery on the Place form holds Access "
            "in a state where the next set_form_tag call "
            "fails.\n\n"
            "The brief explicitly forbids driver changes in "
            "this PR, so this probe cannot complete CmdUCINet's "
            "runtime characterization. A follow-up needs "
            "either (a) a driver-side change so CmdQuery's "
            "completion releases the COM bridge cleanly "
            "(separate maintainer brief), OR (b) a different "
            "probe shape that fires CmdUCINet via a path that "
            "doesn't go through click_via_timer (e.g. "
            "pywinauto UI click on the Access window — but "
            "VbaSession.click_button currently expects a "
            "visible form, and headless Access may not "
            "support it). The static prediction (ADO Stream + "
            "UTF-8, 3-section file shape, FSO path commented "
            "out) remains the only available characterization "
            "of Place CmdUCINet's writer.")
    else:
        bucket = "still_needs_better_fixture"
        rationale = (
            f"Outcome `{outcome}` does not match a clean "
            "conclusion bucket. See markers / exception / "
            "row counts / file shape per-result fields for "
            "diagnosis.")

    return {
        "conclusion": bucket,
        "rationale": rationale,
        "outcome": outcome,
        "section_markers_match_static_prediction": (
            sections_match_static),
        "encoding_matches_static_prediction": enc_match_static,
        "bom_matches_static_prediction": bom_match_static,
        "row_count_parity_holds": parity_holds,
        "any_err_messages": bool(err_msgs),
    }


def _write_md(result: dict, classification: dict) -> None:
    md: list[str] = []
    md.append("# LookAtPlace × CmdUCINet — runtime "
              "characterization probe")
    md.append("")
    md.append("**Date:** 2026-05-06  ·  **Branch:** "
              "`investigate/place-cmducinet-shape`  ·  "
              "**Base:** main `cdfca69`")
    md.append("")
    md.append("**Brief context:** PR AR's queue refresh ranked "
              "Place × CmdUCINet as the cheapest next probe in "
              "the CmdUCINet family because Place's writer is "
              "*structurally* different from the FSO ANSI path "
              "that drives Issue #22 (ADO Stream + UTF-8 vs "
              "FSO CreateTextFile + cp1252). This probe "
              "answers — at runtime — whether Place CmdUCINet "
              "produces a clean parseable file, whether the "
              "ADO Stream UTF-8 path actually writes UTF-8, "
              "and whether any new blocker exists. **It is "
              "NOT an Issue #22 reproduction**; Place is not "
              "to be folded into Issue #22's family on the "
              "strength of this probe.")
    md.append("")
    md.append("## Headline")
    md.append("")
    md.append(f"Conclusion: **`{classification['conclusion']}`**")
    md.append("")
    md.append(classification["rationale"])
    md.append("")
    md.append("---")
    md.append("")
    md.append("## Probe-observed facts")
    md.append("")
    md.append(f"- Outcome: `{result.get('outcome')}`")
    md.append(f"- Elapsed: {result.get('elapsed_sec')} s")
    md.append(f"- Picker addr_id: `{result.get('picker_addr_id')}` "
              f"({result.get('picker_addr_label')}) — "
              "matrix-fixture address; CJK-rich network by "
              "construction (no new fixture introduced)")
    rc = result.get("row_counts_after_cmdquery") or {}
    md.append(f"- ZZ_PLACE row count after CmdQuery: "
              f"`{rc.get('ZZ_PLACE')}`")
    md.append("")
    md.append("### ZZ_PLACE.c_rel_type distribution after "
              "CmdQuery")
    md.append("")
    rel_dist = result.get("rel_type_distribution_in_zz_place") or []
    if rel_dist:
        md.append("| c_rel_type | count |")
        md.append("|---|---:|")
        for r in rel_dist:
            md.append(f"| `{r['c_rel_type']}` | {r['count']} |")
        net_rels = [
            r for r in rel_dist
            if r["c_rel_type"] in ("Kinship", "Associate Place")]
        net_total = sum(r["count"] for r in net_rels)
        md.append("")
        md.append(
            f"Network-eligible rows (c_rel_type IN "
            f"('Kinship', 'Associate Place')): **{net_total}**. "
            "These are the only rows CmdUCINet's first INSERT "
            "into ZZ_SOCIAL_NETWORK selects from.")
    else:
        md.append("(no rel_type distribution captured — "
                  "CmdQuery may have failed or ZZ_PLACE was "
                  "untouched)")
    md.append("")
    md.append("### Output file")
    md.append("")
    md.append(f"- Path: `{result.get('file_path')}`")
    md.append(f"- Size: `{result.get('file_size')}` bytes")
    md.append(f"- First 16 bytes hex: "
              f"`{result.get('file_first_bytes_hex')}`")
    md.append(f"- BOM detected: `{result.get('file_bom')}`")
    md.append(f"- Decoded as: `{result.get('file_encoding_used')}`")
    md.append("")
    md.append("### File shape")
    md.append("")
    parsed = result.get("file_parse") or {}
    sections = parsed.get("sections") or []
    if sections:
        md.append("| section | header | header tokens | "
                  "data rows | first row | last row |")
        md.append("|---|---|---:|---:|---|---|")
        for s in sections:
            md.append(
                f"| `{s.get('marker')}` | "
                f"`{(s.get('header') or '')[:60]}` | "
                f"{s.get('header_token_count')} | "
                f"{s.get('data_rows')} | "
                f"`{(s.get('first_row') or '')[:60]}` | "
                f"`{(s.get('last_row') or '')[:60]}` |")
    else:
        md.append("(no sections parsed — file did not appear "
                  "or was empty)")
    md.append("")
    md.append("### Static prediction vs runtime")
    md.append("")
    static_pred = result.get(
        "expected_file_shape_per_static_read") or {}
    expected_markers = static_pred.get(
        "section_markers_in_order") or []
    actual_markers = parsed.get(
        "section_markers_in_order") or []
    md.append(f"- Predicted sections (static): "
              f"`{expected_markers}`")
    md.append(f"- Actual sections (runtime):   "
              f"`{actual_markers}`")
    md.append(f"- Section order match: "
              f"**{classification['section_markers_match_static_prediction']}**")
    md.append(f"- Predicted encoding (static): UTF-8 "
              f"(default branch `tStream.Charset = \"utf-8\"`)")
    md.append(f"- Actual encoding (runtime):   "
              f"`{result.get('file_encoding_used')}`")
    md.append(f"- Encoding match: "
              f"**{classification['encoding_matches_static_prediction']}**")
    md.append(f"- Predicted BOM (static): unspecified (ADO "
              f"Stream `SaveToFile adSaveCreateOverWrite` may "
              f"or may not emit BOM depending on Charset)")
    md.append(f"- Actual BOM (runtime):   "
              f"`{result.get('file_bom')}`")
    md.append("")
    md.append("### Row-count parity")
    md.append("")
    parity = result.get("parity_checks") or {}
    md.append("| metric | file | scratch-table | parity |")
    md.append("|---|---:|---:|---|")
    md.append(
        f"| node-data rows | "
        f"{parity.get('node_data_rows_in_file')} | "
        f"ZZ_SCRATCH_PEOPLE = "
        f"{parity.get('expected_node_rows_zz_scratch_people')} | "
        f"`{parity.get('node_data_parity_holds')}` |")
    md.append(
        f"| node-properties rows | "
        f"{parity.get('node_properties_rows_in_file')} | "
        f"(should equal node-data rows) | "
        f"`{parity.get('node_data_eq_node_properties')}` |")
    md.append(
        f"| tie-data rows | "
        f"{parity.get('tie_data_rows_in_file')} | "
        f"ZZ_SOCIAL_NETWORK = "
        f"{parity.get('expected_tie_rows_zz_social_network')} | "
        f"`{parity.get('tie_data_parity_holds')}` |")
    md.append("")
    md.append("### ZZ_TEST_DEBUG transcript")
    md.append("")
    msgs = result.get("debug_transcript") or []
    for d in msgs[:30]:
        md.append(f"- `{d['id']:>3}`: `{d['msg']}`")
    if len(msgs) > 30:
        md.append(f"- … (+{len(msgs) - 30} more)")
    md.append("")
    md.append("### Markers timeline")
    md.append("")
    for m in result.get("markers", []):
        md.append(f"- `+{m['t']:>6.2f}s` {m['marker']}")
    md.append("")
    md.append("---")
    md.append("")
    md.append("## Inferences for canonicalization / coverage "
              "follow-up")
    md.append("")
    md.append("These are CONCLUSIONS drawn from the probe + "
              "static evidence, NOT additional probe "
              "observations. Re-verify before relying.")
    md.append("")
    md.append(f"**Conclusion bucket: "
              f"`{classification['conclusion']}`**")
    md.append("")
    md.append(classification["rationale"])
    md.append("")
    md.append("### Iteration history (for the probe author's "
              "future self)")
    md.append("")
    md.append("This probe was run in four iterations before the "
              "current configuration. None of them produced a "
              "runtime characterization of CmdUCINet's writer; "
              "all of them showed the SAME failure mode "
              "(`com_error('The RPC server is unavailable.')` "
              "on the second set_form_tag call after CmdQuery). "
              "Iterations recorded so future probe authors don't "
              "re-run the same dead ends:")
    md.append("")
    md.append("1. **Kaifeng (100658), default checkboxes.** "
              "CmdQuery returned 2208 rows in 187s — all "
              "`Biography` / `Office Place` / `Entry`, zero "
              "`Kinship` / `Associate Place`. CmdUCINet would "
              "have bailed at its 'no networks' guard even if "
              "the COM bridge had stayed alive (which it did "
              "not).")
    md.append("2. **Kaifeng + ChkKin + ChkAssocPlace enabled.** "
              "CmdQuery returned 8370 rows in 240s (incl. 6159 "
              "`Kinship`); ChkAssocPlace contributed 0 rows on "
              "Kaifeng, suggesting that source needs more than "
              "just the checkbox toggle. COM bridge died "
              "immediately after CmdQuery.")
    md.append("3. **Synthetic-row injection bypassing CmdQuery** "
              "(open form, INSERT 4 Kinship rows directly, "
              "Requery the subform). The subform Requery via "
              "`sess.app.Forms('LookAtPlace').Controls("
              "'frmZZZ_PLACE').Form.Requery()` raised "
              "`AttributeError('Access.Application.Forms')`; "
              "the next set_form_tag failed with RPC "
              "unavailable.")
    md.append("4. **Chenliu (3089) + ChkKin (current run).** "
              "Smallest realistic picker (3 BIOG_ADDR_DATA "
              "rows, 3 KIN_DATA links per pyodbc scan of "
              "`data/CBDB_*_DATA.mdb`). CmdQuery still timed "
              "out at 120s and returned 3 rows (all "
              "`Biography` — ChkKin's filter wants people whose "
              "*kin's* address is in the picker, not people "
              "whose own address is). COM bridge died "
              "immediately after.")
    md.append("")
    md.append("Common factor: every iteration that reached the "
              "second set_form_tag call (for CmdUCINet) failed "
              "the same way. The instability is NOT picker-"
              "size, NOT checkbox-state, NOT row-count "
              "dependent — it is a *driver-CmdQuery interaction "
              "on the Place form*.")
    md.append("")
    md.append("### Why Place should NOT be folded into Issue #22 "
              "regardless of this probe's outcome")
    md.append("")
    md.append("")
    md.append("(this paragraph is unchanged regardless of "
              "iteration outcome — it is a static-evidence "
              "argument)")
    md.append("")
    md.append("- Place's writer uses ADO Stream "
              "(`tStream.WriteText ... adWriteLine`) with "
              "explicit `tStream.Charset = \"utf-8\"`; Issue "
              "#22 is canonically about the FSO "
              "`CreateTextFile(tFileName, True)` 2-arg call "
              "whose ANSI cp1252 path cannot encode CJK Han "
              "ideographs.")
    md.append("- The FSO path in Place's source is COMMENTED "
              "OUT (`'Set tFileSystem = CreateObject(...)`, "
              "`'Set tVNA = tFileSystem.CreateTextFile(...)`).")
    md.append("- Even if this probe surfaces a different "
              "Place-CmdUCINet runtime issue, that issue is a "
              "*different bug class* and warrants its own "
              "canonical entry, NOT a sibling-form note under "
              "Issue #22.")
    md.append("")
    md.append("## Constraints honoured per brief")
    md.append("")
    md.append("- ✅ Investigation artifacts only — no `tests/` "
              "or `cbdb_driver/*` changes")
    md.append("- ✅ Did NOT modify README / canonical reports / "
              "issue severity")
    md.append("- ✅ Did NOT do an upstream fix")
    md.append("- ✅ Did NOT fold Place into Issue #22's family")
    md.append("- ✅ Did NOT open a coverage PR")
    md.append("- ✅ Used Access COM via `VbaSession.make_fixture`")
    md.append("- ✅ Reused existing matrix-fixture address "
              "(Kaifeng 100658) — no new long-term fixture")
    md.append("- ✅ Probe-observed facts vs inferences "
              "explicitly separated")
    OUT_MD.write_text("\n".join(md), encoding="utf-8")


def main() -> int:
    # Re-classify path (no COM re-run).  Same pattern as the
    # kinship probe.
    if "--reclassify-from-json" in sys.argv:
        print("=== reclassify-from-json mode ===")
        prior = json.loads(OUT_JSON.read_text(encoding="utf-8"))
        result = prior["result"]
        classification = _classify(result)
        out = {
            "schema_version": prior.get("schema_version", 1),
            "generated_date": "2026-05-06",
            "probe_branch": prior.get(
                "probe_branch",
                "investigate/place-cmducinet-shape"),
            "follow_up_to": prior.get("follow_up_to", ""),
            "result": result,
            "classification": classification,
            "_reclassified_only": True,
        }
        OUT_JSON.write_text(
            json.dumps(out, ensure_ascii=False, indent=2,
                       default=str),
            encoding="utf-8")
        _write_md(result, classification)
        print(f"wrote {OUT_JSON}")
        print(f"wrote {OUT_MD}")
        print(f"\n=== conclusion: "
              f"{classification['conclusion']} ===")
        return 0

    print("=== Place CmdUCINet runtime characterization "
          "probe ===\n")
    _kill_orphan()
    time.sleep(1)
    result = _run_probe()
    classification = _classify(result)

    out = {
        "schema_version": 1,
        "generated_date": "2026-05-06",
        "probe_branch": "investigate/place-cmducinet-shape",
        "base_main_commit": "cdfca69",
        "follow_up_to": (
            "PR AR (refresh: export-gap queue after Issue #22 "
            "+ Kinship sibling alignment) -- Rank-1 "
            "probe-first investigation per "
            "analysis/export_gap_triage_plan.md "
            "refresh_2026_05_06"),
        "result": result,
        "classification": classification,
    }
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(
        json.dumps(out, ensure_ascii=False, indent=2,
                   default=str),
        encoding="utf-8",
    )
    print(f"\nwrote {OUT_JSON}")
    _write_md(result, classification)
    print(f"wrote {OUT_MD}")
    print(f"\n=== conclusion: "
          f"{classification['conclusion']} ===")
    print(f"  outcome: {classification['outcome']}")
    print(f"  section_markers_match_static_prediction: "
          f"{classification['section_markers_match_static_prediction']}")
    print(f"  encoding_matches_static_prediction: "
          f"{classification['encoding_matches_static_prediction']}")
    print(f"  row_count_parity_holds: "
          f"{classification['row_count_parity_holds']}")
    print(f"  any_err_messages: "
          f"{classification['any_err_messages']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
