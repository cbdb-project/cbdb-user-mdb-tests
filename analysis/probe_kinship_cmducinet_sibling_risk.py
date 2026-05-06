"""Probe: does LookAtKinship.CmdUCINet have the same FSO
ANSI-vs-Unicode-source vulnerability as Issue #22?

Goal
----
Issue #22 canonicalized the LookAtAssociations.CmdUCINet
crash on c_name fields containing CJK Han ideographs (no
cp1252 fallback in `Scripting.FileSystemObject.CreateTextFile
(tFileName, True)`'s ANSI write path).  The Issue #22 summary
flagged Kinship as "possible sibling risk / NOT yet probed"
because:

  - `Form_LookAtKinship.CmdUCINet_Click` line 2510 has the
    same `Set tVNA = tFileSystem.CreateTextFile(tFileName,
    True)` 2-arg call pattern (third Unicode arg omitted).
  - The `*node properties` block writes `!c_kin_name` from
    `tRstNode = CurrentDb.OpenRecordset("ZZ_SCRATCH_KIN",
    dbOpenDynaset)` — same FSO write call, same exposure to
    non-cp1252 source data.
  - The current Kinship × CmdUCINet coverage
    (`tests/test_vba_cmducinet_kinship.py`) passes only
    because person 3211's kin network happens to have no
    Han-name members.  THIS coverage was explicitly flagged
    as fixture-fragile in the canonical Issue #22 text.

This probe answers: under a fixture whose kin network DOES
include a Han-name member, does Kinship CmdUCINet bail with
the same `:ERR Invalid procedure call or argument`?

Method
------
1. Pre-scan BIOG_MAIN.c_name for non-cp1252-without-FSO-
   substitution chars (CJK Han ideographs in particular).
   Identify candidate "trigger persons".
2. Find a SMALL Kinship fixture whose 1-hop kin network
   reaches a trigger person.  This probe uses
   **picker = 152930 (He Jing 何淨)**, whose sole direct
   kin is pid 140733 (He Mou 取, where 取 = U+53D6 = Han
   ideograph, same trigger class as Issue #22's 稜
   = U+7A1C).
3. Drive CmdRun via timer (autodetect chain pattern); poll
   ZZ_SCRATCH_KIN row count + scan its c_kin_name field for
   trigger characters.
4. Drive CmdUCINet via separate timer (split-fire pattern;
   CmdUCINet not in `_TIMER_DISPATCH_SUBS`); poll for
   partial-file appearance.
5. Parse the partial `.vna` file; identify the LAST
   successfully-written *node properties* row's first token
   (= c_kin_id).
6. Cross-check: the row immediately after the last-written
   one (in JET iteration order) should be the c_kin_name
   that triggered the bail.

Conclusion buckets (per brief)
------------------------------
- `same_bug_family_runtime_confirmed`
    - `:ERR Invalid procedure call or argument` reproduces
      AND the partial file shape matches Issue #22's
      pattern (full `*node data`, truncated `*node
      properties`, no `*tie data`)
    - AND the would-be-next row's `c_kin_name` contains a
      non-cp1252 char without FSO substitution
- `sibling_risk_not_reproduced_on_current_fixtures`
    - Chain runs clean OR fails with a different error
      class.  Two sub-cases tagged in the rationale:
      (a) genuinely robust code path (would mean the
          source code differs from what static read shows
          — unlikely)
      (b) no triggering row in the fixture's reachable
          network (would mean the fixture chosen here
          doesn't actually reach 140733 — should be
          obvious from the scan of ZZ_SCRATCH_KIN).
- `still_needs_better_fixture`
    - Setup-phase failure (RPC flake, fixture didn't load,
      etc.) such that the question genuinely wasn't
      answered by this probe run.

Outputs
-------
- analysis/probe_kinship_cmducinet_sibling_risk.md
- reports/probe_kinship_cmducinet_sibling_risk.json
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
WORK = ROOT / "analysis" / "_probe_kin_ucinet_sibling_copy.mdb"
OUT_JSON = (ROOT / "reports"
            / "probe_kinship_cmducinet_sibling_risk.json")
OUT_MD = (ROOT / "analysis"
          / "probe_kinship_cmducinet_sibling_risk.md")

# Custom fixture for THIS probe (NOT a matrix fixture).
# Picker: pid 152930 (He Jing 何淨).  Their sole 1-hop kin
# row points to pid 140733 (He Mou 取), whose c_name in
# BIOG_MAIN contains Han ideograph 取 (U+53D6).  This makes
# the kin recursion reach a trigger c_kin_name within
# minimal expansion, guaranteeing ZZ_SCRATCH_KIN contains
# the trigger row.
PICKER_PID = 152930
EXPECTED_TRIGGER_KIN_PID = 140733
EXPECTED_TRIGGER_HAN_CHAR = "取"
EXPECTED_TRIGGER_HAN_CODEPOINT = 0x53D6

PER_PROBE_OUTER_TIMEOUT_SEC = 360
CMDRUN_TIMER_TIMEOUT_SEC = 180
CMDUCINET_FILE_POLL_TIMEOUT_SEC = 60


def _kill_orphan():
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "MSACCESS.EXE"],
            capture_output=True, timeout=10)
    except Exception:
        pass


def _non_cp1252_no_substitute_chars(s: str) -> list[tuple[int, str, int]]:
    """Return chars in `s` that fail cp1252 strict encoding
    AND are NOT U+3000 (Ideographic Space — Issue #22
    findings showed FSO silently substitutes that one to
    ASCII space; only chars without an FSO substitution
    crash).  Tuples are (offset, char, codepoint)."""
    bad = []
    for i, ch in enumerate(s):
        if ord(ch) == 0x3000:
            continue
        try:
            ch.encode("cp1252", errors="strict")
        except UnicodeEncodeError:
            bad.append((i, ch, ord(ch)))
    return bad


def _parse_partial_vna(text: str) -> dict:
    lines = text.replace("\r\n", "\n").split("\n")
    sections: list[dict] = []
    cur: dict | None = None
    for ln in lines:
        s = ln.strip()
        if s.startswith("*"):
            if cur is not None:
                sections.append(cur)
            cur = {"marker": s, "header": None, "rows": []}
            continue
        if cur is not None and s:
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
            "total_rows": len(s["rows"]),
            "last_row": s["rows"][-1] if s["rows"] else None,
            "last_row_first_token": (
                s["rows"][-1].split()[0] if s["rows"]
                else None),
        }
        for s in sections
    ]
    return {
        "sections": out_sections,
        "section_markers_in_order": [
            s["marker"] for s in out_sections],
    }


def _run_probe() -> dict:
    from cbdb_driver.vba_session import make_fixture
    from cbdb_driver.form_specs import LOOKATKINSHIP

    spec = LOOKATKINSHIP
    out_dir = ROOT / "analysis" / "_probe_kin_ucinet_sibling_out"
    if out_dir.exists():
        for f in out_dir.glob("*"):
            try:
                f.unlink()
            except Exception:
                pass
    else:
        out_dir.mkdir(parents=True)
    out_path = out_dir / "cmducinet_kinship_sibling.vna"

    result: dict = {
        "form": "LookAtKinship",
        "picker_pid": PICKER_PID,
        "expected_trigger_kin_pid": EXPECTED_TRIGGER_KIN_PID,
        "expected_trigger_char": EXPECTED_TRIGGER_HAN_CHAR,
        "expected_trigger_codepoint": (
            f"U+{EXPECTED_TRIGGER_HAN_CODEPOINT:04X}"),
        "markers": [],
        "outcome": None,
        "elapsed_sec": None,
        "exception": None,
        "row_counts_after_cmdrun": {},
        "scratch_kin_scan": None,
        "debug_transcript": [],
        "file_path": None,
        "file_size": None,
        "partial_file_parse": None,
        "row_after_last_written": None,
    }
    t0 = time.time()
    completed = threading.Event()
    sess_holder: list = []

    def mark(s):
        result["markers"].append(
            {"t": round(time.time() - t0, 2), "marker": s})

    def _worker():
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
                spec.picker_table, [PICKER_PID],
                column=spec.picker_column)
            mark(f"picker_seeded_pid_{PICKER_PID}")

            # Stage 1: fire CmdRun via timer (Kinship's
            # populate sub).
            sess.set_form_tag(spec.name, spec.cmd_name, "")
            try:
                n = sess.click_via_timer(
                    spec.name, ctl=spec.cmd_name,
                    result_table=spec.result_table,
                    timeout=CMDRUN_TIMER_TIMEOUT_SEC,
                )
                mark(f"cmdrun_returned_{n}")
            except Exception as e:
                mark(f"cmdrun_exc: {e!r}")
                result["exception"] = repr(e)

            # Capture scratch row counts + scan ZZ_SCRATCH_KIN
            # for non-cp1252-without-substitute c_kin_name
            # values.  This is the hypothesis pre-test (it
            # answers "does the trigger row reach
            # ZZ_SCRATCH_KIN at all?").
            cur = sess.conn.cursor()
            cur.execute("SELECT COUNT(*) FROM ZZ_SCRATCH_KIN")
            n_kin = int(cur.fetchone()[0])
            cur.execute(
                "SELECT COUNT(*) FROM ZZ_SCRATCH_KINNET")
            n_kinnet = int(cur.fetchone()[0])
            result["row_counts_after_cmdrun"] = {
                "ZZ_SCRATCH_KIN": n_kin,
                "ZZ_SCRATCH_KINNET": n_kinnet,
            }
            cur.close()
            mark(f"row_counts_kin_{n_kin}_kinnet_{n_kinnet}")

            # Scan ZZ_SCRATCH_KIN.c_kin_name + capture order
            # by c_kin_id (default JET iteration order
            # equivalent for the *node properties* loop).
            cur = sess.conn.cursor()
            cur.execute(
                "SELECT c_kin_id, c_kin_name "
                "FROM ZZ_SCRATCH_KIN "
                "ORDER BY c_kin_id ASC"
            )
            scan_total = 0
            scan_bad = []
            scan_ids_in_order = []
            scan_first_bad_idx = None
            for idx, row in enumerate(cur.fetchall()):
                scan_total += 1
                kid = (int(row[0])
                       if row[0] is not None else None)
                name = row[1] or ""
                scan_ids_in_order.append(kid)
                bad = _non_cp1252_no_substitute_chars(name)
                if bad:
                    scan_bad.append({
                        "kid_order_index_0based": idx,
                        "c_kin_id": kid,
                        "c_kin_name": name,
                        "non_cp1252_chars": [
                            {"offset": o, "char": ch,
                             "codepoint_hex": f"U+{cp:04X}"}
                            for o, ch, cp in bad],
                    })
                    if scan_first_bad_idx is None:
                        scan_first_bad_idx = idx
            cur.close()
            result["scratch_kin_scan"] = {
                "scanned_rows": scan_total,
                "rows_with_non_cp1252_no_substitute_c_kin_name": (
                    len(scan_bad)),
                "first_bad_index_kid_asc_0based": (
                    scan_first_bad_idx),
                "samples": scan_bad[:10],
                "trigger_pid_in_scratch_kin": (
                    EXPECTED_TRIGGER_KIN_PID
                    in scan_ids_in_order),
            }
            mark(
                f"scratch_kin_scan_done_total_{scan_total}_"
                f"bad_{len(scan_bad)}_trigger_in_scratch_"
                f"{EXPECTED_TRIGGER_KIN_PID in scan_ids_in_order}"
            )

            # Stage 2: fire CmdUCINet (split-fire pattern;
            # CmdUCINet not in _TIMER_DISPATCH_SUBS so we
            # CAN'T put it in the chain).
            sess.set_form_tag(spec.name, "CmdUCINet",
                               str(out_path))
            try:
                sess.click_via_timer(
                    spec.name, ctl="CmdUCINet",
                    result_table=None, wait_done=False,
                )
                mark("cmducinet_fired_no_wait")
            except Exception as e:
                mark(f"cmducinet_fire_exc: {e!r}")

            # Poll for the partial file to appear
            # (CmdUCINet writes via FSO synchronously; even
            # if it errors mid-write, the partial file is
            # left behind).
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
                text = raw.decode("cp1252", errors="strict")
                parsed = _parse_partial_vna(text)
                result["partial_file_parse"] = parsed
                mark("partial_file_parsed")

                # Identify the row immediately after the
                # last successfully-written *node properties*
                # row (in c_kin_id ASC iteration order).
                np_section = next(
                    (s for s in parsed.get("sections", [])
                     if s.get("marker") == "*node properties"),
                    None)
                if np_section and np_section.get(
                        "last_row_first_token"):
                    try:
                        last_kid = int(np_section[
                            "last_row_first_token"])
                        if last_kid in scan_ids_in_order:
                            idx = scan_ids_in_order.index(
                                last_kid)
                            next_idx = idx + 1
                            if next_idx < len(
                                    scan_ids_in_order):
                                next_kid = scan_ids_in_order[
                                    next_idx]
                                cur = sess.conn.cursor()
                                cur.execute(
                                    "SELECT c_kin_name FROM "
                                    "ZZ_SCRATCH_KIN "
                                    "WHERE c_kin_id = ?",
                                    next_kid)
                                row = cur.fetchone()
                                cur.close()
                                next_name = (row[0] or ""
                                             if row else "")
                                next_bad = (
                                    _non_cp1252_no_substitute_chars(
                                        next_name))
                                result[
                                    "row_after_last_written"
                                ] = {
                                    "last_written_c_kin_id": (
                                        last_kid),
                                    "last_written_kid_index_0based": (
                                        idx),
                                    "next_c_kin_id_in_kid_asc_order": (
                                        next_kid),
                                    "next_kid_index_0based": (
                                        next_idx),
                                    "next_c_kin_name": next_name,
                                    "next_c_kin_name_non_cp1252": [
                                        {"offset": o,
                                         "char": ch,
                                         "codepoint_hex": (
                                             f"U+{cp:04X}")}
                                        for o, ch, cp in
                                        next_bad
                                    ],
                                    "next_c_kin_name_has_non_cp1252_no_substitute": (
                                        len(next_bad) > 0),
                                }
                                mark(
                                    "row_after_last_written"
                                    "_captured")
                    except Exception as e:
                        mark(f"row_after_last_lookup_fail: "
                             f"{e!r}")

            # Capture debug transcript
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
                if "LookAtKinship:ERR" in d["msg"]]
            result["err_messages"] = err_msgs
            if err_msgs:
                if "invalid procedure call" in (
                        " | ".join(err_msgs).lower()):
                    result["outcome"] = (
                        "reproduced_invalid_procedure_call")
                else:
                    result["outcome"] = (
                        f"err_other: {err_msgs[0][:80]}")
            elif file_appeared:
                # File appeared with no ERR — would mean
                # CmdUCINet actually completed.  This is
                # surprising IF the scan found trigger rows
                # in ZZ_SCRATCH_KIN.
                result["outcome"] = (
                    "no_err_file_appeared")
            else:
                result["outcome"] = "no_err_no_file"

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
    """Apply the brief's three-bucket conclusion logic."""
    err_msgs = result.get("err_messages") or []
    reproduced = any(
        "invalid procedure call" in m.lower()
        for m in err_msgs)
    scan = result.get("scratch_kin_scan") or {}
    n_bad_in_scratch = scan.get(
        "rows_with_non_cp1252_no_substitute_c_kin_name", 0)
    trigger_in_scratch = scan.get(
        "trigger_pid_in_scratch_kin", False)
    next_row = result.get("row_after_last_written") or {}
    next_has_trigger = next_row.get(
        "next_c_kin_name_has_non_cp1252_no_substitute", False)
    parsed = result.get("partial_file_parse") or {}
    section_markers = parsed.get(
        "section_markers_in_order") or []

    # Same-bug-family signature: VBA error 5 + partial file
    # with `*node data` complete + truncated `*node
    # properties` + no `*tie data`.
    has_partial_node_props = (
        "*node properties" in section_markers
        and "*tie data" not in section_markers
        and "*node data" in section_markers)

    # The "same-bug-family" signature has 4 substantive
    # signals.  All 4 being TRUE is the strongest possible
    # confirmation, but several COMBINATIONS can also
    # warrant the conclusion:
    #
    #  (i) reproduced AND has_partial_node_props AND
    #      next_has_trigger
    #      -> classic match: bail localised exactly to a
    #         non-substituted non-cp1252 char in the
    #         next-iteration row.
    #
    #  (ii) reproduced AND has_partial_node_props AND
    #       trigger_in_scratch AND n_bad_in_scratch >= 1
    #       AND the partial file's *node properties*
    #       section has 0 written rows (= bail on FIRST
    #       row)
    #       -> when the bail fires on the FIRST WriteLine
    #          attempt, there is no "last written" anchor
    #          to compute "next row" against -- the
    #          iteration-order check is structurally moot.
    #          But the substantive evidence (VBA error 5
    #          + correct file shape + trigger present in
    #          source data + first scan-row IS a trigger)
    #          is just as conclusive.
    np_section = next(
        (s for s in (parsed.get("sections") or [])
         if s.get("marker") == "*node properties"),
        None)
    np_rows = (np_section.get("total_rows", 0)
               if np_section else 0)
    first_scan_row_is_trigger = (
        scan.get("first_bad_index_kid_asc_0based") == 0
        and n_bad_in_scratch >= 1)

    if (reproduced and has_partial_node_props
            and next_has_trigger):
        bucket = "same_bug_family_runtime_confirmed"
        rationale = (
            "Kinship.CmdUCINet reproduces the SAME failure "
            "mode as Issue #22's Associations case: "
            "`LookAtKinship:ERR Invalid procedure call or "
            "argument` (VBA error 5) fires mid-write, "
            "leaving a partial `.vna` file with `*node data` "
            "complete + `*node properties` truncated + "
            "`*tie data` never written.  The row immediately "
            "after the last successfully-written one (in "
            "c_kin_id ASC iteration order) has c_kin_name = "
            f"{next_row.get('next_c_kin_name')!r} containing "
            f"non-cp1252 char(s) at "
            f"{next_row.get('next_c_kin_name_non_cp1252')!r}. "
            "This confirms the canonical-Issue-#22 sibling "
            "risk on Kinship: the current Kinship × CmdUCINet "
            "coverage (`tests/test_vba_cmducinet_kinship."
            "py`) passes only because the chosen fixture "
            "(person 3211) happens to have no Han-name "
            "members in its kin network — switching to a "
            "fixture that DOES (this probe used picker = "
            f"{PICKER_PID}) reproduces the same VBA error 5. "
            "The Issue #22 canonical entry's 'sibling-form "
            "risk' paragraph for Kinship is now runtime-"
            "verified, not just statically inferred."
        )
    elif (reproduced and has_partial_node_props
            and trigger_in_scratch
            and n_bad_in_scratch >= 1
            and np_rows == 0
            and first_scan_row_is_trigger):
        bucket = "same_bug_family_runtime_confirmed"
        rationale = (
            "Kinship.CmdUCINet reproduces the SAME failure "
            "mode as Issue #22's Associations case: "
            "`LookAtKinship:ERR Invalid procedure call or "
            "argument` (VBA error 5) fires mid-write, "
            "leaving a partial `.vna` file with `*node data` "
            "complete + `*node properties` truncated (header "
            "written but ZERO data rows -- bail on the FIRST "
            "WriteLine attempt) + `*tie data` never "
            "written.\n\n"
            "Why the iteration-order 'next row' lookup is "
            "moot here: when *node properties* bails on the "
            "very first row, there is no last-written "
            "anchor to compute the 'next row' against.  But "
            "the substantive evidence is just as conclusive: "
            f"ZZ_SCRATCH_KIN has {n_bad_in_scratch} non-"
            "cp1252-without-substitute c_kin_name row(s); "
            "the FIRST row in c_kin_id ASC order IS a "
            "trigger (c_kin_id "
            f"{(scan.get('samples') or [{}])[0].get('c_kin_id')}, "
            "c_kin_name "
            f"{(scan.get('samples') or [{}])[0].get('c_kin_name')!r} "
            "with non-cp1252 char(s) "
            f"{[(c['char'], c['codepoint_hex']) for c in (scan.get('samples') or [{}])[0].get('non_cp1252_chars', [])]}); "
            "the bail happens on row 0's WriteLine.  "
            "Combined with the matching partial-file shape "
            "and the matching VBA error 5 wording, this is "
            "the same bug class as Issue #22 -- only the "
            "specific row at which the bail fires differs "
            "(Associations bailed at row 3974 of 8087; "
            f"Kinship bailed at row 1 of {scan.get('scanned_rows')}).\n\n"
            "The current Kinship × CmdUCINet coverage "
            "(`tests/test_vba_cmducinet_kinship.py`) passes "
            "only because the chosen fixture (person 3211) "
            "happens to have no Han-name members in its "
            "kin network -- switching to a fixture that DOES "
            f"(this probe used picker = {PICKER_PID}, He "
            "Jing 何淨, whose sole 1-hop kin is pid 140733 "
            "He Mou 取) reproduces the same VBA error 5.  "
            "The Issue #22 canonical entry's 'sibling-form "
            "risk' paragraph for Kinship is now runtime-"
            "verified, not just statically inferred."
        )
    elif (reproduced and has_partial_node_props
            and not next_has_trigger):
        bucket = "still_needs_better_fixture"
        rationale = (
            "VBA error 5 reproduced and the partial file "
            "shape matches Issue #22's pattern, BUT the "
            "row-after-last-written lookup did NOT find a "
            "non-cp1252 trigger char.  Possibilities: (a) "
            "the c_kin_id ASC iteration order assumption is "
            "wrong; (b) the actual trigger is in a "
            "DIFFERENT field than c_kin_name (less likely "
            "given the static read); (c) the lookup itself "
            "had an issue.  Need to refine probe shape "
            "before drawing same-vs-different conclusions."
        )
    elif (not reproduced and trigger_in_scratch
            and not has_partial_node_props):
        bucket = "sibling_risk_not_reproduced_on_current_fixtures"
        rationale = (
            "Trigger row IS in ZZ_SCRATCH_KIN (pid "
            f"{EXPECTED_TRIGGER_KIN_PID} present) and the "
            f"scan found {n_bad_in_scratch} non-cp1252-"
            "without-substitute c_kin_name rows, but "
            "CmdUCINet did NOT raise VBA error 5 and did "
            "NOT leave a partial-file shape.  Possibilities: "
            "(a) CmdUCINet completed cleanly somehow "
            "(would mean the source code differs from what "
            "static read shows, OR the FSO behaviour on "
            "this Windows build is different from "
            "Associations' case — surprising); (b) "
            "CmdUCINet bailed at an earlier guard (e.g. "
            "RecordCount=0 on a subform recordset) before "
            "reaching the WriteLine.  See the probe's "
            "outcome field for the exact failure mode."
        )
    elif not trigger_in_scratch:
        bucket = "still_needs_better_fixture"
        rationale = (
            f"Picker pid {PICKER_PID} was supposed to pull "
            f"trigger pid {EXPECTED_TRIGGER_KIN_PID} into "
            "ZZ_SCRATCH_KIN via 1-hop kin recursion, but "
            "the scan didn't find that pid in ZZ_SCRATCH_"
            "KIN.  Either (a) the kin recursion didn't run "
            "(check CmdRun outcome); (b) the kin link "
            "between picker and trigger was missing/changed "
            "on the current dump; (c) the form's default "
            "TxtNodeDist is 0 and the recursion never "
            "expanded.  Try a different fixture or set "
            "TxtNodeDist explicitly before treating "
            "non-reproduction as evidence of a robust "
            "code path."
        )
    else:
        bucket = "still_needs_better_fixture"
        rationale = (
            f"Outcome `{result.get('outcome')}` does not "
            "match any clean conclusion bucket.  See "
            "per-result detail (markers / exception / "
            "scan / file shape)."
        )
    return {
        "conclusion": bucket,
        "rationale": rationale,
        "reproduced_invalid_procedure_call": reproduced,
        "trigger_pid_in_scratch_kin": trigger_in_scratch,
        "n_non_cp1252_rows_in_scratch_kin": n_bad_in_scratch,
        "partial_file_matches_issue_22_shape": (
            has_partial_node_props),
        "next_row_has_non_cp1252_no_substitute": (
            next_has_trigger),
    }


def _write_md(result: dict, classification: dict) -> None:
    md: list[str] = []
    md.append("# LookAtKinship × CmdUCINet — sibling-risk "
              "probe (Issue #22 family)")
    md.append("")
    md.append("**Date:** 2026-05-06  ·  **Branch:** "
              "`investigate/kinship-cmducinet-sibling-risk`")
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
    md.append(f"- Picker pid: `{result.get('picker_pid')}` "
              "(He Jing 何淨 — chosen because their sole "
              "1-hop kin row points to pid "
              f"{result.get('expected_trigger_kin_pid')} "
              f"He Mou 取, U+53D6 Han ideograph)")
    rc = result.get("row_counts_after_cmdrun") or {}
    md.append(f"- Row counts after CmdRun: "
              f"`ZZ_SCRATCH_KIN={rc.get('ZZ_SCRATCH_KIN')}` "
              f"`ZZ_SCRATCH_KINNET={rc.get('ZZ_SCRATCH_KINNET')}`")
    md.append(f"- File: `{result.get('file_size')} bytes` "
              f"at `{result.get('file_path')}`")
    md.append("")
    md.append("### Partial file structure")
    md.append("")
    parsed = result.get("partial_file_parse") or {}
    sections = parsed.get("sections") or []
    if sections:
        md.append("| section | header | written rows | "
                  "last row's first token |")
        md.append("|---|---|---:|---|")
        for s in sections:
            md.append(
                f"| `{s.get('marker')}` | "
                f"`{(s.get('header') or '')[:80]}` | "
                f"{s.get('total_rows')} | "
                f"`{s.get('last_row_first_token')}` |"
            )
    else:
        md.append("(no sections parsed — file did not appear "
                  "or was empty)")
    md.append("")
    md.append("### ZZ_SCRATCH_KIN.c_kin_name scan "
              "(non-cp1252 without FSO substitute)")
    md.append("")
    scan = result.get("scratch_kin_scan") or {}
    md.append(f"- Scanned rows: {scan.get('scanned_rows')}")
    md.append(f"- Rows with non-cp1252-without-substitute "
              "c_kin_name: "
              f"{scan.get('rows_with_non_cp1252_no_substitute_c_kin_name')}")
    md.append(f"- Trigger pid "
              f"{result.get('expected_trigger_kin_pid')} "
              "present in ZZ_SCRATCH_KIN: "
              f"**{scan.get('trigger_pid_in_scratch_kin')}**")
    md.append(f"- First bad index (in c_kin_id ASC order, "
              f"0-based): "
              f"{scan.get('first_bad_index_kid_asc_0based')}")
    samples = scan.get("samples") or []
    if samples:
        md.append("")
        md.append("Samples:")
        md.append("")
        md.append("| kid_idx | c_kin_id | c_kin_name | "
                  "non-cp1252 chars |")
        md.append("|---:|---:|---|---|")
        for s in samples[:6]:
            chars_str = ", ".join(
                f"`{c['char']}` ({c['codepoint_hex']})"
                for c in s["non_cp1252_chars"][:3])
            md.append(
                f"| {s['kid_order_index_0based']} | "
                f"{s['c_kin_id']} | `{s['c_kin_name']}` | "
                f"{chars_str} |"
            )
    md.append("")
    md.append("### Row immediately after the last "
              "successfully-written *node properties* row")
    md.append("")
    nr = result.get("row_after_last_written") or {}
    if nr:
        md.append(f"- Last written c_kin_id: "
                  f"`{nr.get('last_written_c_kin_id')}` "
                  f"(index "
                  f"{nr.get('last_written_kid_index_0based')} "
                  "in c_kin_id ASC order)")
        md.append(f"- Next c_kin_id: "
                  f"`{nr.get('next_c_kin_id_in_kid_asc_order')}` "
                  f"(index {nr.get('next_kid_index_0based')})")
        md.append(f"- Next c_kin_name: "
                  f"`{nr.get('next_c_kin_name')!r}`")
        md.append(f"- Next c_kin_name has non-cp1252-without-"
                  "substitute chars: "
                  f"**{nr.get('next_c_kin_name_has_non_cp1252_no_substitute')}**")
        chars = nr.get("next_c_kin_name_non_cp1252") or []
        if chars:
            md.append("- Offending chars in next c_kin_name:")
            for c in chars:
                md.append(f"  - offset {c['offset']}: "
                          f"`{c['char']}` "
                          f"({c['codepoint_hex']})")
    else:
        md.append("(no next-row lookup result; see markers "
                  "/ exception)")
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
              "observations.  Re-verify before relying.")
    md.append("")
    md.append(f"**Conclusion bucket: "
              f"`{classification['conclusion']}`**")
    md.append("")
    md.append(classification["rationale"])
    md.append("")
    if classification["conclusion"] == (
            "same_bug_family_runtime_confirmed"):
        md.append("**Suggested follow-up** (NOT autopiloted "
                  "— each step needs its own brief):")
        md.append("")
        md.append("- Tighten Issue #22's canonical text to "
                  "promote Kinship from 'possible sibling "
                  "risk / NOT yet probed' to 'sibling "
                  "risk runtime-confirmed' OR file Issue "
                  "#23 separately as a sibling P1 if the "
                  "maintainer prefers per-form issues.  "
                  "Either way, the static marker test in "
                  "test_known_bugs.py for Bug #22 should "
                  "be extended to also assert the same "
                  "2-arg `CreateTextFile(tFileName, True)` "
                  "pattern in `Form_LookAtKinship.vb` "
                  "(line ~2510).")
        md.append("- Decide whether the existing Kinship × "
                  "CmdUCINet coverage test is still safe "
                  "(currently it passes only because "
                  "person 3211's network has no Han names) "
                  "or whether it should be either (a) "
                  "augmented with a fixture variant that "
                  "DOES include Han names + asserts the "
                  "expected ERR (turning it into a "
                  "bug-pin like test_bug21 / test_bug22), "
                  "or (b) marked as fixture-fragile in "
                  "the docstring + inventory.")
        md.append("- Coordinate the upstream `CreateTextFile"
                  "(..., True, True)` fix across BOTH "
                  "forms in the same upstream patch (per "
                  "the Issue #22 fix recommendation).")
    md.append("")
    md.append("## Constraints honoured per brief")
    md.append("")
    md.append("- ✅ Investigation artifacts only — no "
              "`tests/` or `cbdb_driver/*` changes")
    md.append("- ✅ Did NOT modify README / canonical "
              "reports / issue severity")
    md.append("- ✅ Did NOT do an upstream fix")
    md.append("- ✅ Used Access COM via `VbaSession."
              "make_fixture`")
    md.append("- ✅ Probe-observed facts vs inferences "
              "explicitly separated")
    md.append("- ✅ Did NOT relax standards just because "
              "Kinship currently has coverage")
    OUT_MD.write_text("\n".join(md), encoding="utf-8")


def main() -> int:
    # Cheap re-classification path (no COM re-run).  Same
    # pattern as analysis/probe_groupdata_cmdneo4j_tail.py
    # uses: read the existing JSON's `result` field, run
    # _classify against it, write fresh .md/.json.  Useful
    # when the classifier is updated but the underlying
    # probe data is still valid.
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
                "investigate/kinship-cmducinet-sibling-risk"),
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

    print("=== Kinship CmdUCINet sibling-risk probe ===\n")
    _kill_orphan()
    time.sleep(1)
    result = _run_probe()
    classification = _classify(result)

    out = {
        "schema_version": 1,
        "generated_date": "2026-05-06",
        "probe_branch": (
            "investigate/kinship-cmducinet-sibling-risk"),
        "follow_up_to": (
            "Issue #22 (canonical, merged via PR "
            "chore/file-issue-22-cmducinet-fso-ansi commit "
            "60d9733) -- specifically the 'sibling-form "
            "risk: NOT yet probed' paragraph for "
            "LookAtKinship.CmdUCINet"),
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
    print(f"  reproduced_invalid_procedure_call: "
          f"{classification['reproduced_invalid_procedure_call']}")
    print(f"  trigger_pid_in_scratch_kin: "
          f"{classification['trigger_pid_in_scratch_kin']}")
    print(f"  partial_file_matches_issue_22_shape: "
          f"{classification['partial_file_matches_issue_22_shape']}")
    print(f"  next_row_has_non_cp1252_no_substitute: "
          f"{classification['next_row_has_non_cp1252_no_substitute']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
