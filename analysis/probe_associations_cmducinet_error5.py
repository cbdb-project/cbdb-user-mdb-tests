"""Focused investigation of the LookAtAssociations CmdUCINet
mid-write VBA error 5.

Background
----------
The CmdUCINet family probe (`investigate/cmducinet-family-
shape`, commit 4e8e0d2 on main) found that
`Form_LookAtAssociations.CmdUCINet_Click` bails mid-write with
`LookAtAssociations:ERR Invalid procedure call or argument`
(VBA error 5) after writing 3973 of 8087 expected rows in the
`*node properties` section.  That blocks Associations from the
CmdUCINet family coverage that LookAtKinship just landed
(`cover/kinship-cmducinet-3211`, commit fc0dc13).

Goal
----
Localize the error to a specific section / row class / field
class so the maintainer can decide whether this is:

  A. a NEW bug candidate (e.g. unguarded encoding mismatch
     between FSO ASCII writer and Unicode source data) that
     should be canonicalized as Issue #N
  B. a fixture-/data-specific blocker (specific to person
     437's network on the current dump; would not affect a
     different fixture)
  C. still ambiguous (need follow-up probe shape)

Strong static-source candidate hypothesis (to test)
---------------------------------------------------
`Form_LookAtAssociations.vb`'s `*node properties` block
(inside `CmdUCINet_Click`) writes:

    If IsNull(!c_name) Then
        tStr = tStr + "[Missing]"
    Else
        tStr = tStr + tQuote + !c_name + tQuote
    End If
    tVNA.WriteLine (tStr)

`tVNA` is created via
`Scripting.FileSystemObject.CreateTextFile(tFileName, True)`.
The 2nd argument is `Overwrite=True`; the 3rd argument
(`Unicode`) is OMITTED so it defaults to FALSE → the file is
written in the system default ANSI code page (cp1252 on
Windows-EN).  If `!c_name` contains a character outside
cp1252 (for example: Pinyin diacritics like `ǎ`, `ǐ`, `ǒ`, or
any U+0100+ char), `tVNA.WriteLine` raises VBA error 5
("Invalid procedure call or argument").

Static evidence that this is plausible:
  - The *node data* section (which writes c_person_id /
    c_index_year / c_sex / c_x_coord / c_y_coord — all
    numeric / single-char ASCII) completed all 8087 rows.
  - Only the *node properties* section (which writes
    c_name = the historical figure's romanized name) failed
    mid-write.
  - This suggests the failure is field-specific to c_name
    rather than recordset-wide.

What this probe does
--------------------
Single COM session, in order:

1. Open Associations + apply matrix `assoc_437_unfiltered`
   fixture; fire CmdQuery via timer (wait for autodetect
   DONE marker).
2. Capture row count of `ZZ_SCRATCH_P_ASSOC` (= the
   recordset CmdUCINet's `*node properties` walks).
3. Scan `ZZ_SCRATCH_P_ASSOC.c_name` for non-cp1252-encodable
   characters; record:
     - count of rows whose c_name contains any non-cp1252
       char
     - first such row's c_person_id + c_name + offending
       characters
4. Fire CmdUCINet via separate timer with `wait_done=False`
   + file-poll; wait for the partial file to appear (it will
   even though the body errors).
5. Parse the partial .vna file to extract the LAST
   successfully-written *node properties* row's
   c_person_id.
6. Query the ZZ_SCRATCH_P_ASSOC recordset in the same
   iteration order CmdUCINet would use (default `OpenRecordset
   ("ZZ_SCRATCH_P_ASSOC", dbOpenDynaset)` → JET internal /
   primary-key order, typically by c_person_id ASC).
7. Find the row IMMEDIATELY AFTER the last successful one
   in that iteration order.  Inspect its c_name for non-
   cp1252 characters.

If step 7's "next row" has a non-cp1252 c_name, the
hypothesis is confirmed and the bug class is "FSO ASCII
writer chokes on non-cp1252 source data" — a NEW bug
candidate (option A above).

Outputs
-------
- analysis/probe_associations_cmducinet_error5.md
- reports/probe_associations_cmducinet_error5.json
"""
from __future__ import annotations

import json
import re
import subprocess
import sys
import threading
import time
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))

USER_MDB = ROOT / "data" / "CBDB_BJ_User.mdb"
WORK = ROOT / "analysis" / "_probe_assoc_cmducinet_err5_copy.mdb"
OUT_JSON = ROOT / "reports" / "probe_associations_cmducinet_error5.json"
OUT_MD = ROOT / "analysis" / "probe_associations_cmducinet_error5.md"

PER_PROBE_OUTER_TIMEOUT_SEC = 360
CMDQUERY_TIMER_TIMEOUT_SEC = 180
CMDUCINET_FILE_POLL_TIMEOUT_SEC = 60


def _kill_orphan():
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "MSACCESS.EXE"],
            capture_output=True, timeout=10)
    except Exception:
        pass


def _matrix_fixture_assoc():
    from test_vba_matrix_all_forms import _all_fixtures
    for fx in _all_fixtures():
        if fx.spec.name == "LookAtAssociations":
            return fx
    return None


def _non_cp1252_char_offsets(s: str) -> list[tuple[int, str, int]]:
    """Return list of (offset, char, codepoint) for chars in
    `s` that cannot be encoded in cp1252.  An empty list means
    `s` is fully cp1252-encodable."""
    bad = []
    for i, ch in enumerate(s):
        try:
            ch.encode("cp1252", errors="strict")
        except UnicodeEncodeError:
            bad.append((i, ch, ord(ch)))
    return bad


def _parse_partial_vna(text: str) -> dict:
    """Parse a partial CmdUCINet .vna file.  Find the last
    successfully-written row in each section so we can identify
    where the body bailed.  Returns:
      {
        sections: [
          {marker, header, last_row, last_row_n, total_rows},
          ...
        ],
        last_section_marker, last_row_text, last_row_first_token
      }"""
    lines = [ln for ln in text.replace("\r\n", "\n").split("\n")]
    sections: list[dict] = []
    cur: dict | None = None
    for ln in lines:
        s = ln.strip()
        if s.startswith("*"):
            if cur is not None:
                sections.append(cur)
            cur = {
                "marker": s,
                "header": None,
                "rows": [],
            }
            continue
        if cur is not None and s:
            if cur["header"] is None:
                cur["header"] = s
            else:
                cur["rows"].append(s)
    if cur is not None:
        sections.append(cur)
    out_sections = []
    for sec in sections:
        rows = sec["rows"]
        last = rows[-1] if rows else None
        out_sections.append({
            "marker": sec["marker"],
            "header": sec["header"],
            "total_rows": len(rows),
            "last_row": last,
            "last_row_first_token": (
                last.split()[0] if last else None),
        })
    last_section = out_sections[-1] if out_sections else {}
    return {
        "sections": out_sections,
        "last_section_marker": last_section.get("marker"),
        "last_row_text": last_section.get("last_row"),
        "last_row_first_token": last_section.get(
            "last_row_first_token"),
    }


def _run_probe() -> dict:
    """Execute the probe as a single COM session."""
    from cbdb_driver.vba_session import make_fixture
    from cbdb_driver.form_specs import LOOKATASSOCIATIONS

    spec = LOOKATASSOCIATIONS
    out_dir = ROOT / "analysis" / "_probe_assoc_cmducinet_err5_out"
    if out_dir.exists():
        for f in out_dir.glob("*"):
            try:
                f.unlink()
            except Exception:
                pass
    else:
        out_dir.mkdir(parents=True)
    out_path = out_dir / "cmducinet_associations.vna"

    result: dict = {
        "form": "LookAtAssociations",
        "fixture_name": None,
        "fixture_picker_ids": None,
        "markers": [],
        "outcome": None,
        "elapsed_sec": None,
        "exception": None,
        "row_counts_after_cmdquery": {},
        "debug_transcript": [],
        "file_path": None,
        "file_size": None,
        "partial_file_parse": None,
        "non_cp1252_scan": None,
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

            fx = _matrix_fixture_assoc()
            if fx is None:
                raise RuntimeError("no matrix fixture")
            result["fixture_name"] = fx.name
            result["fixture_picker_ids"] = list(
                fx.picker_ids or [])

            sess.open_form(spec.name)
            for ctl, val in (fx.controls or {}).items():
                try:
                    sess.set_control(spec.name, ctl, val)
                except Exception as e:
                    print(f"  warn setting {ctl}={val!r}: {e}")
            if fx.picker_ids and spec.picker_table:
                sess.set_picker_codes(
                    spec.picker_table, fx.picker_ids,
                    column=spec.picker_column)
            mark("form_seeded")

            # Stage 1: fire CmdQuery via timer
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

            cur = sess.conn.cursor()
            cur.execute(
                "SELECT COUNT(*) FROM ZZ_SCRATCH_P_ASSOC")
            n_p_assoc = int(cur.fetchone()[0] or 0)
            result["row_counts_after_cmdquery"][
                "ZZ_SCRATCH_P_ASSOC"] = n_p_assoc
            cur.execute(
                "SELECT COUNT(*) FROM ZZ_SCRATCH_ASSOC")
            n_assoc = int(cur.fetchone()[0] or 0)
            result["row_counts_after_cmdquery"][
                "ZZ_SCRATCH_ASSOC"] = n_assoc
            cur.close()
            mark(f"row_counts_p_assoc_{n_p_assoc}_assoc_"
                 f"{n_assoc}")

            # Hypothesis pre-test: scan c_name for non-cp1252
            # chars BEFORE we fire CmdUCINet.  This way we
            # know the answer regardless of whether the
            # CmdUCINet fire itself reproduces the error.
            cur = sess.conn.cursor()
            cur.execute(
                "SELECT c_person_id, c_name "
                "FROM ZZ_SCRATCH_P_ASSOC "
                "ORDER BY c_person_id ASC"
            )
            scan_total = 0
            scan_bad = []
            scan_first_bad_idx_in_pid_order = None
            scan_pids_in_order = []
            for idx, row in enumerate(cur.fetchall()):
                scan_total += 1
                pid = int(row[0])
                name = row[1] or ""
                scan_pids_in_order.append(pid)
                bad = _non_cp1252_char_offsets(name)
                if bad:
                    scan_bad.append({
                        "pid_order_index_0based": idx,
                        "c_person_id": pid,
                        "c_name": name,
                        "non_cp1252_chars": [
                            {"offset": o, "char": ch,
                             "codepoint_hex": f"U+{cp:04X}"}
                            for o, ch, cp in bad],
                    })
                    if scan_first_bad_idx_in_pid_order is None:
                        scan_first_bad_idx_in_pid_order = idx
            cur.close()
            result["non_cp1252_scan"] = {
                "scanned_rows": scan_total,
                "rows_with_non_cp1252_c_name": len(scan_bad),
                "first_bad_index_pid_asc_0based": (
                    scan_first_bad_idx_in_pid_order),
                "samples": scan_bad[:10],
            }
            mark(f"non_cp1252_scan_done_total_{scan_total}_bad"
                 f"_{len(scan_bad)}")

            # Stage 2: fire CmdUCINet
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

            # Poll for partial file (CmdUCINet errors but
            # still leaves the FSO-created file with all
            # written-so-far data).
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
                # Decode via cp1252 (FSO default).  Any bytes
                # the writer DID write are by definition
                # cp1252-encodable.
                text = raw.decode("cp1252", errors="strict")
                parsed = _parse_partial_vna(text)
                result["partial_file_parse"] = parsed
                mark("partial_file_parsed")

                # Identify the row immediately after the last
                # successfully-written one, in the c_person_id-
                # ASC iteration order our scan used.  If the
                # bail point is ALSO indexed by c_person_id,
                # this row is what CmdUCINet was about to
                # write when the error fired.
                last_pid_str = parsed.get("last_row_first_token")
                if last_pid_str:
                    try:
                        last_pid = int(last_pid_str)
                        if last_pid in scan_pids_in_order:
                            idx = scan_pids_in_order.index(
                                last_pid)
                            next_idx = idx + 1
                            if next_idx < len(
                                    scan_pids_in_order):
                                next_pid = scan_pids_in_order[
                                    next_idx]
                                # Look up that row's c_name
                                cur = sess.conn.cursor()
                                cur.execute(
                                    "SELECT c_name FROM "
                                    "ZZ_SCRATCH_P_ASSOC "
                                    "WHERE c_person_id = ?",
                                    next_pid)
                                row = cur.fetchone()
                                cur.close()
                                next_name = (row[0] or ""
                                             if row else "")
                                next_bad = (
                                    _non_cp1252_char_offsets(
                                        next_name))
                                result[
                                    "row_after_last_written"
                                ] = {
                                    "last_written_pid": last_pid,
                                    "last_written_pid_index_0based": idx,
                                    "next_pid_in_pid_asc_order": next_pid,
                                    "next_pid_index_0based": next_idx,
                                    "next_c_name": next_name,
                                    "next_c_name_non_cp1252": [
                                        {"offset": o,
                                         "char": ch,
                                         "codepoint_hex": (
                                             f"U+{cp:04X}")}
                                        for o, ch, cp in
                                        next_bad
                                    ],
                                    "next_c_name_has_non_cp1252": (
                                        len(next_bad) > 0),
                                }
                                mark(
                                    "row_after_last_written_"
                                    "captured")
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

            err_msgs = [d["msg"]
                        for d in result["debug_transcript"]
                        if "LookAtAssociations:ERR" in d["msg"]]
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
                result["outcome"] = (
                    "no_err_but_partial_or_full_file")
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
    """Apply the brief's three-bucket conclusion logic.

    A. new_bug_candidate_fso_ascii_writer_vs_unicode_source
       - Reproduced the error AND
       - The row immediately after the last successfully-
         written one (in c_person_id ASC order) has a c_name
         containing chars outside cp1252.
    B. fixture_or_data_specific_blocker
       - Reproduced the error AND
       - No non-cp1252 chars found in any c_name (would mean
         the cause is something other than encoding, AND
         hypothesis-specific to person 437's network).  In
         this branch, we'd recommend trying a different
         Associations fixture before drawing a wider
         conclusion.
    C. still_ambiguous
       - Anything else (didn't reproduce, file didn't
         appear, can't identify the next row, etc.).
    """
    err_msgs = result.get("err_messages") or []
    reproduced = any(
        "invalid procedure call" in m.lower()
        for m in err_msgs)
    scan = result.get("non_cp1252_scan") or {}
    n_bad_overall = scan.get(
        "rows_with_non_cp1252_c_name", 0)
    next_row = result.get("row_after_last_written") or {}
    next_has_non_cp1252 = next_row.get(
        "next_c_name_has_non_cp1252", False)

    if reproduced and next_has_non_cp1252:
        bucket = (
            "A_new_bug_candidate_fso_ascii_writer_vs_"
            "unicode_source")
        rationale = (
            "VBA error 5 reproduced.  Source-data scan found "
            f"{n_bad_overall} rows in ZZ_SCRATCH_P_ASSOC "
            "whose c_name contains non-cp1252 characters.  "
            "The row immediately after the last successfully-"
            "written one (in c_person_id ASC iteration order) "
            f"has c_name = {next_row.get('next_c_name')!r} "
            "with non-cp1252 chars at "
            f"{next_row.get('next_c_name_non_cp1252')!r}.  "
            "This matches the hypothesis exactly: "
            "Form_LookAtAssociations.CmdUCINet_Click writes "
            "via Scripting.FileSystemObject.CreateTextFile "
            "with the Unicode flag omitted (defaults FALSE -> "
            "cp1252 ANSI), and tVNA.WriteLine raises VBA "
            "error 5 when the string contains characters "
            "outside cp1252.  Bug class candidate: a real "
            "CBDB-source defect (the export should either use "
            "the Unicode flag = TRUE, or strip / transliterate "
            "non-cp1252 chars in c_name before WriteLine).  "
            "Recommend filing as a P1 visible-crash issue if "
            "the maintainer confirms this is user-reachable."
        )
    elif reproduced and n_bad_overall == 0:
        bucket = "B_fixture_or_data_specific_blocker"
        rationale = (
            "VBA error 5 reproduced, but ZERO rows in "
            "ZZ_SCRATCH_P_ASSOC have non-cp1252 c_name "
            "characters.  The encoding hypothesis is "
            "FALSIFIED on this fixture; the actual cause "
            "is something else (not encoding-related).  "
            "Recommend either (1) trying a different "
            "Associations fixture to see if the failure "
            "still reproduces, or (2) a per-block isolation "
            "probe to localize the failure to a specific "
            "field expression rather than just the section."
        )
    elif reproduced and not next_has_non_cp1252:
        bucket = "C_still_ambiguous"
        rationale = (
            "VBA error 5 reproduced and there ARE rows with "
            f"non-cp1252 c_name globally ({n_bad_overall}), "
            "but the row immediately after the last "
            "successfully-written one (in c_person_id ASC "
            "order) does NOT contain non-cp1252 chars.  "
            "Possibilities: (a) iteration order is not "
            "c_person_id ASC; (b) failure is in a different "
            "field than c_name; (c) failure is data-shape "
            "rather than encoding.  Need to refine probe."
        )
    else:
        bucket = "C_still_ambiguous"
        rationale = (
            f"Did not reproduce or insufficient data.  "
            f"outcome={result.get('outcome')}.  "
            f"err_msgs={err_msgs[:2]}"
        )
    return {"conclusion": bucket, "rationale": rationale,
            "reproduced": reproduced,
            "non_cp1252_rows_in_scratch": n_bad_overall,
            "next_row_has_non_cp1252": next_has_non_cp1252}


def _write_md(result: dict, classification: dict) -> None:
    md: list[str] = []
    md.append("# LookAtAssociations × CmdUCINet — error 5 "
              "localisation probe")
    md.append("")
    md.append("**Date:** 2026-05-06  ·  **Branch:** "
              "`investigate/associations-cmducinet-error5`")
    md.append("")
    md.append("## Headline")
    md.append("")
    md.append("Conclusion: **`"
              f"{classification['conclusion']}`**")
    md.append("")
    md.append(classification["rationale"])
    md.append("")
    md.append("---")
    md.append("")
    md.append("## Probe-observed facts")
    md.append("")
    md.append(f"- Outcome: `{result.get('outcome')}`")
    md.append(f"- Elapsed: {result.get('elapsed_sec')} s")
    md.append(f"- Fixture used: `{result.get('fixture_name')}` "
              f"(picker_ids={result.get('fixture_picker_ids')})")
    rc = result.get("row_counts_after_cmdquery") or {}
    md.append(f"- Row counts after CmdQuery: "
              f"`ZZ_SCRATCH_P_ASSOC = {rc.get('ZZ_SCRATCH_P_ASSOC')}`, "
              f"`ZZ_SCRATCH_ASSOC = {rc.get('ZZ_SCRATCH_ASSOC')}`")
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
    md.append("### Non-cp1252 c_name scan over "
              "ZZ_SCRATCH_P_ASSOC")
    md.append("")
    scan = result.get("non_cp1252_scan") or {}
    md.append(f"- Scanned rows: {scan.get('scanned_rows')}")
    md.append(f"- Rows with non-cp1252 c_name: "
              f"{scan.get('rows_with_non_cp1252_c_name')}")
    md.append(f"- First bad index (in c_person_id ASC "
              f"order, 0-based): "
              f"{scan.get('first_bad_index_pid_asc_0based')}")
    samples = scan.get("samples") or []
    if samples:
        md.append("")
        md.append("First few samples:")
        md.append("")
        md.append("| pid_order_idx | c_person_id | c_name | "
                  "non_cp1252 chars |")
        md.append("|---:|---:|---|---|")
        for s in samples[:6]:
            chars_str = ", ".join(
                f"`{c['char']}` ({c['codepoint_hex']})"
                for c in s["non_cp1252_chars"][:3])
            md.append(
                f"| {s['pid_order_index_0based']} | "
                f"{s['c_person_id']} | `{s['c_name']}` | "
                f"{chars_str} |"
            )
    md.append("")
    md.append("### Row immediately after the last "
              "successfully-written one")
    md.append("")
    nr = result.get("row_after_last_written") or {}
    if nr:
        md.append(f"- Last written pid: "
                  f"`{nr.get('last_written_pid')}` "
                  f"(index {nr.get('last_written_pid_index_0based')} "
                  "in c_person_id ASC order)")
        md.append(f"- Next pid: "
                  f"`{nr.get('next_pid_in_pid_asc_order')}` "
                  f"(index "
                  f"{nr.get('next_pid_index_0based')})")
        md.append(f"- Next c_name: "
                  f"`{nr.get('next_c_name')!r}`")
        md.append(f"- Next c_name has non-cp1252 chars: "
                  f"**{nr.get('next_c_name_has_non_cp1252')}**")
        chars = nr.get("next_c_name_non_cp1252") or []
        if chars:
            md.append("- Offending chars in next c_name:")
            for c in chars:
                md.append(f"  - offset {c['offset']}: "
                          f"`{c['char']}` "
                          f"({c['codepoint_hex']})")
    else:
        md.append("(no next-row lookup result; see markers / "
                  "exception)")
    md.append("")
    md.append("### Per-form ZZ_TEST_DEBUG transcript")
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
    md.append("## Inferences for future canonicalization / "
              "coverage decisions")
    md.append("")
    md.append("These are CONCLUSIONS drawn from the probe + "
              "static evidence, NOT additional probe "
              "observations.  Re-verify before relying.")
    md.append("")
    md.append(f"**Conclusion bucket: `"
              f"{classification['conclusion']}`**")
    md.append("")
    md.append(classification["rationale"])
    md.append("")
    if classification["conclusion"].startswith(
            "A_new_bug_candidate"):
        md.append("**Suggested follow-up** (NOT autopiloted):")
        md.append("")
        md.append("- File a candidate issue for the FSO ASCII "
                  "vs Unicode-source mismatch.  Suggested "
                  "shape: P1 visible crash; affected sub "
                  "`Form_LookAtAssociations.CmdUCINet_Click`; "
                  "fix recommendation = `Set tVNA = "
                  "tFileSystem.CreateTextFile(tFileName, "
                  "True, True)` (3rd arg = Unicode = TRUE, "
                  "writes UTF-16LE) OR strip non-cp1252 "
                  "chars before `tVNA.WriteLine`.  Same "
                  "pattern likely affects the *node "
                  "properties* block in `Form_LookAtKinship."
                  "vb` (the c_name shortlabel writer); "
                  "Kinship's coverage just landed because "
                  "person 3211's network happened to have no "
                  "non-cp1252 c_name values, but a different "
                  "Kinship fixture might surface the same "
                  "crash there.")
        md.append("- A future canonicalization PR should "
                  "include both static marker (grep for the "
                  "missing 3rd arg in CreateTextFile) and "
                  "runtime behavioural pin (drive CmdUCINet "
                  "on a fixture known to contain non-cp1252 "
                  "c_name and assert `:ERR Invalid procedure "
                  "call` reproduces).")
        md.append("- A future driver brief might consider "
                  "patching CmdUCINet to use Unicode mode "
                  "in the test driver so coverage tests "
                  "don't need a non-cp1252-free fixture, "
                  "but that's a workaround masking the "
                  "real CBDB bug — not recommended without "
                  "explicit maintainer authorization.")
    md.append("")
    md.append("## Constraints honoured per brief")
    md.append("")
    md.append("- ✅ Investigation artifacts only — no "
              "`tests/` or `cbdb_driver/*` changes")
    md.append("- ✅ Did NOT modify README / canonical "
              "reports / issue severity")
    md.append("- ✅ Used Access COM via "
              "`VbaSession.make_fixture`")
    md.append("- ✅ Reused matrix-supplied "
              "`assoc_437_unfiltered` fixture; no new "
              "fixture design")
    md.append("- ✅ Did NOT file an issue (this PR is the "
              "evidence base for a maintainer's later "
              "filing decision)")
    OUT_MD.write_text("\n".join(md), encoding="utf-8")


def main() -> int:
    print("=== Associations CmdUCINet error 5 probe ===\n")
    _kill_orphan()
    time.sleep(1)
    result = _run_probe()
    classification = _classify(result)

    out = {
        "schema_version": 1,
        "generated_date": "2026-05-06",
        "probe_branch": (
            "investigate/associations-cmducinet-error5"),
        "follow_up_to": (
            "PR investigate/cmducinet-family-shape "
            "commit 4e8e0d2"),
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
    print(f"\n=== conclusion: {classification['conclusion']} ===")
    print(f"  reproduced: {classification['reproduced']}")
    print(f"  non_cp1252_rows_in_scratch: "
          f"{classification['non_cp1252_rows_in_scratch']}")
    print(f"  next_row_has_non_cp1252: "
          f"{classification['next_row_has_non_cp1252']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
