"""Systematically re-verify every documented issue against the
current dump.  Prompted by Bug #3 turning out to be NOT reproducible
— need to confirm we haven't been over-claiming on the others.

For each issue, ask one or both of:

  - Is the BUG TRIGGER reachable in the UI?  (e.g. the corresponding
    button exists, the field is interactable, the form is on the
    standard navigation path).
  - Does the reported SYMPTOM actually occur given current source
    data + schema?  (e.g. does the sub-form's RecordSource really
    omit the column the bound control wants).

This is a static / SQL verification only — no Access COM.  Bugs
that need behavioural verification (chain dispatch, runtime ERR
banner) are flagged for follow-up.
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

import pyodbc

# Several note strings carry → / ≠ unicode arrows.  On Windows
# cp1252 the default sys.stdout chokes on them when the script is
# run from PowerShell without PYTHONIOENCODING.  Reconfigure once
# at import so a plain `python analysis/reverify_all_issues.py`
# always works.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

ROOT = Path(__file__).resolve().parent.parent
USER_MDB = ROOT / "data" / "CBDB_BJ_User.mdb"
INV = json.loads(
    (ROOT / "analysis" / "dump" / "control_inventory.json")
    .read_text(encoding="utf-8")
)
TABLES = json.loads(
    (ROOT / "analysis" / "dump" / "tables.json")
    .read_text(encoding="utf-8")
)
QUERIES = json.loads(
    (ROOT / "analysis" / "dump" / "queries.json")
    .read_text(encoding="utf-8")
)


def _table_cols(name: str) -> set[str]:
    for t in TABLES:
        if (t.get("name") or "").upper() == name.upper():
            return {(c.get("name") or "").lower()
                    for c in t.get("columns", [])}
    return set()


def _saved_query_sql(name: str) -> str:
    for q in QUERIES:
        if (q.get("name") or "").upper() == name.upper():
            return q.get("sql") or ""
    return ""


def _has_form(name: str) -> bool:
    return any(k.lower() == name.lower() for k in INV)


def _form_controls(name: str) -> dict[str, dict]:
    for k, info in INV.items():
        if k.lower() == name.lower():
            if not isinstance(info, dict):
                return {}
            return {(c.get("name") or "").lower(): c
                    for c in info.get("controls", [])
                    if c.get("name")}
    return {}


def _has_control(form: str, ctl: str) -> bool:
    return ctl.lower() in _form_controls(form)


def _control_source_used(col: str) -> list[tuple[str, str]]:
    """[(form, control_name), ...] for every control whose ControlSource is
    exactly ``col``, across all forms in the dump.  Lets a re-verify check
    decide whether an issue's offending binding STILL EXISTS, instead of
    trusting a stale hardcoded "the control is hidden" note (the binding may
    have been removed or re-pointed in a later build)."""
    hits: list[tuple[str, str]] = []
    for k, info in INV.items():
        if not isinstance(info, dict):
            continue
        for c in info.get("controls", []):
            if str(c.get("control_source") or "").lower() == col.lower():
                hits.append((k, c.get("name") or ""))
    return hits


def main() -> int:
    print("=" * 70)
    print("Issue-by-issue re-verification (UI reachability + symptom)")
    print("=" * 70)

    findings: list[tuple[int, str, str]] = []
    # Each tuple: (bug_id, status, note)
    # status ∈ {"REAL", "LATENT", "DORMANT", "NEEDS_RECHECK", "REVIEW"}

    # ---- Bug #1: alias swap (already marked dormant) -------------
    cur = pyodbc.connect(
        f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};"
        f"DBQ={USER_MDB};", autocommit=True
    ).cursor()
    cur.execute("SELECT COUNT(*) FROM STATUS_DATA "
                "WHERE c_fy_range > 0 AND c_ly_range > 0 "
                "  AND c_fy_range <> c_ly_range")
    n = int(cur.fetchone()[0])
    # The defect is an ALIAS SWAP in View_StatusData: c_fy_range_desc/_chn
    # aliased off YEAR_RANGE_CODES_1 (the c_ly_range join) instead of the
    # plain YEAR_RANGE_CODES.  Re-derive presence from the saved-query SQL —
    # do NOT assume it is still there (it was corrected on build 20260602).
    sql1 = _saved_query_sql("View_StatusData")
    swap_present = bool(re.search(
        r"YEAR_RANGE_CODES_1\s*\.\s*\[?c_range(?:_chn)?\]?\s+AS\s+"
        r"\[?c_fy_range", sql1, re.IGNORECASE))
    if not swap_present:
        findings.append((1, "REVIEW",
                         "alias swap corrected in source — c_fy_range_desc/"
                         "_chn now alias off YEAR_RANGE_CODES (not "
                         "YEAR_RANGE_CODES_1); defect not present this dump"))
    else:
        findings.append((1, "DORMANT" if n == 0 else "REAL",
                         f"alias swap present in View_StatusData; "
                         f"{n} STATUS_DATA rows trigger it"))

    # ---- Bug #2: dao360 — handled by check_vba_refs.py at open ----
    findings.append((2, "REAL", "shipped .mdb has broken dao360 ref; "
                     "driver autopatches it on open"))

    # ---- Bug #3 was removed from the documented ISSUES set on
    # 2026-05-03 — re-verification on the current dump found 0 NULL
    # backfills out of 92,514 rows, AND there's no upstream source-
    # level fix to point at, so per the marker-failure-≠-fix policy
    # it was treated as an early false positive (testing infrastructure
    # / fixture / driver) rather than a CBDB-maintainer bug.  Kept
    # this comment so the bug-id sequence is self-explanatory.

    # ---- Bug #4: LookAtPlace.CmdGIS GISFrame -----------------------
    # The control 'GISFrame' doesn't exist on LookAtPlace, so calling
    # CmdGIS_Click raises 'Object required'.  BUT — Issue #15 says
    # there's no CmdGIS button on LookAtPlace either.  So a real user
    # CAN'T click it.  The bug is LATENT — only triggered if Issue #15
    # is fixed first.
    has_gisframe = _has_control("LookAtPlace", "GISFrame")
    has_gis_button = _has_control("LookAtPlace", "CmdGIS")
    if has_gisframe:
        findings.append((4, "REVIEW", "GISFrame DOES exist now? bug premise wrong"))
    elif not has_gis_button:
        findings.append((4, "LATENT",
                         "GISFrame missing AND CmdGIS button missing; "
                         "user can't trigger the bug from the UI"))
    else:
        findings.append((4, "REAL",
                         "GISFrame missing, CmdGIS button present → "
                         "user clicking GIS gets 'Object required'"))

    # ---- Bug #5: LookAtStatus.CmdPajek ChkIDs + SQL ----------------
    has_chkids = _has_control("LookAtStatus", "ChkIDs")
    has_pajek_button = _has_control("LookAtStatus", "CmdPajek")
    status_body = (ROOT / "analysis/dump/vba/Form_LookAtStatus.vb"
                   ).read_bytes().decode("utf-8")
    status_cols = _table_cols("ZZ_SCRATCH_STATUS")
    bad_sql_cols = [c for c in ("c_person_id", "c_status_id", "c_status_count")
                     if c not in status_cols]
    if "CmdPajek_Click" not in status_body:
        findings.append((5, "REVIEW",
                         "LookAtStatus.CmdPajek_Click handler (and its ChkIDs "
                         "SQL) removed this build — defect gone"))
    elif has_chkids and not bad_sql_cols:
        findings.append((5, "REVIEW", "neither ChkIDs nor bad SQL cols anymore"))
    elif not has_pajek_button:
        findings.append((5, "LATENT",
                         f"ChkIDs missing={not has_chkids}, "
                         f"bad SQL cols still missing={bad_sql_cols}; "
                         "but no Pajek button → user can't trigger"))
    else:
        findings.append((5, "REAL", "Pajek button present, defects remain"))

    # ---- Bug #6: LookAtGroupData.queryEntry — verified behaviorally
    # (test_vba_bug_behaviors.py).  Real.
    body = (ROOT / "analysis/dump/vba/Form_LookAtGroupData.vb").read_bytes().decode("utf-8")
    if "ENTRY_DATA.c_parental_status " in body:
        findings.append((6, "REAL",
                         "buggy `ENTRY_DATA.c_parental_status` still in "
                         "queryEntry SQL; behavioral test confirms ERR"))
    else:
        findings.append((6, "REVIEW", "buggy substring no longer in source"))

    # ---- Bug #7: Place.CmdNeo4j — verified behaviorally ------------
    body7 = (ROOT / "analysis/dump/vba/Form_LookAtPlace.vb").read_bytes().decode("utf-8")
    buggy7 = ("SELECT DISTINCT ZZ_SCRATCH_P_TEXT.c_person_id, "
              "ZZ_SCRATCH_P_TEXT.c_name, ZZ_SCRATCH_P_TEXT.c_name_chn, "
              "ZZ_SCRATCH_P_TEXT.c_index_year")
    has_place_neo4j = _has_control("LookAtPlace", "CmdNeo4j")
    # buggy7 is a substring of the FIXED SELECT too (the fix appends columns
    # after c_index_year), so also check whether the SELECT now projects the
    # dynasty/female columns the row loop reads.
    proj_fixed7 = "DYNASTIES.c_dynasty" in body7
    if buggy7 in body7 and not proj_fixed7 and has_place_neo4j:
        findings.append((7, "REAL",
                         "buggy 4-col SELECT still in source; CmdNeo4j "
                         "button exists; verified ERR via behavioral test"))
    elif proj_fixed7:
        findings.append((7, "REVIEW",
                         "CmdNeo4j SELECT now projects DYNASTIES.c_dynasty / "
                         "c_dynasty_chn / BIOG_MAIN.c_female (the columns the "
                         "row loop reads) — fixed this build"))
    else:
        findings.append((7, "REVIEW",
                         f"premise changed (buggy SELECT={buggy7 in body7}, "
                         f"button={has_place_neo4j})"))

    # ---- Bug #8: Networks.CmdNeo4j — Form_Open hangs in driver -----
    body8 = (ROOT / "analysis/dump/vba/Form_LookAtNetworks.vb"
             ).read_bytes().decode("utf-8")
    has_net_neo4j = _has_control("LookAtNetworks", "CmdNeo4j")
    if "ADDR_CODES.x_coord" in body8:
        findings.append((8, "REVIEW",
                         "CmdNeo4j place SELECT now projects ADDR_CODES.x_coord"
                         " / y_coord (the columns the loop reads) — fixed this "
                         "build"))
    else:
        findings.append((8, "REAL" if has_net_neo4j else "LATENT",
                         f"static auditor confirms buggy SQL; CmdNeo4j "
                         f"button on Networks={has_net_neo4j}; "
                         "behavioral repro blocked by driver Form_Open hang"))

    # ---- Bug #9: Entry.CmdNeo4j — gated by institution rows --------
    # Re-verified 2026-05-04: source-level typo on
    # Form_LookAtEntry.vb:1425 still present, but the entire SaveAs +
    # buggy `With` block sits inside `If tRecDeleted > 0 Then` at
    # line 1389, where tRecDeleted = row count of an `INSERT ... WHERE
    # ZZ_SCRATCH_ENTRY.c_inst_code > 0`.  CmdQuery copies
    # ENTRY_DATA.c_inst_code verbatim into ZZ_SCRATCH_ENTRY, so we
    # can ask the question with a simple SQL pre-image.
    body9 = (ROOT / "analysis/dump/vba/Form_LookAtEntry.vb"
             ).read_bytes().decode("utf-8")
    # The bug was the Institutions block opening tRstInstitutions but using
    # `With tRstAssocCodes` (the already-closed recordset).  `With
    # tRstAssocCodes` also appears legitimately in the Associations block, so
    # the fix marker is the Institutions block now using `With tRstInstitutions`.
    inst_fixed9 = "With tRstInstitutions" in body9
    cur.execute("SELECT COUNT(*) FROM ENTRY_DATA "
                "WHERE c_inst_code > 0")
    n_inst = int(cur.fetchone()[0])
    cur.execute("SELECT COUNT(*) FROM ENTRY_DATA "
                "WHERE c_inst_name_code > 0")
    n_inst_name = int(cur.fetchone()[0])
    if inst_fixed9:
        findings.append((9, "REVIEW",
                         "Institutions block now `Set tRstInstitutions` then "
                         "`With tRstInstitutions` (consistent recordset) — the "
                         "wrong-variable typo is fixed this build."))
    elif n_inst == 0 and n_inst_name == 0:
        findings.append((9, "LATENT",
                         f"`With tRstAssocCodes` typo confirmed at "
                         f"line 1425, but gated unreachable: "
                         f"ENTRY_DATA c_inst_code>0={n_inst}, "
                         f"c_inst_name_code>0={n_inst_name}.  Will "
                         f"re-promote to P1 the moment any future "
                         f"MDB drop introduces inst rows."))
    else:
        findings.append((9, "REAL",
                         f"`With tRstAssocCodes` typo confirmed AND "
                         f"gate is open: ENTRY_DATA c_inst_code>0="
                         f"{n_inst}, c_inst_name_code>0={n_inst_name}."
                         f"  Re-promote to P1."))

    # ---- Bug #10: EVENT_ADDR_2 TxtAddrCHN/TxtAddrPY -----------------
    # Verify ControlSource columns are NOT in View_EventAddrData
    # projection (would-be visible blank).  Also check if the columns
    # happen to be in some underlying table that DAO might still
    # resolve — but for a saved-query RecordSource this isn't possible.
    sql10 = _saved_query_sql("View_EventAddrData")
    has10 = ("c_name_chn" in sql10.lower().split())  # rough
    # More precise: scan for ' c_name_chn,' or ' c_name_chn ' as a token
    proj_tokens = set(re.findall(r"\b([a-z_][\w]*)\b", sql10.lower()))
    # Check the OUTPUT name (what control_source can resolve to).
    # The projection has aliases; what matters is the alias names.
    aliased_names = set(re.findall(
        r"as\s+([a-z_]\w*)", sql10, flags=re.IGNORECASE
    ))
    aliased_names = {n.lower() for n in aliased_names}
    # Plus un-aliased trailing-identifier columns (e.g. plain
    # "EVENTS_ADDR.c_personid" exposes "c_personid").
    unaliased = set()
    for tok in re.findall(
        r"[A-Za-z_]\w*\.([a-z_]\w*)(?=\s*(?:,|FROM|from))", sql10
    ):
        unaliased.add(tok.lower())
    output_names = aliased_names | unaliased
    bad10 = [n for n in ("c_name_chn", "c_name") if n not in output_names]
    # Re-derive whether the offending BINDING still exists, not just whether
    # the view projects the column.  On build 20260602 the EVENT_ADDR_2
    # Subform controls were re-bound from c_name_chn/c_name to the projected
    # aliases (c_event_addr_chn / c_event_addr_name), so no control renders
    # blank even though those column names remain absent from the projection.
    subform10 = _form_controls("EVENT_ADDR_2 Subform")
    still_bound10 = sorted({
        (c.get("control_source") or "").lower()
        for c in subform10.values()
        if (c.get("control_source") or "").lower() in ("c_name_chn", "c_name")
    })
    if not still_bound10:
        findings.append((10, "REVIEW",
                         "EVENT_ADDR_2 Subform controls re-bound to the "
                         "projected aliases (c_event_addr_chn / "
                         "c_event_addr_name); no control binds c_name_chn/"
                         "c_name — blank-column defect not present this dump"))
    elif bad10:
        findings.append((10, "REAL",
                         f"View_EventAddrData projection lacks {bad10} AND "
                         f"controls still bind {still_bound10}; sub-form "
                         f"controls render blank"))
    else:
        findings.append((10, "REVIEW",
                         "view actually exposes c_name_chn / c_name now"))

    # ---- Bug #11/#12: hidden internal controls bound to non-projected
    # columns.  The static defect (control bound to a column that the
    # form's RecordSource doesn't project) IS real for both, but a live
    # COM probe in 2026-05-03 confirmed that the offending controls are
    # `Visible = False` with sub-5mm widths — i.e. hidden internal
    # join-key holders, never user-facing.  Reclass: LATENT, not REAL.
    # See `analysis/probe_bug_10_11_12_visibility.py` for the probe
    # and `analysis/dump/bug_10_11_12_visibility.json` for the result.
    sql11 = _saved_query_sql("View_EventsData")
    aliased11 = {n.lower() for n in re.findall(
        r"as\s+([a-z_]\w*)", sql11, flags=re.IGNORECASE
    )}
    unaliased11 = set()
    for tok in re.findall(
        r"[A-Za-z_]\w*\.([a-z_]\w*)(?=\s*(?:,|FROM|from))", sql11
    ):
        unaliased11.add(tok.lower())
    out11 = aliased11 | unaliased11
    in_evts = "c_event_record_id" in _table_cols("EVENTS_DATA")
    in_view = "c_event_record_id" in out11
    # Re-derive whether a control actually BINDS the column before judging it
    # latent — do NOT trust a stale hardcoded "Visible=False/240 twips" note.
    bound11 = _control_source_used("c_event_record_id")
    if not bound11:
        findings.append((11, "REVIEW",
                         "no control binds ControlSource=c_event_record_id "
                         "in the current forms dump — the offending binding "
                         "is gone; defect not present this dump"))
    elif not in_evts and not in_view:
        findings.append((11, "LATENT",
                         f"c_event_record_id is in NEITHER EVENTS_DATA nor "
                         f"View_EventsData (would render blank); bound by "
                         f"{bound11} — confirm the control's visibility with "
                         f"a live COM probe before promoting"))
    else:
        findings.append((11, "REVIEW",
                         f"in EVENTS_DATA={in_evts}, in view projection={in_view}"))

    # Bug #12: same shape — c_appt_type_code is a hidden internal
    # control.  The user-facing appointment-type controls
    # (TxtApptType / TxtApptTypeChn) are bound to c_appt_desc /
    # c_appt_desc_chn, which ARE in View_PostingOfficeData and render
    # correctly.
    sql12 = _saved_query_sql("View_PostingOfficeData")
    aliased12 = {n.lower() for n in re.findall(
        r"as\s+([a-z_]\w*)", sql12, flags=re.IGNORECASE
    )}
    unaliased12 = set()
    for tok in re.findall(
        r"[A-Za-z_]\w*\.([a-z_]\w*)(?=\s*(?:,|FROM|from))", sql12
    ):
        unaliased12.add(tok.lower())
    out12 = aliased12 | unaliased12
    # Re-derive whether a control actually BINDS the column (the user-facing
    # TxtApptType / TxtApptTypeChn bind c_appt_desc / c_appt_desc_chn, which
    # ARE projected and work) — do NOT trust a stale "Visible=False/180 twips".
    bound12 = _control_source_used("c_appt_type_code")
    if not bound12:
        findings.append((12, "REVIEW",
                         "no control binds ControlSource=c_appt_type_code "
                         "in the current forms dump — the offending binding "
                         "is gone; defect not present this dump"))
    elif "c_appt_type_code" not in out12:
        findings.append((12, "LATENT",
                         f"c_appt_type_code not in View_PostingOfficeData "
                         f"projection (would render blank); bound by "
                         f"{bound12} — confirm the control's visibility with "
                         f"a live COM probe before promoting"))
    else:
        findings.append((12, "REVIEW",
                         "view now exposes c_appt_type_code"))

    # ---- Bug #13/14: cross-form picker refs to missing forms -------
    # The static defect is "the picker form doesn't exist in
    # inventory".  But that only matters if a user can actually reach
    # the button that calls OpenForm.  For #13 the host sub-form
    # (BIOG_MAIN_2_Subform) is reachable via CBDB_Browser_2; for #14
    # the host sub-form (KIN_DATA Subform — note the variant *with
    # the picker button*) is not currently embedded in any active
    # form (BIOG_MAIN_2 embeds KIN_DATA_2 Subform instead, which has
    # no CmdPickKinRel button).  So #14 is LATENT until the host is
    # re-embedded.
    def _is_subform_reachable(needle: str) -> bool:
        """Does any form in inventory contain `needle` as a sub-form
        control (by control name or source_object)?  TMPCLP* design-
        time backups are excluded — they aren't user-navigable."""
        for fname, info in INV.items():
            if not isinstance(info, dict):
                continue
            if fname.startswith("Form__TMPCLP"):
                continue
            for c in info.get("controls", []):
                if (c.get("name") == needle
                        or c.get("source_object") == needle):
                    return True
        return False

    has_nh = _has_form("frmPickNIAN_HAO")
    has_kc = _has_form("frmPickKINSHIP_CODES")

    biog_main_2_reachable = _is_subform_reachable("BIOG_MAIN_2_Subform")
    # The defect only exists if the host's CODE still calls the missing picker.
    _b2 = ROOT / "analysis/dump/vba/Form_BIOG_MAIN_2_Subform.vb"
    calls_nh = (_b2.exists()
                and "frmPickNIAN_HAO" in _b2.read_bytes().decode("utf-8", "replace"))
    if not calls_nh:
        findings.append((13, "REVIEW",
                         "no code in BIOG_MAIN_2_Subform references "
                         "frmPickNIAN_HAO this build — defect gone"))
    else:
        findings.append(
            (13, "REAL" if (not has_nh and biog_main_2_reachable) else
             ("LATENT" if not has_nh else "REVIEW"),
             f"frmPickNIAN_HAO in inventory={has_nh}, "
             f"BIOG_MAIN_2_Subform reachable={biog_main_2_reachable}")
        )

    kin_subform_reachable = _is_subform_reachable("KIN_DATA Subform")
    # The buggy host form 'KIN_DATA Subform' (the variant WITH CmdPickKinRel)
    # may have been removed from the build entirely (only KIN_DATA_2 remains).
    _kh = ROOT / "analysis/dump/vba/Form_KIN_DATA_Subform.vb"
    if not _kh.exists():
        findings.append((14, "REVIEW",
                         "KIN_DATA Subform host (with CmdPickKinRel) removed "
                         "from this build — defect gone"))
    else:
        findings.append(
            (14, "REAL" if (not has_kc and kin_subform_reachable) else
             ("LATENT" if not has_kc else "REVIEW"),
             f"frmPickKINSHIP_CODES in inventory={has_kc}, "
             f"KIN_DATA Subform (host of CmdPickKinRel) "
             f"reachable={kin_subform_reachable}")
        )

    # ---- Bug #15-#19: design-time, sub exists but no UI button ----
    cases15 = [
        (15, "LookAtPlace", "CmdGIS"),
        (16, "LookAtStatus", "CmdPajek"),
        (17, "LookAtStatus", "CmdGephi"),
        (18, "LookAtStatus", "CmdUCINet"),
        (19, "LookAtOffice", "CmdGUESS"),
    ]
    for bug_id, form, btn in cases15:
        sub_exists = (
            f"Sub {btn}_Click("
            in (ROOT / f"analysis/dump/vba/Form_{form}.vb")
            .read_bytes().decode("utf-8")
        )
        button_exists = _has_control(form, btn)
        if sub_exists and not button_exists:
            findings.append((bug_id, "REAL",
                             f"{form}.{btn}_Click sub exists, no button"))
        else:
            findings.append((bug_id, "REVIEW",
                             f"sub={sub_exists}, button={button_exists}"))

    # ---- Print results ---------------------------------------------
    by_status: dict[str, list[tuple[int, str]]] = {}
    for bug_id, status, note in sorted(findings):
        by_status.setdefault(status, []).append((bug_id, note))

    for status in ("REAL", "REAL_BUT_GATED", "LATENT", "DORMANT",
                   "NEEDS_RECHECK", "REVIEW"):
        items = by_status.get(status, [])
        if not items:
            continue
        print(f"\n=== {status} ({len(items)} issues) ===")
        for bug_id, note in items:
            print(f"  Bug #{bug_id}: {note}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
