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
    findings.append((1, "DORMANT" if n == 0 else "REAL",
                     f"{n} STATUS_DATA rows trigger the alias swap"))

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
    status_cols = _table_cols("ZZ_SCRATCH_STATUS")
    bad_sql_cols = [c for c in ("c_person_id", "c_status_id", "c_status_count")
                     if c not in status_cols]
    if has_chkids and not bad_sql_cols:
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
    if buggy7 in body7 and has_place_neo4j:
        findings.append((7, "REAL",
                         "buggy 4-col SELECT still in source; CmdNeo4j "
                         "button exists; verified ERR via behavioral test"))
    else:
        findings.append((7, "REVIEW",
                         f"premise changed (buggy SELECT={buggy7 in body7}, "
                         f"button={has_place_neo4j})"))

    # ---- Bug #8: Networks.CmdNeo4j — Form_Open hangs in driver -----
    has_net_neo4j = _has_control("LookAtNetworks", "CmdNeo4j")
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
    typo_present = "With tRstAssocCodes" in body9
    cur.execute("SELECT COUNT(*) FROM ENTRY_DATA "
                "WHERE c_inst_code > 0")
    n_inst = int(cur.fetchone()[0])
    cur.execute("SELECT COUNT(*) FROM ENTRY_DATA "
                "WHERE c_inst_name_code > 0")
    n_inst_name = int(cur.fetchone()[0])
    if not typo_present:
        findings.append((9, "REVIEW",
                         "`With tRstAssocCodes` typo no longer in "
                         "Form_LookAtEntry.vb — flip this branch."))
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
    import re as _re
    proj_tokens = set(_re.findall(r"\b([a-z_][\w]*)\b", sql10.lower()))
    # Check the OUTPUT name (what control_source can resolve to).
    # The projection has aliases; what matters is the alias names.
    aliased_names = set(_re.findall(
        r"as\s+([a-z_]\w*)", sql10, flags=_re.IGNORECASE
    ))
    aliased_names = {n.lower() for n in aliased_names}
    # Plus un-aliased trailing-identifier columns (e.g. plain
    # "EVENTS_ADDR.c_personid" exposes "c_personid").
    unaliased = set()
    for tok in _re.findall(
        r"[A-Za-z_]\w*\.([a-z_]\w*)(?=\s*(?:,|FROM|from))", sql10
    ):
        unaliased.add(tok.lower())
    output_names = aliased_names | unaliased
    bad10 = [n for n in ("c_name_chn", "c_name") if n not in output_names]
    if bad10:
        findings.append((10, "REAL",
                         f"View_EventAddrData projection lacks {bad10}; "
                         f"sub-form controls bound to those names show blank"))
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
    aliased11 = {n.lower() for n in _re.findall(
        r"as\s+([a-z_]\w*)", sql11, flags=_re.IGNORECASE
    )}
    unaliased11 = set()
    for tok in _re.findall(
        r"[A-Za-z_]\w*\.([a-z_]\w*)(?=\s*(?:,|FROM|from))", sql11
    ):
        unaliased11.add(tok.lower())
    out11 = aliased11 | unaliased11
    in_evts = "c_event_record_id" in _table_cols("EVENTS_DATA")
    in_view = "c_event_record_id" in out11
    if not in_evts and not in_view:
        findings.append((11, "LATENT",
                         "c_event_record_id is in NEITHER EVENTS_DATA "
                         "nor View_EventsData → would render blank, BUT "
                         "the control is Visible=False (hidden internal "
                         "control, width=240 twips per COM probe) — user "
                         "doesn't see it"))
    else:
        findings.append((11, "REVIEW",
                         f"in EVENTS_DATA={in_evts}, in view projection={in_view}"))

    # Bug #12: same shape — c_appt_type_code is a hidden internal
    # control.  The user-facing appointment-type controls
    # (TxtApptType / TxtApptTypeChn) are bound to c_appt_desc /
    # c_appt_desc_chn, which ARE in View_PostingOfficeData and render
    # correctly.
    sql12 = _saved_query_sql("View_PostingOfficeData")
    aliased12 = {n.lower() for n in _re.findall(
        r"as\s+([a-z_]\w*)", sql12, flags=_re.IGNORECASE
    )}
    unaliased12 = set()
    for tok in _re.findall(
        r"[A-Za-z_]\w*\.([a-z_]\w*)(?=\s*(?:,|FROM|from))", sql12
    ):
        unaliased12.add(tok.lower())
    out12 = aliased12 | unaliased12
    if "c_appt_type_code" not in out12:
        findings.append((12, "LATENT",
                         "c_appt_type_code not in View_PostingOfficeData "
                         "projection → would render blank, BUT the "
                         "control is Visible=False (hidden internal "
                         "control, width=180 twips per COM probe).  "
                         "The user-facing TxtApptType / TxtApptTypeChn "
                         "are bound to c_appt_desc / c_appt_desc_chn "
                         "which ARE in projection and work correctly."))
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
    findings.append(
        (13, "REAL" if (not has_nh and biog_main_2_reachable) else
         ("LATENT" if not has_nh else "REVIEW"),
         f"frmPickNIAN_HAO in inventory={has_nh}, "
         f"BIOG_MAIN_2_Subform reachable={biog_main_2_reachable}")
    )

    kin_subform_reachable = _is_subform_reachable("KIN_DATA Subform")
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
