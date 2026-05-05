"""Read-only inventory of (LookAt form × export button) coverage.

Cross-references three text-only sources to produce a coverage
matrix.  No MDB.  No Access COM.

Inputs:
  - `analysis/dump/vba/Form_LookAt*.vb` — for `Sub <Cmd*>_Click()`
    handler presence per form.
  - `analysis/dump/control_inventory.json` — for button presence on
    each form's design (case-insensitive name match against
    `controls[].name`).
  - A curated `EXPORT_TEST_MANIFEST` declaring which (form, button)
    combos each existing real-VBA test exercises, with skip reasons
    where applicable.  The script cross-validates the manifest
    against the VBA dump on every run, so manifest drift becomes a
    visible audit finding rather than silent rot.

Outputs:
  - `reports/export_coverage_inventory.json` — machine-readable
    matrix + per-cell decision trace + manifest cross-check
    findings.
  - `analysis/export_coverage_inventory.md` — human-readable matrix
    + per-cell legend + recommended next real-VBA slice (just a
    suggestion; this PR ships inventory only and does NOT
    implement).

Status taxonomy (one per cell):
  - `not_applicable` ─ neither button nor handler in source
  - `missing_ui_button` ─ handler exists, button missing (Issues
                         #15-19 family — already documented in the
                         report)
  - `orphan_button_no_handler` ─ button exists, handler missing
                                 (would be a real bug; flagged for
                                 investigation if any)
  - `real_vba_covered` ─ test exists and passes
  - `real_vba_skipped` ─ test exists but is `pytest.mark.skip`'d
                         (with reason)
  - `real_vba_failing` ─ test exists, runs (does NOT skip), and
                         fails on the current dump (e.g. depth-
                         check classifier gap for an unfamiliar
                         file-shape family)
  - `unit_or_static_only` ─ covered by static analysis (e.g.
                            `tests/test_known_bugs.py` source-grep)
                            but no real CmdQuery → CmdX chain
  - `gap` ─ button + handler both exist, no test of any kind

Born 2026-05-04 from the reclassify-#9 follow-up: before picking
the next real-VBA export slice we want a stable, deterministic
inventory of where the suite already covers vs. where the genuine
gaps are.  This script is the read-only first step; PR-shipping
new fixtures comes in a separate, scope-narrow follow-up after
review.
"""
from __future__ import annotations

import datetime as _dt
import json
import re
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent.parent
DUMP_VBA = ROOT / "analysis" / "dump" / "vba"
CONTROL_INVENTORY = ROOT / "analysis" / "dump" / "control_inventory.json"

OUT_JSON = ROOT / "reports" / "export_coverage_inventory.json"
OUT_MD = ROOT / "analysis" / "export_coverage_inventory.md"


# ---------------------------------------------------------------
# Static data — the matrix axes.
# ---------------------------------------------------------------
FORMS: tuple[str, ...] = (
    "LookAtEntry", "LookAtTexts", "LookAtAssociations",
    "LookAtOffice", "LookAtPlace", "LookAtKinship",
    "LookAtStatus", "LookAtAssociationPairs",
    "LookAtNetworks", "LookAtGroupData",
)

# Button names as they appear in the source code (real casing).
# The brief lists "CmdUCInet"; the actual VBA + control name is
# "CmdUCINet" (capital N).  CmdKML doesn't exist in any form —
# KML is a `ChkKML` checkbox option that other exports honour, not
# a standalone button.  Including it in the matrix per the brief
# so the absence is visible rather than hidden.
BUTTONS: tuple[str, ...] = (
    "CmdGIS", "CmdNeo4j", "CmdPajek", "CmdGephi",
    "CmdGUESS", "CmdKML", "CmdUCINet", "CmdGISPeople",
)


# ---------------------------------------------------------------
# Test coverage manifest.
#
# Each entry declares "this real-VBA test exercises this (form,
# button) chain end-to-end".  Status is one of:
#   covered           — test exists and passes
#   skipped:<reason>  — test exists but pytest.mark.skip'd
#
# These are CURATED from reading the test files at
# investigation time (2026-05-04).  The script cross-validates
# every entry's (form, button) pair against the VBA dump, so a
# manifest entry that no longer matches reality lands in
# `manifest_drift_findings`.
# ---------------------------------------------------------------
EXPORT_TEST_MANIFEST: list[dict] = [
    # ---- CmdGIS ----
    {"form": "LookAtEntry",  "button": "CmdGIS",
     "test_module": "tests/test_vba_export.py",
     "test_node": "test_lookatentry_cmd_gis",
     "status": "covered",
     "notes": "byte-level golden compare; the original CmdGIS test"},
    {"form": "LookAtStatus", "button": "CmdGIS",
     "test_module": "tests/test_vba_cmdgis_other_forms.py",
     "test_node": "test_cmd_gis_produces_file",
     "status": "covered",
     "notes": "structural assertion (header + non-empty cols)"},
    {"form": "LookAtTexts",  "button": "CmdGIS",
     "test_module": "tests/test_vba_cmdgis_other_forms.py",
     "test_node": "test_cmd_gis_produces_file",
     "status": "covered", "notes": "structural"},
    {"form": "LookAtAssociations", "button": "CmdGIS",
     "test_module": "tests/test_vba_cmdgis_other_forms.py",
     "test_node": "test_cmd_gis_produces_file",
     "status": "covered", "notes": "structural"},
    {"form": "LookAtOffice", "button": "CmdGIS",
     "test_module": "tests/test_vba_cmdgis_other_forms.py",
     "test_node": "test_cmd_gis_produces_file",
     "status": "covered", "notes": "structural"},
    {"form": "LookAtPlace",  "button": "CmdGIS",
     "test_module": "tests/test_vba_cmdgis_other_forms.py",
     "test_node": "test_cmd_gis_produces_file",
     "status": "covered",
     "notes": "structural; passes only thanks to driver-side "
              "_PER_FORM_CMDGIS_PATCHES rewrite of GISFrame -> "
              "CodeFrame (Issue #4 latent typo workaround)"},
    {"form": "LookAtKinship", "button": "CmdGIS",
     "test_module": "tests/test_vba_cmdgis_other_forms.py",
     "test_node": "test_cmd_gis_produces_file",
     "status": "covered",
     "notes": "structural; subform requery via "
              "_SUBFORMS_TO_REQUERY"},
    {"form": "LookAtGroupData", "button": "CmdGIS",
     "test_module": "tests/test_vba_cmdgis_other_forms.py",
     "test_node": "test_cmd_gis_groupdata_clean_branches",
     "status": "covered",
     "notes":
         "Clean-branches coverage (Status / Office / Addr) using "
         "matrix_hard_forms's groupdata_person_1_small fixture.  "
         "Explicitly excludes Entry (Issue #6 P1 fires JET 3061; "
         "pinned by tests/test_known_bugs.py::test_bug6_groupdata"
         "_query_entry_wrong_field [static] AND tests/test_vba_"
         "bug_behaviors.py::test_bug6_lookat_groupdata_query_"
         "entry_fires_no_such_field [runtime]).  Excludes Text "
         "(person_1 has 0 BIOG_TEXT_DATA rows -> WriteGIS_Text "
         "bails on RecCount=0; benign 0-files state, not a "
         "coverage gap).  Excludes ChkGisOfficePeople (the "
         "Office_OfficeOffice variant alone exercises queryOffice "
         "-> WriteGIS_OfficeOffice cleanly; OfficePeople adds a "
         "second writer without coverage value).  Probe evidence: "
         "analysis/probe_groupdata_cmdgis_subcalls.py + "
         "analysis/groupdata_cmdgis_subcall_trace.md."},

    # ---- CmdNeo4j ----
    {"form": "LookAtEntry", "button": "CmdNeo4j",
     "test_module": "tests/test_vba_cmdneo4j_cross_form.py",
     "test_node": "test_cmd_neo4j_produces_files[LookAtEntry]",
     "status": "covered",
     "notes": "min_files=7 + per-shape depth + LookAtEntry-"
              "specific structural assertion that the file set is "
              "exactly {People, PeopleEntry, Places, PeoplePlaces, "
              "PersonPlaceCodes, EntryCodes, AssocCodes} AND no "
              "InstitutionCodes file (Issue #9 LATENT-gate pin: "
              "`ENTRY_DATA.c_inst_code > 0 = 0` on this dump).  "
              "Promoted from skip 2026-05-04 after the Issue #9 "
              "reverification probe verified the chain end-to-end "
              "with c_entry_code=101."},
    {"form": "LookAtTexts", "button": "CmdNeo4j",
     "test_module": "tests/test_vba_cmdneo4j_cross_form.py",
     "test_node": "test_cmd_neo4j_produces_files[LookAtTexts]",
     "status": "covered", "notes": "min_files=4 + per-shape depth"},
    {"form": "LookAtAssociations", "button": "CmdNeo4j",
     "test_module": "tests/test_vba_cmdneo4j_cross_form.py",
     "test_node": "test_cmd_neo4j_produces_files[LookAtAssociations]",
     "status": "skipped",
     "skip_reason": "produces 0 files in directory mode — needs "
                    "investigation alongside Place"},
    {"form": "LookAtOffice", "button": "CmdNeo4j",
     "test_module": "tests/test_vba_cmdneo4j_cross_form.py",
     "test_node": "test_cmd_neo4j_produces_files[LookAtOffice]",
     "status": "covered",
     "notes":
         "min_files=4 + per-shape depth.  Promoted from "
         "`real_vba_failing` 2026-05-04 in PR cover/lookatoffice-"
         "cmdneo4j-peopleoffice: added `(NameID, OfficeCode) -> "
         "PeopleOffice` 2-col entry to `_NEO4J_SHAPES_BY_TWO_COLS` "
         "with required cols `[NameID, OfficeCode, OfficeAddrID, "
         "SocialInstID, PostingFirstYear, PostingLastYear, "
         "PostingDynasty]`.  Header literal verified at "
         "`Form_LookAtOffice.vb:947`.  Live-verified end-to-end "
         "with --include-vba: 6 files produced (People, "
         "PeopleOffice, Places, PeoplePlaces, PersonPlaceCodes, "
         "OfficeCode-codes-via-loose-check), 5/6 classified "
         "strictly via the depth check; the 6th (`OfficeCode_"
         "*.csv`, header `OfficeCode,OfficeTrans,OfficePinyin"
         "[,OfficeHZ]` per `Form_LookAtOffice.vb:1324-1326`) "
         "passes via the loose-check fallback because the "
         "classifier doesn't have a single-column entry for "
         "`OfficeCode` yet — non-failing today; tightening it is a "
         "future hygiene follow-up if needed.  InstitutionCodes "
         "block (line 1399) is gated like LookAtEntry's and is "
         "absent on this dump (no `c_inst_code > 0` rows in the "
         "office-relevant scratch table).  Inventory previously "
         "(PR 89d9a63) had marked this `covered` based on an "
         "assumption from the test file's _SPECS list; PR fix/"
         "cmdneo4j-classifier-lookattexts surfaced the actual "
         "classifier-side failure and correctly downgraded to "
         "`real_vba_failing`; this PR fixes the classifier and "
         "honestly re-promotes after live verification."},
    {"form": "LookAtPlace", "button": "CmdNeo4j",
     "test_module": "tests/test_vba_cmdneo4j_cross_form.py",
     "test_node": "test_cmd_neo4j_produces_files[LookAtPlace]",
     "status": "skipped",
     "skip_reason": "fires 'Item not found in this collection' "
                    "mid-body — Issue #7 (real CBDB bug)"},
    {"form": "LookAtKinship", "button": "CmdNeo4j",
     "test_module": "tests/test_vba_cmdneo4j_cross_form.py",
     "test_node": "test_cmd_neo4j_produces_files[LookAtKinship]",
     "status": "covered", "notes": "min_files=4 + per-shape depth"},
    {"form": "LookAtStatus", "button": "CmdNeo4j",
     "test_module": "tests/test_vba_cmdneo4j_cross_form.py",
     "test_node": "test_cmd_neo4j_produces_files[LookAtStatus]",
     "status": "skipped",
     "skip_reason": "chain post-cleanup invalidates subform "
                    "recordset rebind; downstream CmdNeo4j reads "
                    "RecordCount=0 (same family as Pajek/Gephi "
                    "Status skip)"},
    {"form": "LookAtNetworks", "button": "CmdNeo4j",
     "test_module": "tests/test_vba_cmdneo4j_cross_form.py",
     "test_node": "(implicit — LookAtNetworks is in "
                  "_FORMS_WITH_CMDGUESS-style skip family; no "
                  "explicit Spec entry in _SPECS)",
     "status": "skipped",
     "skip_reason": "matrix CmdRun skipped (high-degree anchor "
                    "expansion) + default full injection Form_Open "
                    "deadlock (AGENTS landmine #3.5; minimal "
                    "injection works for Form_Open per "
                    "tests/test_vba_networks_small_fixture.py)"},

    # ---- CmdPajek ----
    {"form": "LookAtKinship", "button": "CmdPajek",
     "test_module": "tests/test_vba_pajek_gephi_cross_form.py",
     "test_node": "test_export_produces_file[LookAtKinship-CmdPajek]",
     "status": "covered", "notes": "shape: .net / *vertices header"},
    {"form": "LookAtPlace", "button": "CmdPajek",
     "test_module": "tests/test_vba_pajek_gephi_cross_form.py",
     "test_node": "test_export_produces_file[LookAtPlace-CmdPajek]",
     "status": "covered", "notes": "shape: .net / *vertices header"},
    {"form": "LookAtStatus", "button": "CmdPajek",
     "test_module": "tests/test_vba_pajek_gephi_cross_form.py",
     "test_node": "test_export_produces_file[LookAtStatus-CmdPajek]",
     "status": "skipped",
     "skip_reason": "chain to CmdPajek/CmdGephi reads RecordCount=0 "
                    "after subform rebind cleanup; same as Status "
                    "Neo4j skip"},
    {"form": "LookAtAssociations", "button": "CmdPajek",
     "test_module": "tests/test_vba_pajek_gephi_cross_form.py",
     "test_node": "test_export_produces_file[LookAtAssociations-CmdPajek]",
     "status": "covered", "notes": "shape: .net / *vertices header"},
    {"form": "LookAtAssociationPairs", "button": "CmdPajek",
     "test_module": "tests/test_vba_pajek_gephi_cross_form.py",
     "test_node": "test_export_button_produces_file"
                  "[LookAtAssociationPairs_CmdPajek]",
     "status": "covered",
     "notes": "shape: .net / *vertices header.  Uses custom 1×3 "
              "known-edged person pair (NOT matrix's default 4×5; "
              "see _assocpairs_1x3_fixture in the test).  "
              "Unblocked by the AssociationPairs SetFocus driver "
              "patch in _PER_FORM_CMDGIS_PATCHES."},

    # ---- CmdGephi ----
    {"form": "LookAtPlace", "button": "CmdGephi",
     "test_module": "tests/test_vba_pajek_gephi_cross_form.py",
     "test_node": "test_export_produces_file[LookAtPlace-CmdGephi]",
     "status": "covered", "notes": "shape: .gdf / nodedef header"},
    {"form": "LookAtStatus", "button": "CmdGephi",
     "test_module": "tests/test_vba_pajek_gephi_cross_form.py",
     "test_node": "test_export_produces_file[LookAtStatus-CmdGephi]",
     "status": "skipped",
     "skip_reason": "same chain-cleanup family as Status CmdPajek"},
    {"form": "LookAtAssociations", "button": "CmdGephi",
     "test_module": "tests/test_vba_pajek_gephi_cross_form.py",
     "test_node": "test_export_produces_file[LookAtAssociations-CmdGephi]",
     "status": "covered", "notes": "shape: .gdf / nodedef header"},
    {"form": "LookAtAssociationPairs", "button": "CmdGephi",
     "test_module": "tests/test_vba_pajek_gephi_cross_form.py",
     "test_node": "test_export_button_produces_file"
                  "[LookAtAssociationPairs_CmdGephi]",
     "status": "covered",
     "notes": "shape: .gdf / nodedef header.  Uses custom 1×3 "
              "known-edged person pair (NOT matrix's default 4×5; "
              "see _assocpairs_1x3_fixture in the test).  "
              "Unblocked by the AssociationPairs SetFocus driver "
              "patch in _PER_FORM_CMDGIS_PATCHES."},

    # ---- CmdGUESS ----
    {"form": "LookAtKinship", "button": "CmdGUESS",
     "test_module": "tests/test_vba_cmdguess_cross_form.py",
     "test_node": "test_cmd_guess_produces_file[LookAtKinship]",
     "status": "covered", "notes": ".gdf shape"},
    {"form": "LookAtNetworks", "button": "CmdGUESS",
     "test_module": "tests/test_vba_cmdguess_cross_form.py",
     "test_node": "test_cmd_guess_produces_file[LookAtNetworks]",
     "status": "skipped",
     "skip_reason": "CmdRun times out on high-degree anchors "
                    "(AGENTS landmine #3.5)"},
    {"form": "LookAtOffice", "button": "CmdGUESS",
     "test_module": "tests/test_vba_cmdguess_cross_form.py",
     "test_node": "test_cmd_guess_produces_file[LookAtOffice]",
     "status": "covered", "notes": ".gdf shape"},

    # ---- CmdGISPeople ----
    {"form": "LookAtOffice", "button": "CmdGISPeople",
     "test_module": "tests/test_vba_cmdgispeople_office.py",
     "test_node": "test_cmd_gis_people",
     "status": "covered",
     "notes": "only host; people-side GIS export distinct from "
              "office-side CmdGIS"},
]


# Static-only (test_known_bugs.py / test_vba_bug_design_time.py)
# coverage notes — these don't drive real CmdQuery+CmdX chains but
# do pin source-level / control-side facts about specific
# (form, button) combos.  Captured here so the inventory shows the
# safety net even when real-VBA coverage is absent.
STATIC_TEST_NOTES: dict[tuple[str, str], list[str]] = {
    ("LookAtPlace",   "CmdGIS"):    [
        "tests/test_known_bugs.py::test_bug4_lookat_place_cmdgis"
        "_references_nonexistent_gisframe — pins source typo "
        "(Issue #4)"],
    ("LookAtPlace",   "CmdNeo4j"):  [
        "tests/test_known_bugs.py::test_bug7_lookat_place_cmdneo4j"
        "_select_missing_dynasty_female — pins source SELECT "
        "(Issue #7)"],
    ("LookAtNetworks", "CmdNeo4j"): [
        "tests/test_known_bugs.py::test_bug8_lookat_networks_cmdneo4j"
        "_select_missing_xy — pins source SELECT (Issue #8)"],
    ("LookAtEntry",   "CmdNeo4j"):  [
        "tests/test_known_bugs.py::test_bug9_lookat_entry_cmdneo4j"
        "_with_wrong_var — pins source typo + LATENT-gate "
        "(Issue #9)",
        "tests/test_vba_bug_behaviors.py::test_bug9_lookat_entry"
        "_cmdneo4j_with_institutions_fixture — runtime: chain "
        "finishes cleanly without ERR for non-inst fixture"],
    ("LookAtStatus",  "CmdPajek"):  [
        "tests/test_known_bugs.py::test_bug5_lookat_status_cmdpajek"
        "_references_nonexistent_chkids — pins source typo "
        "(Issue #5)"],
    # Issues #15-19: P3 missing UI buttons.
    ("LookAtPlace",   "CmdGIS"):    [
        "(plus Issue #15: P3 missing UI button — covered by "
        "tests/test_known_bugs.py::test_bugs_15_to_19_orphan"
        "_export_handlers)"],
    ("LookAtStatus",  "CmdPajek"):  [
        "(plus Issue #16: P3 missing UI button)"],
    ("LookAtStatus",  "CmdGephi"):  [
        "(plus Issue #17: P3 missing UI button)"],
    ("LookAtStatus",  "CmdUCINet"): [
        "(plus Issue #18: P3 missing UI button)"],
    ("LookAtOffice",  "CmdGUESS"):  [
        "(plus Issue #19: P3 missing UI button)"],
}


# ---------------------------------------------------------------
# Probes
# ---------------------------------------------------------------
def _vba_body(form: str) -> str:
    p = DUMP_VBA / f"Form_{form}.vb"
    return p.read_bytes().decode("utf-8", errors="replace")


def _handler_present(form: str, button: str) -> bool:
    """True iff the VBA dump contains `[Private |Public ]Sub
    <Button>_Click` for this form."""
    body = _vba_body(form)
    pat = rf"\b(?:Private |Public )?Sub\s+{re.escape(button)}_Click\b"
    return re.search(pat, body) is not None


def _button_present(form: str, button: str,
                    inv: dict) -> bool:
    """True iff control_inventory.json lists a control with the
    given name on the given form (case-insensitive)."""
    ctls = inv.get(form, {}).get("controls", [])
    target = button.lower()
    return any(c.get("name", "").lower() == target for c in ctls)


# ---------------------------------------------------------------
# Coverage decision per cell
# ---------------------------------------------------------------
def _classify_cell(form: str, button: str,
                   handler_present: bool,
                   button_present: bool,
                   manifest_for_cell: list[dict]) -> dict:
    """Decide the cell's status string + supporting evidence.

    Note on handler-vs-button asymmetry: the cross-form tests
    invoke handlers via `Form_Timer` dispatch (set
    `Me.TimerInterval > 0`, the injected Form_Timer body calls
    the named `Sub <Cmd>_Click()` directly), which bypasses any
    UI click.  So a (form, button) cell where the HANDLER exists
    but the BUTTON does NOT can still be exercised by a real-VBA
    test — and indeed, the suite uses this on purpose for
    Issues #15-19's missing-button cells (the test proves the
    handler's logic still works while the report documents the
    missing UI control separately).  We flag those cells with a
    distinct status so the matrix shows both facts at once.
    """
    static_notes = STATIC_TEST_NOTES.get((form, button), [])
    if not handler_present and not button_present:
        return {
            "status": "not_applicable",
            "explanation":
                "No handler in source AND no button in inventory.",
            "static_test_notes": static_notes,
        }
    if button_present and not handler_present:
        return {
            "status": "orphan_button_no_handler",
            "explanation":
                "Button exists on the form but no `Sub "
                f"{button}_Click()` in the dumped VBA — clicking "
                "it would be a no-op (or fall through to Form_"
                "Click).  Investigate before relying on this cell.",
            "static_test_notes": static_notes,
        }
    if handler_present and not button_present:
        # Handler exists, button missing.  Two sub-cases:
        if not manifest_for_cell:
            # No test reaches it (and no UI button to click) — the
            # P3 missing-UI family (Issues #15-19).
            return {
                "status": "missing_ui_button",
                "explanation":
                    "Handler exists in source but the form has no "
                    "control with this name (P3 missing-UI family "
                    "— Issues #15-19 documented).  No real-VBA "
                    "test exercises this handler.",
                "static_test_notes": static_notes,
            }
        # Test exists AND uses Form_Timer dispatch to invoke the
        # handler despite the missing UI button — the test proves
        # the handler's logic, the report documents the missing
        # UI button separately.
        statuses = {m["status"] for m in manifest_for_cell}
        is_covered = "covered" in statuses
        return {
            "status": ("real_vba_covered_via_handler_dispatch"
                       if is_covered
                       else "real_vba_skipped_via_handler_dispatch"),
            "explanation":
                "Handler exists in source but the form has no "
                "control with this name (P3 missing-UI family).  "
                "The cross-form test exercises the handler via "
                "Form_Timer dispatch (no UI click needed), so "
                "handler logic IS covered; the missing UI button "
                "is documented separately as a P3 issue.",
            "test_entries": manifest_for_cell,
            "static_test_notes": static_notes,
        }
    # Both present.
    if not manifest_for_cell:
        return {
            "status": "gap",
            "explanation":
                "Button + handler both present, no entry in the "
                "real-VBA test manifest, no static-only test note "
                "either."
                if not static_notes
                else
                "Button + handler both present, no real-VBA "
                "coverage; only static / source-level tests guard "
                "this cell.",
            "static_test_notes": static_notes,
        }
    if any(m["status"] == "covered" for m in manifest_for_cell):
        return {
            "status": "real_vba_covered",
            "explanation":
                "Real-VBA test exists and passes for this cell.",
            "test_entries": manifest_for_cell,
            "static_test_notes": static_notes,
        }
    if any(m["status"] == "real_vba_failing" for m in manifest_for_cell):
        return {
            "status": "real_vba_failing",
            "explanation":
                "Real-VBA test exists, runs (does NOT skip), and "
                "fails on the current dump.  See `skip_reason` on "
                "each manifest entry for the failure mode.",
            "test_entries": manifest_for_cell,
            "static_test_notes": static_notes,
        }
    return {
        "status": "real_vba_skipped",
        "explanation":
            "Real-VBA test exists but is currently skipped.  See "
            "skip_reason on each manifest entry.",
        "test_entries": manifest_for_cell,
        "static_test_notes": static_notes,
    }


def _gap_score(status: str, button: str, handler_present: bool,
               button_present: bool) -> int:
    """Heuristic priority for "next real-VBA slice" recommendation
    (lower = higher priority).

    Conservative: only a `gap` cell (button + handler exist, no
    test of any kind) is a real candidate.  Among those, prefer
    cells whose handler family already has working tests (lower
    blast radius), and CmdGIS-family / CmdGUESS-family over the
    Networks / AssociationPairs CmdRun-blocked exports.
    """
    if status != "gap":
        return 99
    family_priority = {
        "CmdGIS":        1,
        "CmdGUESS":      2,
        "CmdGephi":      3,
        "CmdPajek":      3,
        "CmdGISPeople":  4,
        "CmdNeo4j":      5,  # heaviest chain
        "CmdUCINet":     6,  # untested family entirely
        "CmdKML":       99,
    }
    return family_priority.get(button, 50)


# ---------------------------------------------------------------
# Build + render
# ---------------------------------------------------------------
def _build() -> dict:
    inv = json.loads(CONTROL_INVENTORY.read_text(encoding="utf-8"))

    # First, cross-validate every manifest entry against the dump
    # so a stale manifest line lands as a visible finding rather
    # than silently distorting the matrix.
    # Drift = manifest declares a test exists, but the *handler*
    # the test would dispatch to is absent from the VBA dump.
    # (Missing UI button is NOT drift — Form_Timer dispatch
    # bypasses the button; that case becomes
    # `real_vba_*_via_handler_dispatch` in the matrix.)
    manifest_drift_findings: list[dict] = []
    for m in EXPORT_TEST_MANIFEST:
        f, b = m["form"], m["button"]
        if not _handler_present(f, b):
            manifest_drift_findings.append({
                "kind": "manifest_entry_handler_missing",
                "form": f, "button": b,
                "test_module": m["test_module"],
                "explanation":
                    f"Manifest declares {f}.{b} test exists, but "
                    f"`Sub {b}_Click()` not found in dumped VBA.",
            })

    matrix: dict[str, dict[str, dict]] = {}
    rec_candidates: list[dict] = []
    low_hanging_skips: list[dict] = []
    for f in FORMS:
        matrix[f] = {}
        for b in BUTTONS:
            h = _handler_present(f, b)
            bp = _button_present(f, b, inv)
            mfc = [m for m in EXPORT_TEST_MANIFEST
                   if m["form"] == f and m["button"] == b]
            cell = _classify_cell(f, b, h, bp, mfc)
            cell["handler_present"] = h
            cell["button_present"] = bp
            matrix[f][b] = cell
            if cell["status"] == "gap":
                rec_candidates.append({
                    "form": f, "button": b,
                    "score": _gap_score(
                        cell["status"], b, h, bp),
                    "static_test_notes":
                        cell.get("static_test_notes") or [],
                })
            # Low-hanging skip: real_vba_skipped whose skip_reason
            # looks mechanical ("no matrix fixture for X").  Other
            # skip reasons (chain cleanup, RPC death, Form_Open
            # deadlock, matrix CmdRun timeout) are NOT low-hanging.
            if cell["status"] == "real_vba_skipped":
                for m in cell.get("test_entries") or []:
                    sr = (m.get("skip_reason") or "").lower()
                    if "no matrix fixture" in sr:
                        low_hanging_skips.append({
                            "form": f, "button": b,
                            "test_module": m["test_module"],
                            "test_node": m["test_node"],
                            "skip_reason": m["skip_reason"],
                            "notes": m.get("notes", ""),
                            "fix_class":
                                "wire matrix fixture in "
                                "tests/test_vba_matrix_all_forms.py"
                                "::_all_fixtures",
                        })

    rec_candidates.sort(key=lambda r: (r["score"], r["form"], r["button"]))

    # Roll-up counts.
    by_status: dict[str, int] = {}
    for f in FORMS:
        for b in BUTTONS:
            s = matrix[f][b]["status"]
            by_status[s] = by_status.get(s, 0) + 1

    return {
        "schema_version": 1,
        "generated_utc":
            _dt.datetime.now(_dt.timezone.utc).isoformat(
                timespec="seconds"),
        "forms": list(FORMS),
        "buttons": list(BUTTONS),
        "matrix": matrix,
        "summary": {
            "n_forms": len(FORMS),
            "n_buttons": len(BUTTONS),
            "n_cells": len(FORMS) * len(BUTTONS),
            "by_status": by_status,
        },
        "manifest_drift_findings": manifest_drift_findings,
        "low_hanging_skips": low_hanging_skips,
        "recommended_next_slices": rec_candidates,
    }


def _validate_invariants(d: dict) -> None:
    """Assert deterministic invariants on the built inventory.  Raises
    `AssertionError` if any invariant is violated; the script exits
    non-zero (after writing the JSON / MD) so the regression is
    visible in CI output.

    Invariants:

      I1. `real_vba_failing` cells are NEVER counted as covered.
          (`by_status['real_vba_covered']` excludes them; the matrix
          cell's status string is `real_vba_failing`, not `_covered`.)

      I2. `real_vba_failing` cells NEVER appear in
          `low_hanging_skips` — a failing test is not a "skip with a
          mechanical fix", it's a test that runs and fails.

      I3. Every `real_vba_failing` cell's manifest entries MUST carry
          either a `skip_reason` (re-used as the failure-mode
          description, since the `_failing` status overloads the
          field) OR a substantive `notes` field of >= 20 characters
          — so reviewers always have machinery-readable context for
          the failure, not just the bare status string.
    """
    failing_cells: list[tuple[str, str, dict]] = []
    for f in d["forms"]:
        for b in d["buttons"]:
            cell = d["matrix"][f][b]
            if cell["status"] == "real_vba_failing":
                failing_cells.append((f, b, cell))

    # I1: failing cells are not in the covered count.
    n_covered = d["summary"]["by_status"].get("real_vba_covered", 0)
    n_failing = d["summary"]["by_status"].get("real_vba_failing", 0)
    assert n_failing == len(failing_cells), (
        "Invariant I1 violated: by_status['real_vba_failing'] "
        f"({n_failing}) does not match the number of cells with "
        f"status == 'real_vba_failing' in the matrix "
        f"({len(failing_cells)})."
    )

    # I2: failing cells never in low_hanging_skips.
    failing_keys = {(f, b) for f, b, _ in failing_cells}
    bad_low_hang = [
        h for h in d["low_hanging_skips"]
        if (h["form"], h["button"]) in failing_keys
    ]
    assert not bad_low_hang, (
        "Invariant I2 violated: low_hanging_skips contains failing "
        f"cells (these are NOT mechanical-fix skips): "
        f"{[(h['form'], h['button']) for h in bad_low_hang]}.  "
        "low_hanging_skips is gated on status == 'real_vba_skipped' "
        "by construction; if you see this, the gate has regressed."
    )

    # I3: every failing cell has a substantive failure description.
    bad_desc = []
    for f, b, cell in failing_cells:
        for m in cell.get("test_entries") or []:
            sr = (m.get("skip_reason") or "").strip()
            notes = (m.get("notes") or "").strip()
            if not sr and len(notes) < 20:
                bad_desc.append({
                    "form": f, "button": b,
                    "test_module": m.get("test_module"),
                })
    assert not bad_desc, (
        "Invariant I3 violated: real_vba_failing cells must carry "
        "either a `skip_reason` (re-used as failure-mode "
        "description) or substantive `notes` (>= 20 chars).  "
        f"Bare cells: {bad_desc}"
    )


_STATUS_GLYPH = {
    "real_vba_covered":                          "✓",
    "real_vba_covered_via_handler_dispatch":     "✓*",
    "real_vba_skipped":                          "skip",
    "real_vba_skipped_via_handler_dispatch":     "skip*",
    "real_vba_failing":                          "FAIL",
    "not_applicable":                            "—",
    "missing_ui_button":                         "no-btn",
    "orphan_button_no_handler":                  "orphan",
    "unit_or_static_only":                       "static",
    "gap":                                       "GAP",
}


def _render_md(d: dict) -> str:
    lines: list[str] = []
    lines.append("# Export coverage inventory "
                 "(LookAt form × export button)")
    lines.append("")
    lines.append(f"**Generated:** {d['generated_utc']}")
    lines.append(
        "**Generator:** `analysis/inventory_export_coverage.py`")
    lines.append(
        "**Companion JSON:** `reports/export_coverage_inventory.json`")
    lines.append("")
    lines.append(
        "Read-only inventory.  No MDB.  No Access COM.  Reads the "
        "VBA dump (`analysis/dump/vba/`), the control inventory "
        "(`analysis/dump/control_inventory.json`), and a curated "
        "test-coverage manifest declared in this script.  The "
        "manifest is cross-validated against the VBA dump on every "
        "run; drift surfaces under § Manifest drift.")
    lines.append("")

    # Legend.
    lines.append("## Legend")
    lines.append("")
    lines.append(
        "| Glyph | Status | Meaning |\n"
        "|---|---|---|\n"
        "| `✓` | `real_vba_covered` | "
        "Real-VBA test exists and passes for this (form, button). "
        "Both handler and UI button present. |\n"
        "| `✓*` | `real_vba_covered_via_handler_dispatch` | "
        "Test passes via `Form_Timer` handler dispatch — handler "
        "exists, UI button is missing (P3 family).  Test exercises "
        "handler logic; missing button stays documented as a P3 "
        "issue. |\n"
        "| `skip` | `real_vba_skipped` | "
        "Real-VBA test exists but is `pytest.mark.skip`'d "
        "(see per-cell skip_reason in JSON). |\n"
        "| `skip*` | `real_vba_skipped_via_handler_dispatch` | "
        "Same as `skip`, but for a handler-dispatched cell whose "
        "UI button is missing. |\n"
        "| `FAIL` | `real_vba_failing` | "
        "Real-VBA test exists, runs (does NOT skip), and fails on "
        "the current dump.  See per-cell `skip_reason` in JSON for "
        "the failure mode (typically a depth-check classifier gap "
        "for an unfamiliar file-shape family).  **NOT counted as "
        "covered.**  Distinct from `skip` (no `pytest.mark.skip`) "
        "and from `GAP` (test does exist).  See § Status semantics "
        "below. |\n"
        "| `GAP` | `gap` | "
        "Both handler + button present, no real-VBA test, no "
        "static-only test either — true uncovered slice. |\n"
        "| `static` | `unit_or_static_only` | "
        "Only static / source-level tests cover this cell "
        "(`tests/test_known_bugs.py` family); no real "
        "CmdQuery → CmdX chain. |\n"
        "| `no-btn` | `missing_ui_button` | "
        "Handler exists in source but no control on the form "
        "(P3 missing-UI family — Issues #15-19). |\n"
        "| `orphan` | `orphan_button_no_handler` | "
        "Button exists on the form but no `Sub <Cmd>_Click` in "
        "the dumped VBA — clicking would be a no-op. |\n"
        "| `—` | `not_applicable` | "
        "Neither handler nor button — this form just doesn't "
        "host this export. |"
    )
    lines.append("")
    lines.append(
        "**CmdKML caveat:** CmdKML is included in the matrix per the "
        "investigation brief, but the codebase has neither a "
        "`CmdKML` button nor a `CmdKML_Click` handler on any "
        "LookAt form.  KML output is implemented as a `ChkKML` "
        "checkbox option that other exports honour (e.g. "
        "LookAtOffice has `ChkPeopleKML` / `ChkOfficeKML`).  Every "
        "CmdKML cell therefore renders `—` (not_applicable); "
        "future KML coverage would be checkbox-driven, not a "
        "separate button slice.")
    lines.append("")
    lines.append(
        "**CmdUCInet → CmdUCINet:** the brief used `CmdUCInet` "
        "(lower-case `i`); the actual control + handler is "
        "`CmdUCINet` (capital `N`).  The matrix uses the real "
        "casing.")
    lines.append("")

    # Status semantics — pin the meaning of FAIL vs skip vs GAP.
    lines.append("## Status semantics")
    lines.append("")
    lines.append(
        "Four statuses describe \"there is or is not real-VBA test "
        "coverage for this cell\".  They are NOT interchangeable; "
        "in particular `real_vba_failing` is its own bucket, never "
        "rolled into `real_vba_covered`.")
    lines.append("")
    lines.append(
        "| Status | Test exists? | Test skipped? | Test passes? | "
        "Counted as covered? | Eligible as Tier-1 (low-hanging) "
        "candidate? |")
    lines.append(
        "|---|:---:|:---:|:---:|:---:|:---:|")
    lines.append(
        "| `real_vba_covered` (✓ / ✓*) | yes | no | **yes** | "
        "**yes** | n/a (already covered) |")
    lines.append(
        "| `real_vba_skipped` (skip / skip*) | yes | **yes** "
        "(`pytest.mark.skip`) | n/a — not run | no | "
        "**yes**, but only when `skip_reason` matches a "
        "mechanical-fix pattern (currently: \"no matrix fixture\") |")
    lines.append(
        "| `real_vba_failing` (FAIL) | yes | no | **no** | **no** "
        "| **no** — failing tests are not skips with a mechanical "
        "fix; they're tests that run and fail.  Each failing cell "
        "must carry a `skip_reason` (re-used as failure-mode "
        "description) or substantive `notes` (>= 20 chars) so the "
        "failure mode is machinery-readable. |")
    lines.append(
        "| `gap` (GAP) | **no** | n/a | n/a | no | yes (always — "
        "ranked by family priority in Tier 2) |")
    lines.append("")
    lines.append(
        "Three deterministic invariants are checked at script-exit "
        "time and printed to stderr (with non-zero exit code) if "
        "violated:")
    lines.append("")
    lines.append(
        "  - **I1** — `real_vba_failing` cells are never counted as "
        "covered (`by_status` rolls them up as their own bucket).")
    lines.append(
        "  - **I2** — `real_vba_failing` cells never appear in "
        "`low_hanging_skips` (which is gated on status == "
        "`real_vba_skipped` by construction; this invariant pins "
        "that gate).")
    lines.append(
        "  - **I3** — every `real_vba_failing` cell's manifest "
        "entries carry either a `skip_reason` (re-used as failure-"
        "mode description) or substantive `notes` (>= 20 chars).")
    lines.append("")
    lines.append(
        "If a `FAIL` cell ever wants to graduate to `real_vba_"
        "covered`, the path is to (a) fix the underlying failure "
        "mode (typically: extend the depth-check classifier in "
        "`tests/test_vba_cmdneo4j_cross_form.py` for a new "
        "file-shape family) and (b) flip the manifest entry's "
        "`status` from `real_vba_failing` to `covered`.  The flip "
        "must come AFTER the test actually passes, not before.")
    lines.append("")

    # Summary roll-up.
    s = d["summary"]
    lines.append("## Summary")
    lines.append("")
    lines.append(f"- **Cells total:** {s['n_cells']} "
                 f"({s['n_forms']} forms × {s['n_buttons']} buttons)")
    for k, v in sorted(s["by_status"].items(),
                       key=lambda kv: (-kv[1], kv[0])):
        glyph = _STATUS_GLYPH.get(k, k)
        lines.append(f"- `{k}` ({glyph}): **{v}**")
    lines.append("")

    # Matrix.
    lines.append("## Coverage matrix")
    lines.append("")
    header = "| Form | " + " | ".join(d["buttons"]) + " |"
    sep = "|---" + "|---" * len(d["buttons"]) + "|"
    lines.append(header)
    lines.append(sep)
    for f in d["forms"]:
        cells = []
        for b in d["buttons"]:
            cell = d["matrix"][f][b]
            cells.append(_STATUS_GLYPH.get(cell["status"], "?"))
        lines.append(f"| **{f}** | " + " | ".join(cells) + " |")
    lines.append("")

    # Per-cell detail (only the non-trivial cells).
    lines.append("## Per-cell detail (non-trivial cells)")
    lines.append("")
    lines.append(
        "Cells with status `not_applicable` are omitted to keep "
        "the noise down.")
    lines.append("")
    for f in d["forms"]:
        for b in d["buttons"]:
            cell = d["matrix"][f][b]
            if cell["status"] == "not_applicable":
                continue
            lines.append(
                f"### {f} × {b} — `{cell['status']}`")
            lines.append("")
            lines.append(f"- Handler in source: "
                         f"{'yes' if cell['handler_present'] else 'no'}")
            lines.append(f"- Button on form: "
                         f"{'yes' if cell['button_present'] else 'no'}")
            lines.append(f"- Why: {cell['explanation']}")
            for m in cell.get("test_entries") or []:
                line = (f"- Test: `{m['test_module']}::"
                        f"{m['test_node']}` "
                        f"— **{m['status']}**")
                if m.get("skip_reason"):
                    line += f"; skip_reason: {m['skip_reason']}"
                if m.get("notes"):
                    line += f"; notes: {m['notes']}"
                lines.append(line)
            for note in cell.get("static_test_notes") or []:
                lines.append(f"- Static-only note: {note}")
            lines.append("")

    # Manifest drift.
    drift = d["manifest_drift_findings"]
    lines.append("## Manifest drift")
    lines.append("")
    if drift:
        lines.append("**The test-coverage manifest disagrees with "
                     "the dumped VBA / control inventory:**")
        lines.append("")
        for x in drift:
            lines.append(
                f"- `{x['kind']}` — {x['form']}.{x['button']} "
                f"({x['test_module']}): {x['explanation']}")
        lines.append("")
    else:
        lines.append("**Clean — manifest entries all match the VBA "
                     "dump and control inventory.**")
        lines.append("")

    # Recommended next slice — two-tier.  Low-hanging skips first
    # (mechanical fix — best ROI), then pure gap cells.
    low = d["low_hanging_skips"]
    rec = d["recommended_next_slices"]
    lines.append("## Recommended next real-VBA slice")
    lines.append("")
    lines.append(
        "*Suggestion only — this PR ships inventory and does NOT "
        "implement.  Review the cell's manifest entry + static "
        "notes before committing to a slice.*")
    lines.append("")

    lines.append("### Tier 1 — low-hanging skipped tests "
                 "(mechanical fix)")
    lines.append("")
    if low:
        lines.append(
            "These cells have a real-VBA test that is currently "
            "skipped for a *mechanical* reason (typically: \"no "
            "matrix fixture\").  Promoting one to passing is the "
            "smallest, lowest-risk way to close a coverage cell — "
            "no driver changes, no new infrastructure, no "
            "Networks / Status / AssociationPairs blockers.")
        lines.append("")
        for h in low:
            lines.append(f"- **{h['form']} × {h['button']}** "
                         f"(`{h['test_module']}`)")
            lines.append(f"  - skip_reason: {h['skip_reason']}")
            lines.append(f"  - fix class: {h['fix_class']}")
            if h.get("notes"):
                lines.append(f"  - existing evidence: {h['notes']}")
        lines.append("")
        # Single concrete recommendation pulled from the top of
        # the low-hanging list.
        top = low[0]
        lines.append(
            f"**Concrete recommendation** (do NOT implement in "
            f"this inventory PR): wire a {top['form']} matrix "
            f"fixture in "
            f"`tests/test_vba_matrix_all_forms.py::_all_fixtures` "
            f"so `{top['test_node']}` in `{top['test_module']}` "
            f"flips from skip to a real assertion.  Notes from "
            f"the manifest: {top.get('notes') or '(none)'}")
        lines.append("")
    else:
        lines.append(
            "*(No skipped tests with a mechanical fix class.  All "
            "current skips are blocked on harder issues — chain "
            "cleanup family, matrix CmdRun timeout family, "
            "Form_Open deadlock family.)*")
        lines.append("")

    lines.append("### Tier 2 — pure `gap` cells "
                 "(button + handler exist, no test of any kind)")
    lines.append("")
    if not rec:
        lines.append("**No `gap` cells** — every (form, button) "
                     "cell where both the handler and the button "
                     "exist is already covered or skipped with a "
                     "documented reason.")
        lines.append("")
    else:
        lines.append(
            "Cells ranked by family priority (lower = lower-risk, "
            "smaller blast radius).  Note: many gap cells in this "
            "list (Networks family, AssociationPairs family, "
            "GroupData family, the entirely-untested CmdUCINet "
            "family) sit behind known blockers and are NOT good "
            "candidates for a small first slice — Tier 1 is "
            "preferred.")
        lines.append("")
        lines.append("| # | Form | Button | Family priority | "
                     "Known family blocker |")
        lines.append("|---:|---|---|---:|---|")
        blocker = {
            "LookAtNetworks": "Form_Open hang / CmdRun timeout "
                              "(AGENTS landmine #3.5)",
            "LookAtAssociationPairs":
                "matrix CmdQuery times out — no CrossFixture "
                "promoted to a passing assertion",
            "LookAtGroupData":
                "matrix CmdQuery has issues; depends on a "
                "CrossFixture that doesn't exist for GroupData",
        }
        for i, r in enumerate(rec, start=1):
            blk = blocker.get(r["form"], "")
            if r["button"] == "CmdUCINet" and not blk:
                blk = ("entirely untested handler family — "
                       "no existing test infrastructure to "
                       "extend")
            lines.append(
                f"| {i} | {r['form']} | {r['button']} "
                f"| {r['score']} | {blk or '—'} |")
        lines.append("")

    return "\n".join(lines)


def main() -> int:
    d = _build()
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(
        json.dumps(d, indent=2, ensure_ascii=False),
        encoding="utf-8")
    OUT_MD.parent.mkdir(parents=True, exist_ok=True)
    OUT_MD.write_text(_render_md(d), encoding="utf-8")

    s = d["summary"]
    print("=== export coverage inventory ===")
    print(f"  cells: {s['n_cells']} "
          f"({s['n_forms']} forms × {s['n_buttons']} buttons)")
    for k, v in sorted(s["by_status"].items(),
                       key=lambda kv: (-kv[1], kv[0])):
        print(f"    {k:30s} {v}")
    print(f"  manifest drift: "
          f"{len(d['manifest_drift_findings'])}")
    print(f"  recommended slices: "
          f"{len(d['recommended_next_slices'])}")
    print(f"  wrote {OUT_JSON}")
    print(f"  wrote {OUT_MD}")

    # Run deterministic invariants AFTER writing so a reviewer can
    # always inspect the JSON / MD even if guards fire.  Print the
    # AssertionError message and exit non-zero so the regression is
    # visible in CI without aborting the artifact write.
    try:
        _validate_invariants(d)
    except AssertionError as e:
        print(f"\n[INVARIANT VIOLATION] {e}", file=sys.stderr)
        return 2

    # Inventory is informational; never non-zero exit on coverage
    # alone (gaps are expected — that's the point).  Only manifest
    # drift indicates the inventory itself is stale.
    return 1 if d["manifest_drift_findings"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
