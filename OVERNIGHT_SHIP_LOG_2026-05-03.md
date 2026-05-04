# Overnight ship log — 2026-05-03

Per the autopilot brief: 8 PRs were queued (AA–AH) plus
self-directed exploration after.  All merged to `main` and
pushed; branches deleted.  Fast suite stayed at **111 passed,
9 skipped** throughout; audits at **0 above baseline**.

## Merged PRs (in order)

| Branch | SHA | Description |
| --- | --- | --- |
| `feat/pr-aa-lookatnetworks-formopen-investigation` | `19cce20` | LookAtNetworks Form_Open is fine; CmdRun is the real hang |
| `feat/pr-ac-hard-form-skip-inventory` | `1091497` | hard-form skip / xfail inventory + blocker classification |
| `feat/pr-ad-export-coverage-matrix` | `0b5c940` | export coverage matrix across all LookAt* buttons |
| `feat/pr-ae-export-delimiter-risk-audit` | `510498d` | static export delimiter risk audit |
| `feat/pr-af-gis-reach-extension` | `8d32ad8` | GIS reach extension — Office side + Kinship recursive |
| `feat/pr-ag-index-addr-tiebreak-note` | `a3590c6` | index_addr same-candidate tie-break repro note |
| `feat/pr-ah-pytest-marker-inventory` | `c9d47cf` | pytest collection / marker health audit |
| `feat/pr-ai-consistent-within-rule-20-cluster` | `8d29d35` | **K2 diff=-20 cluster: hypothesis REVERSED — actual cause is upstream BIOG_MAIN drift** |
| `feat/pr-aj-rule11-plus1-cluster` | `411a81f` | K2 diff=+1 cluster: KIN_DATA drift dominates |

PR AB was naturally subsumed by AA — its scope assumed AA
would identify a Form_Open mitigation, but AA found Form_Open
doesn't hang at all.  Logged here, not a separate PR.

## Tests run

Per the overnight rules, after each PR:
- `python -m pytest tests/ -W ignore` → 111 passed, 9 skipped
- `python analysis/run_all_audits.py --ci` → 0 above baseline

Where reports were touched:
- `python reports/generate_report.py` (only PR AE invoked
  this; output unchanged in committed `.md` files because
  the new audit doesn't write into the report).

No `--include-vba` runs, no `pytest tests/test_vba_*.py`
runs, no Access COM activity except inside PR AA's probe
script (which I ran manually with watchdog supervision and
killed any wedged MSACCESS.EXE between variations).

## Headline findings

These are the items that *substantively change* what we
believed at the start of the overnight session.

### F1. LookAtNetworks Form_Open is NOT broken (PR AA)

The README / AGENTS / `test_vba_bug_behaviors.py:270` claim
that `LookAtNetworks.Form_Open hangs in this driver` is
**misleading**.  The probe ran `OpenForm("LookAtNetworks", ...)`
in 4 configurations under per-process Access COM and all 4
returned in <2 s once the cold-cache OpenCurrentDatabase
issue was resolved.

The actual blocker is `CmdRun` (already noted accurately at
`test_vba_matrix_all_forms.py:313`: "Networks (CmdRun times
out — Zhu Xi has 2471 assocs)").

**Open question for morning**: should the
`test_vba_bug_behaviors.py:270` skip reason wording be
updated, and should we attempt a smaller-fixture LookAtNetworks
unskip experiment (was-PR-AB scope)?  The `analysis/hard_form_
skip_inventory.md` has 4 occurrences flagged
`lookatnetworks_form_open_hang_legacy_label` waiting on this
decision.

### F2. K2 `consistent_within_rule × 14` bucket is upstream-data drift, NOT algorithm divergence (PR AI + AJ)

PR Y had this bucket labelled `candidate_same_rule_tie_break_
or_aggregation_diff` at confidence `medium`.  Probing all 14
rows shows:

  - **8 rows**: BIOG_MAIN.c_birthyear differs between User MDB
    and SQLite snapshot for the same evidence person, by
    exactly the same amount as the index_year diff.
  - **6 rows**: KIN_DATA evidence pid set differs between sides
    (different father / sibling rows recorded on each side).

100% of the 14 rows are upstream-data drift between the User
MDB and the SQLite snapshot.  Same general class as PR Z's
tcode='05' candidate_php_entry_code_mapping_gap.  The Access
vs PHP rule body is identical and runs correctly on both sides.

**Open question for morning**: rename the bucket label in
`reports/index_drift_cause_summary.json` from
`candidate_same_rule_tie_break_or_aggregation_diff` to
something like `source_data_drift_biog_main_or_kin_data_
between_sides`, and promote confidence from `medium` to
`supported_by_focused_probe`.  Held back overnight per
"don't change severity without overwhelming evidence +
maintainer sign-off" — evidence is overwhelming but the
rename is maintainer-visible.

### F3. PR AE found 4 NEW candidate delimiter-risk findings (parallel to Issue #20)

Static scan of 8 export-bound source tables (32 columns)
found:

  - `OFFICE_CODES.c_office_chn` U+FEFF × **5** rows.  Same
    bug class as Issue #20 (BOM + JET-mangle).  PR AF
    confirmed:
      - All 5 dirty offices have **0 persons posted** to them
        in `POSTED_TO_OFFICE_DATA` — same orphan pattern as
        314/315 dirty ADDR_CODES rows.
      - Reachable via `LookAtOffice.CmdNeo4j_Click` →
        OfficeCodes.csv (line 1360 + 2676), TAB-separated
        without escaping.
      - Not reachable via `LookAtOffice.CmdGIS_Click` (which
        joins ADDR_CODES, not OFFICE_CODES, for AddrName).
  - `BIOG_MAIN.c_notes` LF × 193 / CR × 193 / TAB × 1 — only
    matters if any export emits c_notes; not currently the
    case for the GIS / Neo4j paths I scanned.

**Open question for morning**: promote OFFICE_CODES finding
to a new sub-issue (Issue #20.b) or fold into Issue #20's
"Known reach (PR W)" section.  c_notes findings can wait
unless someone wants a focused reach probe.

### F4. PR Z verdict for tcode='05' × 7 stays — focused probe confirmed PHP-side data gap

(Pre-overnight finding, mentioned for morning context.)
6/7 confirmed `candidate_php_entry_code_mapping_gap` (SQLite
ENTRY_CODE_TYPE_REL missing the '040101' membership);
1/7 confirmed `candidate_php_entry_data_year_missing`
(SQLite ENTRY_DATA.c_year = 0 for the same entry where
User MDB has the year).  All 7 are PHP-side upstream data
gaps; Access fires Rule 05 correctly.

## Candidate findings NOT promoted to ISSUES

None.  Per the overnight rule "Do not change issue severity
unless evidence is overwhelming; record candidates only,"
all four candidate findings above are recorded in
`analysis/` + `reports/` only.  None added to
`reports/generate_report.py`'s ISSUES list.

## Stop conditions encountered

None of the listed stop conditions hit.  Fast suite stayed
green; audits stayed at baseline; no destructive operations
were needed beyond standard branch deletion.

## Files added overnight

By directory:

```
analysis/
  hard_form_skip_inventory.py
  hard_form_skip_inventory.md
  export_coverage_matrix.py
  export_coverage_matrix.md
  audit_export_delimiter_risk.py
  export_delimiter_risk_audit.md
  analyze_gis_office_addr_reach.py
  gis_office_addr_reach.md
  index_addr_same_candidate_tiebreak.md
  render_index_addr_tiebreak_note.py
  pytest_marker_inventory.py
  pytest_marker_inventory.md
  probe_index_year_diff_minus_20_cluster.py
  index_year_diff_minus_20_cluster_probe.md
  probe_index_year_diff_plus_1_cluster.py
  index_year_diff_plus_1_cluster_probe.md
  probe_lookatnetworks_form_open.py        (PR AA)
  lookatnetworks_form_open_hang.md         (PR AA)

reports/
  hard_form_skip_inventory.json
  export_coverage_matrix.json
  export_delimiter_risk_audit.json
  gis_office_addr_reach.json
  index_addr_same_candidate_tiebreak.json
  pytest_marker_inventory.json
  index_year_diff_minus_20_cluster_probe.json
  index_year_diff_plus_1_cluster_probe.json
  lookatnetworks_form_open_hang_probe.json (PR AA)
```

Plus the memory note at `memory/reference_linklistinit_fast_
path.md` (out of tree).

## Recommended morning sequence

1. **Quick scan**: read this log + the four "headline
   findings" sections in their respective MDs.
2. **Decide on F2 bucket rename** — it's the biggest
   conceptual shift; would propagate into the maintainer
   report appendix.  Cause-summary JSON edit is one-shot.
3. **Decide on F1 / F3 promotions** — if any go to ISSUES,
   that's its own small PR.
4. **Run the validation gauntlet** to confirm post-merge
   state matches expectation:

   ```
   git pull --ff-only
   python -m pytest tests/ -W ignore
   python analysis/run_all_audits.py --ci
   python reports/generate_report.py     # idempotent — should leave clean status
   ```

5. **Decide whether to attempt PR-AB-as-originally-scoped**:
   smaller LookAtNetworks fixture so CmdRun completes.
   Uses Access COM; needs supervision.

## Self-assessment

What went well:
- All 9 PRs committed, pushed, merged with green tests.
- Several substantive findings that would have required
  several daytime sessions to discover.
- Branch hygiene clean (no orphan branches, no force pushes).

What went less well:
- PR AA's first probe attempts popped 2 modal Access
  dialogs because of a LinkListInit oversight (now memorialised).
  Fixed mid-PR; cost ~30 min.
- The export coverage matrix (PR AD) over-attributes
  `manifest` depth; documented as a caveat.

No regressions introduced.  All branches deleted; only `main`
remains tracking `origin/main`.
