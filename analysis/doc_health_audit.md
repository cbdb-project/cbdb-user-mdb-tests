# Doc health audit (PR AL)

Static scan of all `.md` docs in the repo for known-stale phrasing patterns and broken local file links.  No edits made; this is a worklist for morning review.

## Headline

- Docs scanned: 25
- Findings: **21**

- By kind:
  - `stale_claim_pre_PR_AA`: 16
  - `candidate_label_overtaken_by_PR_AI_AJ`: 4
  - `stale_label_pre_PR_X`: 1

## `AGENTS.md` × 2

- line 137 (`stale_claim_pre_PR_AA`)
  - match: `Form_Open hangs`
  - excerpt: ### 3. NAVIGATION_PANE.Form_Open hangs forever if `LinkListInit.c_path`
  - note: PR AA showed Form_Open does NOT hang; the actual blocker is CmdRun (Zhu Xi 2471 assocs)
- line 531 (`candidate_label_overtaken_by_PR_AI_AJ`)
  - match: `candidate_same_rule_tie_break_or_aggregation_diff`
  - excerpt: `candidate_same_rule_tie_break_or_aggregation_diff` (both
  - note: PR AI + AJ showed all 14 rows are upstream BIOG_MAIN / KIN_DATA drift, NOT tie-break.  Bucket label rename queued for morning.

## `OVERNIGHT_SHIP_LOG_2026-05-03.md` × 2

- line 50 (`stale_claim_pre_PR_AA`)
  - match: `Form_Open hangs`
  - excerpt: that `LookAtNetworks.Form_Open hangs in this driver` is
  - note: PR AA showed Form_Open does NOT hang; the actual blocker is CmdRun (Zhu Xi 2471 assocs)
- line 87 (`candidate_label_overtaken_by_PR_AI_AJ`)
  - match: `candidate_same_rule_tie_break_or_aggregation_diff`
  - excerpt: `candidate_same_rule_tie_break_or_aggregation_diff` to
  - note: PR AI + AJ showed all 14 rows are upstream BIOG_MAIN / KIN_DATA drift, NOT tie-break.  Bucket label rename queued for morning.

## `README.md` × 5

- line 239 (`stale_claim_pre_PR_AA`)
  - match: `Form_Open hangs`
  - excerpt: | LookAtNetworks         | ⏭ skipped (Form_Open hangs) | — | ⏭ Recall (Form_Open hangs) | ⏭ ImportPeople / ImportPlaces (Form_Open) | — (no save button) | recursive expansion (Zhu Xi 2 471 assocs); bl
  - note: PR AA showed Form_Open does NOT hang; the actual blocker is CmdRun (Zhu Xi 2471 assocs)
- line 239 (`stale_claim_pre_PR_AA`)
  - match: `Form_Open hangs`
  - excerpt: | LookAtNetworks         | ⏭ skipped (Form_Open hangs) | — | ⏭ Recall (Form_Open hangs) | ⏭ ImportPeople / ImportPlaces (Form_Open) | — (no save button) | recursive expansion (Zhu Xi 2 471 assocs); bl
  - note: PR AA showed Form_Open does NOT hang; the actual blocker is CmdRun (Zhu Xi 2471 assocs)
- line 306 (`stale_claim_pre_PR_AA`)
  - match: `Form_Open hangs`
  - excerpt: | 7 | 🟡 partial | `tests/test_vba_matrix_hard_forms.py` (added 2026-05-02) handles 2 of the 3 forms with hand-picked tiny fixtures: LookAtGroupData (c_person_id=1, 2 entries / 2 statuses — backfill ch
  - note: PR AA showed Form_Open does NOT hang; the actual blocker is CmdRun (Zhu Xi 2471 assocs)
- line 310 (`stale_claim_pre_PR_AA`)
  - match: `Form_Open hangs`
  - excerpt: | 11 | ✅ done | Bilingual UI test (`tests/test_vba_bilingual_ui.py`) — for each of the 9 forms with the standard `CmdFanti` / `CmdJianti` toggle pair (Networks uses different names + Form_Open hangs),
  - note: PR AA showed Form_Open does NOT hang; the actual blocker is CmdRun (Zhu Xi 2471 assocs)
- line 314 (`stale_claim_pre_PR_AA`)
  - match: `Form_Open hangs`
  - excerpt: | 15 | ✅ done | `CmdStoreID` / `CmdRecallID` round-trip (`tests/test_vba_storeid_recallid.py`) — covers all 7 query-runnable forms for Store, 3 of 4 forms for Recall (Networks Form_Open hangs in this 
  - note: PR AA showed Form_Open does NOT hang; the actual blocker is CmdRun (Zhu Xi 2471 assocs)

## `analysis\hard_form_skip_inventory.md` × 8

- line 70 (`stale_claim_pre_PR_AA`)
  - match: `Form_Open hangs`
  - excerpt: - reason: LookAtNetworks Form_Open hangs in this driver —
  - note: PR AA showed Form_Open does NOT hang; the actual blocker is CmdRun (Zhu Xi 2471 assocs)
- line 73 (`stale_claim_pre_PR_AA`)
  - match: `Form_Open hangs`
  - excerpt: - reason: LookAtNetworks Form_Open hangs in this driver
  - note: PR AA showed Form_Open does NOT hang; the actual blocker is CmdRun (Zhu Xi 2471 assocs)
- line 76 (`stale_claim_pre_PR_AA`)
  - match: `Form_Open hangs`
  - excerpt: - reason: LookAtNetworks Form_Open hangs in this driver
  - note: PR AA showed Form_Open does NOT hang; the actual blocker is CmdRun (Zhu Xi 2471 assocs)
- line 82 (`stale_claim_pre_PR_AA`)
  - match: `Form_Open hangs`
  - excerpt: - reason: LookAtNetworks Form_Open hangs in this driver
  - note: PR AA showed Form_Open does NOT hang; the actual blocker is CmdRun (Zhu Xi 2471 assocs)
- line 191 (`stale_claim_pre_PR_AA`)
  - match: `Form_Open hangs`
  - excerpt: - `tests\test_vba_bug_behaviors.py:270` — LookAtNetworks Form_Open hangs in this driver —
  - note: PR AA showed Form_Open does NOT hang; the actual blocker is CmdRun (Zhu Xi 2471 assocs)
- line 192 (`stale_claim_pre_PR_AA`)
  - match: `Form_Open hangs`
  - excerpt: - `tests\test_vba_cmdguess_cross_form.py:50` — LookAtNetworks Form_Open hangs in this driver
  - note: PR AA showed Form_Open does NOT hang; the actual blocker is CmdRun (Zhu Xi 2471 assocs)
- line 193 (`stale_claim_pre_PR_AA`)
  - match: `Form_Open hangs`
  - excerpt: - `tests\test_vba_import_lists.py:134` — LookAtNetworks Form_Open hangs in this driver
  - note: PR AA showed Form_Open does NOT hang; the actual blocker is CmdRun (Zhu Xi 2471 assocs)
- line 194 (`stale_claim_pre_PR_AA`)
  - match: `Form_Open hangs`
  - excerpt: - `tests\test_vba_storeid_recallid.py:224` — LookAtNetworks Form_Open hangs in this driver
  - note: PR AA showed Form_Open does NOT hang; the actual blocker is CmdRun (Zhu Xi 2471 assocs)

## `analysis\index_drift_cause_analysis.md` × 1

- line 245 (`stale_label_pre_PR_X`)
  - match: `blocked_by_missing_frmBaseMaintenance_vba`
  - excerpt: `blocked_by_missing_frmBaseMaintenance_vba` (factually
  - note: PR X renamed to blocked_by_runtime_priority_triage_pending

## `analysis\lookatnetworks_form_open_hang.md` × 1

- line 134 (`stale_claim_pre_PR_AA`)
  - match: `Form_Open hangs`
  - excerpt: behaviors.py:270` from `"LookAtNetworks Form_Open hangs"`
  - note: PR AA showed Form_Open does NOT hang; the actual blocker is CmdRun (Zhu Xi 2471 assocs)

## `reports\CBDB_Issues_Report_EN.md` × 1

- line 659 (`candidate_label_overtaken_by_PR_AI_AJ`)
  - match: `candidate_same_rule_tie_break_or_aggregation_diff`
  - excerpt: - `consistent_within_rule` × 14 → 5 signature groups, all `candidate_same_rule_tie_break_or_aggregation_diff`.  Recurring diff=-20 across Rules 11/13/15/19 is the standout pattern.
  - note: PR AI + AJ showed all 14 rows are upstream BIOG_MAIN / KIN_DATA drift, NOT tie-break.  Bucket label rename queued for morning.

## `reports\CBDB_Issues_Report_ZH-Hant.md` × 1

- line 659 (`candidate_label_overtaken_by_PR_AI_AJ`)
  - match: `candidate_same_rule_tie_break_or_aggregation_diff`
  - excerpt: - `consistent_within_rule` × 14 → 5 個 signature 分組，全部標為 `candidate_same_rule_tie_break_or_aggregation_diff`。最顯眼的是 Rule 11/13/15/19 反覆出現的 diff=-20。
  - note: PR AI + AJ showed all 14 rows are upstream BIOG_MAIN / KIN_DATA drift, NOT tie-break.  Bucket label rename queued for morning.
