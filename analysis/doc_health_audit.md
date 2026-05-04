# Doc health audit (PR AL)

Static scan of all `.md` docs in the repo for known-stale phrasing patterns and broken local file links.  No edits made; this is a worklist for morning review.

## Headline

- Docs scanned: 24
- Findings: **6**

- By kind:
  - `stale_claim_pre_PR_AA`: 3
  - `candidate_label_overtaken_by_PR_AI_AJ`: 2
  - `stale_label_pre_PR_X`: 1

## `AGENTS.md` × 2

- line 137 (`stale_claim_pre_PR_AA`)
  - match: `Form_Open hangs`
  - excerpt: ### 3. NAVIGATION_PANE.Form_Open hangs forever if `LinkListInit.c_path`
  - note: PR AA showed Form_Open does NOT hang; the actual blocker is CmdRun (Zhu Xi 2471 assocs)
- line 530 (`candidate_label_overtaken_by_PR_AI_AJ`)
  - match: `candidate_same_rule_tie_break_or_aggregation_diff`
  - excerpt: `candidate_same_rule_tie_break_or_aggregation_diff` (both
  - note: PR AI + AJ showed all 14 rows are upstream BIOG_MAIN / KIN_DATA drift, NOT tie-break.  Bucket label rename queued for morning.

## `OVERNIGHT_SHIP_LOG_2026-05-03.md` × 3

- line 241 (`stale_label_pre_PR_X`)
  - match: `blocked_by_missing_frmBaseMaintenance_vba`
  - excerpt: - **1** stray `blocked_by_missing_frmBaseMaintenance_vba`
  - note: PR X renamed to blocked_by_runtime_priority_triage_pending
- line 50 (`stale_claim_pre_PR_AA`)
  - match: `Form_Open hangs`
  - excerpt: that `LookAtNetworks.Form_Open hangs in this driver` is
  - note: PR AA showed Form_Open does NOT hang; the actual blocker is CmdRun (Zhu Xi 2471 assocs)
- line 87 (`candidate_label_overtaken_by_PR_AI_AJ`)
  - match: `candidate_same_rule_tie_break_or_aggregation_diff`
  - excerpt: `candidate_same_rule_tie_break_or_aggregation_diff` to
  - note: PR AI + AJ showed all 14 rows are upstream BIOG_MAIN / KIN_DATA drift, NOT tie-break.  Bucket label rename queued for morning.

## `analysis\lookatnetworks_form_open_hang.md` × 1

- line 134 (`stale_claim_pre_PR_AA`)
  - match: `Form_Open hangs`
  - excerpt: behaviors.py:270` from `"LookAtNetworks Form_Open hangs"`
  - note: PR AA showed Form_Open does NOT hang; the actual blocker is CmdRun (Zhu Xi 2471 assocs)
