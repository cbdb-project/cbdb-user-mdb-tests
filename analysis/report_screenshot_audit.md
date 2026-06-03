# Report screenshot consistency audit

**Generated:** 2026-06-03T09:56:10+00:00
**Generator:** `analysis/audit_report_screenshot_consistency.py`
**Companion JSON:** `reports/report_screenshot_audit.json`

This audit cross-checks every screenshot caption in `reports/generate_report.py::ISSUES` against the surrounding issue's tier and declared latency state.  Conservative text-only rules; no MDB, no Access COM, no NLP.

## Rules

**Rule A — P5 only.** If a screenshot filename contains any of `popup, runtime, form_open, annotated`, its caption must contain at least one hedge keyword: `Hypothetical, latent, cannot trigger, can't trigger, not currently triggerable, if ... existed, if Issue #, users currently CAN'T trigger, 潛伏, 潜伏, 目前不能觸發, 目前不能触发, 目前無法觸發, 目前无法触发, 不可達, 不可达, 假設, 假设`.  Rationale: the typical bug-report pattern is annotated runtime + faux popup; for P5 issues that are gated/dormant today, the caption MUST disclaim that the popup is hypothetical or that the trigger isn't currently reachable.

**Rule B — all issues.** If a caption contains any active-trigger phrase (`users see, popup users see, the popup users see, appears immediately, 立即彈出, 立即弹出`) AND the issue's title/summary/steps/severity contains any latency marker (`LATENT, DORMANT` case-sensitive, plus `cannot trigger, can't trigger, CAN'T trigger, no UI trigger, not currently triggerable, 潛伏, 潜伏, 不可達, 不可达, 目前不能觸發, 目前不能触发, 目前無法觸發, 目前无法触发` case-insensitive), flag.  Rationale: the literal self-contradiction pattern (Issue #9 class).

**Out of scope by design:**
- **P3 missing UI** is exempt from Rule A — runtime captures are how we PROVE a button is missing.
- **P0 / P1 / P2** are exempt from Rule A — faux popups + annotated runtime are the standard reporting pattern when the bug is user-visible today.
- No MDB / Access COM checks, no image-content inference.

## Summary

| Metric | Value |
|---|---:|
| Issues total | 0 |
| Issues with at least one screenshot | 0 |
| Screenshots inspected | 0 |
| Rule A violations (P5) | 0 |
| Rule B violations (all) | 0 |
| **Mismatches** | **0** |

## Mismatches

**Clean — no mismatches.** Every screenshot caption is consistent with its issue's tier and latency state under the active rules.

## Per-screenshot inventory

| Issue | Tier | Filename | Filename trigger | Caption hedge | Active phrase | Issue latent? | Rule A | Rule B |
|---:|---|---|---|---|---|---|---|---|
