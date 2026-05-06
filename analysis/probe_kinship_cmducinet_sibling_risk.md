# LookAtKinship × CmdUCINet — sibling-risk probe (Issue #22 family)

**Date:** 2026-05-06  ·  **Branch:** `investigate/kinship-cmducinet-sibling-risk`

## Headline

Conclusion: **`same_bug_family_runtime_confirmed`**

Kinship.CmdUCINet reproduces the SAME failure mode as Issue #22's Associations case: `LookAtKinship:ERR Invalid procedure call or argument` (VBA error 5) fires mid-write, leaving a partial `.vna` file with `*node data` complete + `*node properties` truncated (header written but ZERO data rows -- bail on the FIRST WriteLine attempt) + `*tie data` never written.

Why the iteration-order 'next row' lookup is moot here: when *node properties* bails on the very first row, there is no last-written anchor to compute the 'next row' against.  But the substantive evidence is just as conclusive: ZZ_SCRATCH_KIN has 1 non-cp1252-without-substitute c_kin_name row(s); the FIRST row in c_kin_id ASC order IS a trigger (c_kin_id 140733, c_kin_name 'He Mou 取' with non-cp1252 char(s) [('取', 'U+53D6')]); the bail happens on row 0's WriteLine.  Combined with the matching partial-file shape and the matching VBA error 5 wording, this is the same bug class as Issue #22 -- only the specific row at which the bail fires differs (Associations bailed at row 3974 of 8087; Kinship bailed at row 1 of 6).

The current Kinship × CmdUCINet coverage (`tests/test_vba_cmducinet_kinship.py`) passes only because the chosen fixture (person 3211) happens to have no Han-name members in its kin network -- switching to a fixture that DOES (this probe used picker = 152930, He Jing 何淨, whose sole 1-hop kin is pid 140733 He Mou 取) reproduces the same VBA error 5.  The Issue #22 canonical entry's 'sibling-form risk' paragraph for Kinship is now runtime-verified, not just statically inferred.

---

## Probe-observed facts

- Outcome: `reproduced_invalid_procedure_call`
- Elapsed: 12.67 s
- Picker pid: `152930` (He Jing 何淨 — chosen because their sole 1-hop kin row points to pid 140733 He Mou 取, U+53D6 Han ideograph)
- Row counts after CmdRun: `ZZ_SCRATCH_KIN=6` `ZZ_SCRATCH_KINNET=5`
- File: `343 bytes` at `C:\Users\how612\Documents\GitHub\cbdb-user-mdb-tests\analysis\_probe_kin_ucinet_sibling_out\cmducinet_kinship_sibling.vna`

### Partial file structure

| section | header | written rows | last row's first token |
|---|---|---:|---|
| `*node data` | `ID index_year dy_code dynasty sex x_coord y_coord kindist` | 6 | `152934` |
| `*node properties` | `ID color shape size shortlabel` | 0 | `None` |

### ZZ_SCRATCH_KIN.c_kin_name scan (non-cp1252 without FSO substitute)

- Scanned rows: 6
- Rows with non-cp1252-without-substitute c_kin_name: 1
- Trigger pid 140733 present in ZZ_SCRATCH_KIN: **True**
- First bad index (in c_kin_id ASC order, 0-based): 0

Samples:

| kid_idx | c_kin_id | c_kin_name | non-cp1252 chars |
|---:|---:|---|---|
| 0 | 140733 | `He Mou 取` | `取` (U+53D6) |

### Row immediately after the last successfully-written *node properties* row

(no next-row lookup result; see markers / exception)

### ZZ_TEST_DEBUG transcript

- `  1`: `LookAtKinship:ENTER`
- `  2`: `LookAtKinship:DONE`
- `  3`: `LookAtKinship:ERR Invalid procedure call or argument`

### Markers timeline

- `+  0.00s` constructing_session
- `+  6.30s` session_opened_attempt_1
- `+  6.81s` filedialog_patched
- `+  7.60s` picker_seeded_pid_152930
- `+  8.17s` cmdrun_returned_6
- `+  8.17s` row_counts_kin_6_kinnet_5
- `+  8.17s` scratch_kin_scan_done_total_6_bad_1_trigger_in_scratch_True
- `+ 10.21s` cmducinet_fired_no_wait
- `+ 10.21s` file_appeared_True
- `+ 10.21s` partial_file_parsed
- `+ 10.21s` debug_captured

---

## Inferences for canonicalization / coverage follow-up

These are CONCLUSIONS drawn from the probe + static evidence, NOT additional probe observations.  Re-verify before relying.

**Conclusion bucket: `same_bug_family_runtime_confirmed`**

Kinship.CmdUCINet reproduces the SAME failure mode as Issue #22's Associations case: `LookAtKinship:ERR Invalid procedure call or argument` (VBA error 5) fires mid-write, leaving a partial `.vna` file with `*node data` complete + `*node properties` truncated (header written but ZERO data rows -- bail on the FIRST WriteLine attempt) + `*tie data` never written.

Why the iteration-order 'next row' lookup is moot here: when *node properties* bails on the very first row, there is no last-written anchor to compute the 'next row' against.  But the substantive evidence is just as conclusive: ZZ_SCRATCH_KIN has 1 non-cp1252-without-substitute c_kin_name row(s); the FIRST row in c_kin_id ASC order IS a trigger (c_kin_id 140733, c_kin_name 'He Mou 取' with non-cp1252 char(s) [('取', 'U+53D6')]); the bail happens on row 0's WriteLine.  Combined with the matching partial-file shape and the matching VBA error 5 wording, this is the same bug class as Issue #22 -- only the specific row at which the bail fires differs (Associations bailed at row 3974 of 8087; Kinship bailed at row 1 of 6).

The current Kinship × CmdUCINet coverage (`tests/test_vba_cmducinet_kinship.py`) passes only because the chosen fixture (person 3211) happens to have no Han-name members in its kin network -- switching to a fixture that DOES (this probe used picker = 152930, He Jing 何淨, whose sole 1-hop kin is pid 140733 He Mou 取) reproduces the same VBA error 5.  The Issue #22 canonical entry's 'sibling-form risk' paragraph for Kinship is now runtime-verified, not just statically inferred.

**Suggested follow-up** (NOT autopiloted — each step needs its own brief):

- Tighten Issue #22's canonical text to promote Kinship from 'possible sibling risk / NOT yet probed' to 'sibling risk runtime-confirmed' OR file Issue #23 separately as a sibling P1 if the maintainer prefers per-form issues.  Either way, the static marker test in test_known_bugs.py for Bug #22 should be extended to also assert the same 2-arg `CreateTextFile(tFileName, True)` pattern in `Form_LookAtKinship.vb` (line ~2510).
- Decide whether the existing Kinship × CmdUCINet coverage test is still safe (currently it passes only because person 3211's network has no Han names) or whether it should be either (a) augmented with a fixture variant that DOES include Han names + asserts the expected ERR (turning it into a bug-pin like test_bug21 / test_bug22), or (b) marked as fixture-fragile in the docstring + inventory.
- Coordinate the upstream `CreateTextFile(..., True, True)` fix across BOTH forms in the same upstream patch (per the Issue #22 fix recommendation).

## Constraints honoured per brief

- ✅ Investigation artifacts only — no `tests/` or `cbdb_driver/*` changes
- ✅ Did NOT modify README / canonical reports / issue severity
- ✅ Did NOT do an upstream fix
- ✅ Used Access COM via `VbaSession.make_fixture`
- ✅ Probe-observed facts vs inferences explicitly separated
- ✅ Did NOT relax standards just because Kinship currently has coverage