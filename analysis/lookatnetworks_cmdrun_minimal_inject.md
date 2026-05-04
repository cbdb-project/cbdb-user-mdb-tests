# LookAtNetworks CmdRun probe — minimal injection (PR AX)

Bisection arc complete: PR AR diagnosis → PR AS-AV ruling out triggers → PR AW finding W3 (keep loaded) and PR AU V12 (minimal injection) as the two viable paths.  This PR uses the minimal-injection path on PR AQ's 3 anchor candidates.

## Setup (constant across candidates)

- skip_inject_autodetect_forms = all 9 sibling Form_LookAt* (Networks autodetect ON only)
- minimal kin-only control state:
  - `TxtNodeDist` = `1`
  - `TxtMaxLoop` = `0`
  - `ChkKin` = `-1`
  - `ChkNonKin` = `0`
  - `ChkMale` = `-1`
  - `ChkFemale` = `-1`
- result_table = ZZ_SOCIAL_NETWORK
- CmdRun timer cap = 120 s; per-candidate hard cap = 240 s

## Outcomes

| pid | name | outcome | elapsed | ZZ_SOCIAL_NETWORK | ZZ_SCRATCH_PEOPLE |
|---:|---|---|---:|---:|---:|
| 30270 | 曹植 (Cao Zhi) | `succeeded_with_output` | 7.86s | 1 | 2 |
| 4 | 查道 (Zha Dao) | `succeeded_with_output` | 7.79s | 14 | 10 |
| 3135 | 張君平 (Zhang Junping) | `succeeded_with_output` | 7.97s | 2 | 3 |

## Headline verdict: `SUCCESS_at_least_one_anchor_completed`

3 of 3 anchors completed CmdRun with non-empty bounded output under 120 s.  Recommend PR AY to add a small-fixture LookAtNetworks test using one of these anchors + the minimal-injection helper.

## Per-candidate detail

### `c_personid = 30270` (曹植 / Cao Zhi)

- est_1hop_assoc_total (PR AQ): 10
- assocs=10, kin=1
- elapsed: 7.86s
- outcome: **`succeeded_with_output`**
- row counts:
  - `ZZ_SOCIAL_NETWORK_via_click_via_timer`: 1
  - `ZZ_SOCIAL_NETWORK`: 1
  - `ZZ_SCRATCH_PEOPLE`: 2
  - `ZZ_SOCIAL_NETWORK_AGGREGATE`: 1
- controls_set: `{'TxtNodeDist': 1, 'TxtMaxLoop': 0, 'ChkKin': -1, 'ChkNonKin': 0, 'ChkMale': -1, 'ChkFemale': -1}`
- ZZ_TEST_DEBUG (2 entries):
  - 1: `LookAtNetworks:ENTER`
  - 2: `LookAtNetworks:DONE`
- markers:
  - +0.0s constructing_session_minimal_inject
  - +0.0s opening_session
  - +2.49s session_opened_only_Networks_injected
  - +3.64s form_opened
  - +3.86s picker_seeded
  - +4.0s controls_set
  - +5.53s cmdrun_returned_1_rows_via_timer
  - +5.53s row_counts_captured
  - +5.53s debug_captured

### `c_personid = 4` (查道 / Zha Dao)

- est_1hop_assoc_total (PR AQ): 99
- assocs=5, kin=9
- elapsed: 7.79s
- outcome: **`succeeded_with_output`**
- row counts:
  - `ZZ_SOCIAL_NETWORK_via_click_via_timer`: 14
  - `ZZ_SOCIAL_NETWORK`: 14
  - `ZZ_SCRATCH_PEOPLE`: 10
  - `ZZ_SOCIAL_NETWORK_AGGREGATE`: 14
- controls_set: `{'TxtNodeDist': 1, 'TxtMaxLoop': 0, 'ChkKin': -1, 'ChkNonKin': 0, 'ChkMale': -1, 'ChkFemale': -1}`
- ZZ_TEST_DEBUG (2 entries):
  - 1: `LookAtNetworks:ENTER`
  - 2: `LookAtNetworks:DONE`
- markers:
  - +0.0s constructing_session_minimal_inject
  - +0.0s opening_session
  - +2.16s session_opened_only_Networks_injected
  - +3.32s form_opened
  - +3.32s picker_seeded
  - +3.46s controls_set
  - +5.49s cmdrun_returned_14_rows_via_timer
  - +5.49s row_counts_captured
  - +5.49s debug_captured

### `c_personid = 3135` (張君平 / Zhang Junping)

- est_1hop_assoc_total (PR AQ): 31
- assocs=5, kin=2
- elapsed: 7.97s
- outcome: **`succeeded_with_output`**
- row counts:
  - `ZZ_SOCIAL_NETWORK_via_click_via_timer`: 2
  - `ZZ_SOCIAL_NETWORK`: 2
  - `ZZ_SCRATCH_PEOPLE`: 3
  - `ZZ_SOCIAL_NETWORK_AGGREGATE`: 2
- controls_set: `{'TxtNodeDist': 1, 'TxtMaxLoop': 0, 'ChkKin': -1, 'ChkNonKin': 0, 'ChkMale': -1, 'ChkFemale': -1}`
- ZZ_TEST_DEBUG (2 entries):
  - 1: `LookAtNetworks:ENTER`
  - 2: `LookAtNetworks:DONE`
- markers:
  - +0.0s constructing_session_minimal_inject
  - +0.0s opening_session
  - +2.12s session_opened_only_Networks_injected
  - +5.01s form_opened
  - +5.02s picker_seeded
  - +5.15s controls_set
  - +5.68s cmdrun_returned_2_rows_via_timer
  - +5.68s row_counts_captured
  - +5.68s debug_captured

## Implications

**The minimal-injection path is viable** for Networks tests.  Recommended follow-up (separate PR AY): add a small-fixture `tests/test_vba_networks_small_fixture.py` that uses one of the successful anchors and the same minimal-injection setup, gated by `--include-vba`.  Use the existing `test_vba_matrix_hard_forms.py` patterns for fixture conventions; do NOT unskip the matrix test in the same PR.

Operating recipe to encode in AY:
1. Construct VbaSession with the 9-sibling skip set (use the same SKIP_SIBLINGS constant as this probe).
2. Open LookAtNetworks (~2 s).
3. Seed picker (`ZZ_SCRATCH_IMPORT_PEOPLE`).
4. Set the 6 minimal-control values above.
5. `click_via_timer` for CmdRun, 120 s cap.
6. Assert ZZ_SOCIAL_NETWORK row count > 0 (or another sanity bound from the per-anchor outcomes captured here).

## Constraints respected per AX brief
- Probe-only.  No matrix unskips.
- Minimal injection (only Form_LookAtNetworks).
- Did NOT use W3 keep-loaded path.
- Per-candidate fresh Access process.
- 120 s CmdRun timer cap + 240 s per-candidate outer cap.
- Killed orphan MSACCESS.EXE between candidates.
- Default driver behaviour preserved (probe uses PR AT `skip_inject_autodetect_forms` kwarg only).
- Fast suite still 111 passed, 9 skipped (pre-probe).