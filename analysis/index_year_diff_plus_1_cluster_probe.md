# Rule 11/11 diff=+1 cluster probe (PR AJ)

Companion to PR AI.  Tests whether the 5-row K2 diff=+1 sub-cluster shares the same upstream BIOG_MAIN.c_birthyear drift mechanism as the diff=-20 cluster did, just in the opposite direction.

## Verdict: `different_mechanism — see outcomes detail`

Outcome counts:
- `different_evidence_pid_set_between_sides`: 4
- `source_data_drift_biog_main_birthyear`: 1

## Per-row detail

### `c_personid = 523820` (金文伯 / Jin Wenbo)
- Access actual: 1398, PHP actual: 1397, diff: 1
- User winner: pid=34486 (金善), birthyear=1368
- SQLite winner: pid=66728 (金善), birthyear=1367
- Outcome: **different_evidence_pid_set_between_sides**

### `c_personid = 523821` (金武伯 / Jin Wubo)
- Access actual: 1398, PHP actual: 1397, diff: 1
- User winner: pid=34486 (金善), birthyear=1368
- SQLite winner: pid=66728 (金善), birthyear=1367
- Outcome: **different_evidence_pid_set_between_sides**

### `c_personid = 523822` (金堅伯 / Jin Jianbo)
- Access actual: 1398, PHP actual: 1397, diff: 1
- User winner: pid=34486 (金善), birthyear=1368
- SQLite winner: pid=66728 (金善), birthyear=1367
- Outcome: **different_evidence_pid_set_between_sides**

### `c_personid = 523823` (金壽伯 / Jin Shoubo)
- Access actual: 1398, PHP actual: 1397, diff: 1
- User winner: pid=34486 (金善), birthyear=1368
- SQLite winner: pid=66728 (金善), birthyear=1367
- Outcome: **different_evidence_pid_set_between_sides**

### `c_personid = 696896` (陳瑜 / Chen Yu)
- Access actual: 1346, PHP actual: 1345, diff: 1
- User winner: pid=108002 (陳嗣), birthyear=1316
- SQLite winner: pid=108002 (陳嗣), birthyear=1315
- BIOG_MAIN drift pid=108002: User=1316 / SQLite=1315 (diff=1)
- Outcome: **source_data_drift_biog_main_birthyear**

## Cumulative picture across PR AI + AJ

With both probes:
- diff=-20 sub-cluster (9 rows): 7 source_data_drift_biog_main_birthyear + 2 different_evidence_pid_set
- diff=+1 sub-cluster (5 rows): see above

Combined, the K2 `consistent_within_rule × 14` bucket is overwhelmingly upstream-data drift, not algorithm divergence.  Cause-summary JSON re-class left for morning review (per overnight rules: don't change severity without overwhelming evidence + maintainer sign-off).