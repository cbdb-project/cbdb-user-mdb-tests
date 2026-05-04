# Index_year diff=-20 cluster deep-dive (PR AI)

PR Y suggested investigation #4: walk the 9 K2 rows in the `consistent_within_rule` bucket whose `(php_tcode, access_tcode, diff)` signature has `diff = -20` across rules 11/13/15/19, and confirm or refute the working hypothesis that they share a single staging-step row pick.

## Verdict: `hypothesis_REVISED — actual cause is upstream BIOG_MAIN.c_birthyear drift between User MDB and SQLite snapshot, NOT staging-step row pick`

Outcome counts:
- `source_data_drift_biog_main_birthyear`: 7
- `different_evidence_pid_set_between_sides`: 2

## Per-row detail

### `c_personid = 228114` (王淑抃 / Wang Shubian) — rule 11

- Rule intent: child = father.c_birthyear + 30
- Access actual index_year: 1587
- PHP actual index_year: 1607  (diff = -20)
- User-side evidence rows: 1
- SQLite-side evidence rows: 1
- Same evidence pid set on both sides: True
- BIOG_MAIN.c_birthyear drift per evidence pid:
  - pid=123710: User=1557 / SQLite=1577 (diff=-20)
- **Winner-pid birthyear drift fully explains the index_year diff.**
- User-inferred winner: pid=123710 (王圖), birthyear=1557 → inferred index_year = 1587
- SQLite-inferred winner: pid=123710 (王圖), birthyear=1577 → inferred index_year = 1607
- Outcome: **source_data_drift_biog_main_birthyear**

### `c_personid = 527169` (王毓玄 / Wang Yuxuan) — rule 11

- Rule intent: child = father.c_birthyear + 30
- Access actual index_year: 1587
- PHP actual index_year: 1607  (diff = -20)
- User-side evidence rows: 1
- SQLite-side evidence rows: 1
- Same evidence pid set on both sides: True
- BIOG_MAIN.c_birthyear drift per evidence pid:
  - pid=123710: User=1557 / SQLite=1577 (diff=-20)
- **Winner-pid birthyear drift fully explains the index_year diff.**
- User-inferred winner: pid=123710 (王圖), birthyear=1557 → inferred index_year = 1587
- SQLite-inferred winner: pid=123710 (王圖), birthyear=1577 → inferred index_year = 1607
- Outcome: **source_data_drift_biog_main_birthyear**

### `c_personid = 228102` (王邦憲 / Wang Bangxian) — rule 13

- Rule intent: father = MIN(child.c_birthyear) - 30
- Access actual index_year: 1527
- PHP actual index_year: 1547  (diff = -20)
- User-side evidence rows: 2
- SQLite-side evidence rows: 2
- Same evidence pid set on both sides: True
- BIOG_MAIN.c_birthyear drift per evidence pid:
  - pid=123710: User=1557 / SQLite=1577 (diff=-20)
- **Winner-pid birthyear drift fully explains the index_year diff.**
- User-inferred winner: pid=123710 (王圖), birthyear=1557 → inferred index_year = 1527
- SQLite-inferred winner: pid=123710 (王圖), birthyear=1577 → inferred index_year = 1547
- Outcome: **source_data_drift_biog_main_birthyear**

### `c_personid = 248797` (武英 / Wu Ying) — rule 13

- Rule intent: father = MIN(child.c_birthyear) - 30
- Access actual index_year: 1402
- PHP actual index_year: 1422  (diff = -20)
- User-side evidence rows: 1
- SQLite-side evidence rows: 1
- Same evidence pid set on both sides: False
- User-inferred winner: pid=198975 (武清), birthyear=1432 → inferred index_year = 1402
- SQLite-inferred winner: pid=199710 (武清), birthyear=1452 → inferred index_year = 1422
- Outcome: **different_evidence_pid_set_between_sides**

### `c_personid = 228104` (某氏(王圖母) / Mou Shi (Mother of Wangtu)) — rule 15

- Rule intent: mother = MIN(child.c_birthyear) - 27
- Access actual index_year: 1530
- PHP actual index_year: 1550  (diff = -20)
- User-side evidence rows: 1
- SQLite-side evidence rows: 1
- Same evidence pid set on both sides: True
- BIOG_MAIN.c_birthyear drift per evidence pid:
  - pid=123710: User=1557 / SQLite=1577 (diff=-20)
- **Winner-pid birthyear drift fully explains the index_year diff.**
- User-inferred winner: pid=123710 (王圖), birthyear=1557 → inferred index_year = 1530
- SQLite-inferred winner: pid=123710 (王圖), birthyear=1577 → inferred index_year = 1550
- Outcome: **source_data_drift_biog_main_birthyear**

### `c_personid = 248799` (張氏(武清母) / Zhang Shi (Mother of Wuqing)) — rule 15

- Rule intent: mother = MIN(child.c_birthyear) - 27
- Access actual index_year: 1405
- PHP actual index_year: 1425  (diff = -20)
- User-side evidence rows: 1
- SQLite-side evidence rows: 1
- Same evidence pid set on both sides: False
- User-inferred winner: pid=198975 (武清), birthyear=1432 → inferred index_year = 1405
- SQLite-inferred winner: pid=199710 (武清), birthyear=1452 → inferred index_year = 1425
- Outcome: **different_evidence_pid_set_between_sides**

### `c_personid = 228111` (王萃 / Wang Cui) — rule 19

- Rule intent: older brother = MAX(sibling.c_birthyear) + 2
- Access actual index_year: 1559
- PHP actual index_year: 1579  (diff = -20)
- User-side evidence rows: 2
- SQLite-side evidence rows: 2
- Same evidence pid set on both sides: True
- BIOG_MAIN.c_birthyear drift per evidence pid:
  - pid=123710: User=1557 / SQLite=1577 (diff=-20)
- **Winner-pid birthyear drift fully explains the index_year diff.**
- User-inferred winner: pid=123710 (王圖), birthyear=1557 → inferred index_year = 1559
- SQLite-inferred winner: pid=123710 (王圖), birthyear=1577 → inferred index_year = 1579
- Outcome: **source_data_drift_biog_main_birthyear**

### `c_personid = 228112` (王圍 / Wang Wei) — rule 19

- Rule intent: older brother = MAX(sibling.c_birthyear) + 2
- Access actual index_year: 1559
- PHP actual index_year: 1579  (diff = -20)
- User-side evidence rows: 1
- SQLite-side evidence rows: 1
- Same evidence pid set on both sides: True
- BIOG_MAIN.c_birthyear drift per evidence pid:
  - pid=123710: User=1557 / SQLite=1577 (diff=-20)
- **Winner-pid birthyear drift fully explains the index_year diff.**
- User-inferred winner: pid=123710 (王圖), birthyear=1557 → inferred index_year = 1559
- SQLite-inferred winner: pid=123710 (王圖), birthyear=1577 → inferred index_year = 1579
- Outcome: **source_data_drift_biog_main_birthyear**

### `c_personid = 228113` (王海實 / Wang Haishi) — rule 19

- Rule intent: older brother = MAX(sibling.c_birthyear) + 2
- Access actual index_year: 1559
- PHP actual index_year: 1579  (diff = -20)
- User-side evidence rows: 1
- SQLite-side evidence rows: 1
- Same evidence pid set on both sides: True
- BIOG_MAIN.c_birthyear drift per evidence pid:
  - pid=123710: User=1557 / SQLite=1577 (diff=-20)
- **Winner-pid birthyear drift fully explains the index_year diff.**
- User-inferred winner: pid=123710 (王圖), birthyear=1557 → inferred index_year = 1559
- SQLite-inferred winner: pid=123710 (王圖), birthyear=1577 → inferred index_year = 1579
- Outcome: **source_data_drift_biog_main_birthyear**

## Implications for PR Y

PR Y's `consistent_within_rule` × 14 bucket (at confidence `medium`) had this 9-row -20 sub-cluster as its top suggested next investigation.  Result: `hypothesis_REVISED — actual cause is upstream BIOG_MAIN.c_birthyear drift between User MDB and SQLite snapshot, NOT staging-step row pick`.

Per the result, the cause-summary JSON's confidence for the diff=-20 cluster can be promoted from `medium` to either `supported_by_focused_probe` (if cleanly supported) or annotated with the partial-support outcome breakdown.  Update is left to a follow-up PR for morning review.