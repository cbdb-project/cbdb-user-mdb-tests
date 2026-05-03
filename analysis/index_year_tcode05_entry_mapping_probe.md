# tcode='05' × 7 — focused probe verdict

PR Y flagged the `access_tcode='05'` × 7 sub-group of
`php_did_not_compute` as the cleanest next investigation:
**candidate_php_entry_code_mapping_gap** (medium-high
confidence).  The hypothesis was that PHP's
`sqlEntryRule('040101', 30, '05')` writes 0 / NULL for these 7
persons because the SQLite snapshot's `ENTRY_CODE_TYPE_REL`
doesn't classify their `c_entry_code` into `'040101'`.

PR Z (this probe — `analysis/probe_index_year_tcode05_entry_
mapping.py` → `reports/index_year_tcode05_entry_mapping_probe
.json`) walks each of the 7 personids on both sides:

  - User MDB `ENTRY_DATA` rows + `ENTRY_CODE_TYPE_REL`
    membership for each `c_entry_code`.
  - SQLite snapshot equivalents.
  - Whether PHP-shape reconstruction (`c_year - 30`) on the
    User-MDB-side row reproduces Access's stored
    `c_index_year`.

## Verdict — `hypothesis_mostly_supported`

| Outcome | Count |
| --- | ---: |
| `mapping_gap_confirmed_user_has_target_sqlite_missing` | **6** |
| `both_sides_have_target_cause_elsewhere` | 1 |

For 6 / 7 rows the verdict is direct: User MDB carries the
`c_entry_code → '040101'` mapping, SQLite does **not**.  PHP's
`sqlEntryRule('040101', 30, '05')` joins on the missing row, no
result, NULL written.  PR Y's `candidate_php_entry_code_mapping_
gap` label is **directly evidenced** for these 6.

The 1 outlier (`c_personid = 93384` / 張文伏 Zhang Wenfu) is
**also a PHP-side data gap**, just a different one:

  - User MDB `ENTRY_DATA` row: `c_entry_code = 36`,
    `c_year = 926`, `c_sequence = 1`.
  - SQLite `ENTRY_DATA` row: `c_entry_code = 36`, **`c_year =
    0`**, `c_sequence = 1`.
  - Both sides have entry_code 36 mapped to `'040101'` in
    `ENTRY_CODE_TYPE_REL`.
  - User-MDB-side PHP-shape reconstruction: `926 - 30 = 896`,
    matches Access's stored `c_index_year = 896` exactly.
  - PHP can't compute a year because its source `c_year` is 0.

So all 7 confirm Access fired Rule 05 correctly; PHP wrote NULL
in 6 cases because of a missing type-mapping row on its side, and
in 1 case because of a missing `c_year` value on its side.  Both
are upstream-data gaps, not algorithm divergence.

## Implication for PR Y's confidence

The sub-group splits cleanly into two PHP-side data-gap classes:

  - `candidate_php_entry_code_mapping_gap` × 6 — confidence
    promotable to **`supported_by_focused_probe`**.
  - `candidate_php_entry_data_year_missing` × 1 — newly named
    by this probe; same severity (PHP-side data gap), narrower
    failure mode.

Both leave Access vindicated on the rule-firing side.  Neither
is a confirmed CBDB bug per the maintainer-report convention
("PHP-side upstream data" is out-of-scope for this repo); the
finding is **maintainer-actionable** as a candidate for the
cbdb-online-main-server / SQLite-snapshot-build team rather than
the User MDB.

## Re-running

```
python analysis/probe_index_year_tcode05_entry_mapping.py
```

Pure pyodbc + sqlite3.  No Access COM.  Re-derive after any
fresh SQLite snapshot pull.
