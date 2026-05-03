# Issue #20 — reach analysis

PR U + V documented that 315 `ADDR_CODES` rows carry a stray
U+FEFF prefix in `c_name` and `c_name_chn`, and that JET's
UPDATE/INSERT chain mangles BOM-prefixed values into strings
containing a literal TAB at position 0 — which then breaks
CmdGIS's tab-separated output without escaping.  This note
quantifies the actual user-visible reach of that bug across the
6 GIS-capable forms.

Headline number is much smaller than 315.  Almost all 315 dirty
ADDR_CODES rows are **orphans** — no person record references
them through `BIOG_MAIN.c_index_addr_id` or `BIOG_ADDR_DATA`.
The single non-orphan is `c_addr_id = 702559` (Wei Shi / 尉氏),
attached to `c_personid = 29619` (Ruan Fu / 阮孚) — exactly the
row PR V's byte trace identified.

Companion JSON: `reports/gis_embedded_delimiter_reach.json`.

## Universal counts (from `analyze_gis_embedded_delim_reach.py`)

| Metric | Count of 315 dirty addrs |
| ------ | -----------------------: |
| Used as anyone's `BIOG_MAIN.c_index_addr_id` | **1** |
| Anywhere in `BIOG_ADDR_DATA` | **1** |

So 314 of the 315 are dormant data-quality issues at the table
level — they exist in `ADDR_CODES` but nothing in the rest of
the schema points at them.  A user can't make them surface
through any person-anchored CmdQuery.

## Per-form reach

| Form | Source chain | Upper-bound dirty reachable | Current fixture reachable | Confidence |
| ---- | ------------ | ---: | ---: | ---------- |
| LookAtStatus | `STATUS_DATA → BIOG_ADDR_DATA → ADDR_CODES` | 1 | 1 (`c_status_code=40`); 0 (`=114`) | **byte_confirmed** (PR V evidence) |
| LookAtTexts | `BIOG_TEXT_DATA → BIOG_MAIN.c_index_addr_id → ADDR_CODES` | 0 | 0 (`c_text_cat_code=1`) | not_reached_by_source_chain |
| LookAtPlace | picker IS the address | **315** (any could be picked) | 0 (`7213`); 0 (`7686`) | picker_addressed |
| LookAtOffice | `POSTED_TO_OFFICE_DATA → BIOG_ADDR_DATA → ADDR_CODES` | 0 | 0 (`80944`); 0 (`87473`) | not_reached_by_source_chain |
| LookAtAssociations | `ASSOC_DATA → BIOG_ADDR_DATA → ADDR_CODES` | 0 | 0 (`437`); 0 (`438`) | not_reached_by_source_chain |
| LookAtKinship | `KIN_DATA(of kin) → BIOG_ADDR_DATA → ADDR_CODES` | 1 | 0 (`c_personid=3211`) | likely_reachable |

Notes:
- **LookAtPlace** is the special case: its picker IS an address.
  Of the 315 dirty addresses, only 1 has at least one
  `BIOG_ADDR_DATA` person attached, so picking any of the other
  314 would give an empty GIS export anyway.  Picking
  `c_addr_id = 702559` (Wei Shi 尉氏) would show Ruan Fu and
  reproduce the column misalignment — same row content as
  PR V's byte-confirmed LookAtStatus row.
- **LookAtKinship** can surface the same dirty row if any of
  the **3 persons who have Ruan Fu as kin** (`SELECT
  c_personid FROM KIN_DATA WHERE c_kin_id = 29619` returns 3)
  is picked.  The repo's current fixture `c_personid = 3211`
  is not one of them.
- Ruan Fu has 0 entries in `BIOG_TEXT_DATA`, `ASSOC_DATA`, and
  `POSTED_TO_OFFICE_DATA`, so LookAtTexts / LookAtAssociations /
  LookAtOffice cannot surface this bug from the current source
  data — even with arbitrary picker choices.

## Practical user impact

Conservatively: **as of CBDB_20260430_DATA.mdb the user-facing
manifestation of Issue #20 is one row** — the same row PR V's
byte trace identified.  It surfaces in:

- LookAtStatus exports whose query returns Ruan Fu (`c_status_
  code=40` does; other status codes may also).
- LookAtKinship exports anchored on one of the 3 persons whose
  kin includes Ruan Fu.
- LookAtPlace exports if a user picks `c_addr_id = 702559`.

The 314 orphan dirty addresses are a **latent data quality
issue**: they don't break any current export, but they would
reproduce the same misalignment the moment any of them gains
its first person link.  The recommended fixes from Issue #20
remain unchanged:

1. One-shot ADDR_CODES BOM cleanup — clears all 315 latent rows.
2. CmdGIS sanitisation — protects against any future similar
   dirt regardless of where it comes from.

## Caveats

- Reach numbers reflect a **single dirty address**
  (`c_addr_id = 702559`).  If upstream cleans this one row
  without fixing the export-writer escaping, the headline
  problem disappears today but the architectural gap remains.
- `LookAtKinship`'s SQL model walks direct kin only; the form's
  CmdQuery walks `KIN_DATA` recursively.  Recursive reach
  could be larger by a small factor; for this single row, the
  delta is bounded by the size of Ruan Fu's extended kin
  network, not by anything that would change the headline.
- `LookAtOffice` joins `POSTED_TO_OFFICE_DATA` to person address
  data; the user-visible export also writes the OFFICE address
  (header column `OfficeAddr`).  Office addresses are stored
  in a different ADDR_CODES join that's not separately probed
  here — the 0-reach for LookAtOffice covers the person
  address only.  An OfficeAddr-side reach probe is left as
  follow-up if anyone wants to harden this further.

## Re-running

```
python analysis/probe_status_gis_embedded_delim.py     # source scan
python analysis/analyze_gis_embedded_delim_reach.py    # this report
```

Both are pure pyodbc; no Access COM required.  Outputs land in
`reports/gis_embedded_delimiter_findings.json` and
`reports/gis_embedded_delimiter_reach.json`.
