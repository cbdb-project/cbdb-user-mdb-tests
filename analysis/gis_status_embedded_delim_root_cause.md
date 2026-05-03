# LookAtStatus GIS export — embedded-delimiter root-cause trace

PR T's `--include-vba` run flagged `LookAtStatus` GIS export row
11476 as having 10 tab-separated cells against a 9-col header.  This
note traces the full chain from source storage to exported file and
explains exactly which step introduces the offending TAB character.

## TL;DR

ADDR_CODES has 315 rows whose `c_name` and `c_name_chn` start with a
stray BOM (U+FEFF) prefix.  When LookAtStatus's CmdQuery copies one
of these rows into the scratch staging table via SQL UPDATE/INSERT,
**JET strips the BOM and re-interprets the remaining UTF-16 bytes as
single-byte code-page-1252 chars**, producing a corrupted Unicode
string that contains a literal TAB character.  CmdGIS then writes
this without escaping — the TAB becomes a delimiter, splitting one
column into two and silently shifting every column to its right.

Two independent candidate fixes:

1. **One-shot data cleanup** — strip the BOM prefix from the 315
   ADDR_CODES rows.  Fixes existing exports immediately; nothing
   protects against the same data sneaking back in.
2. **Defensive sanitisation in CmdGIS** — replace Chr(9) / Chr(10) /
   Chr(13) / U+FEFF / control chars with space before appending each
   cell value.  Architectural fix; protects 6 forms (Status, Texts,
   Place, Associations, Office, Kinship).

Neither yet labelled a confirmed bug.  Severity rationale: a real
user clicking GIS on any LookAt form whose query touches one of the
315 dirty addresses gets a `.tab` file with column misalignment —
numeric coordinates land in the AddrChn column.  This is exactly the
silent-export-corruption pattern the depth check was meant to catch.

## The chain, byte-for-byte

Subject row: `BIOG_MAIN.c_personid = 29619` (Ruan Fu / 阮孚),
`ADDR_CODES.c_addr_id = 702559` (intended: Wei Shi / 尉氏 = the
Wei-clan seat).

### Step 1 — Source: ADDR_CODES.c_name_chn

| Field         | Value as Python repr | UTF-16 LE bytes        |
| ------------- | -------------------- | ---------------------- |
| `c_name`      | `'﻿Wei Shi'`    | `ff fe 57 00 65 00 69 00 20 00 53 00 68 00 69 00` |
| `c_name_chn`  | `'﻿尉氏'`       | `ff fe 09 5c 0f 6c`    |

Decoded as UTF-16 LE the `c_name_chn` is exactly 3 chars:
U+FEFF (BOM) + U+5C09 (尉) + U+6C0F (氏).

### Step 2 — JET SQL UPDATE/INSERT into ZZ_SCRATCH_P_STATUS

`Form_LookAtStatus.vb:1398` issues an UPDATE that copies
`[ADDR_CODES].[c_name_chn]` into `ZZ_SCRATCH_STATUS.c_addr_chn`,
then `Form_LookAtStatus.vb:1409` does an INSERT INTO
`ZZ_SCRATCH_P_STATUS (..., c_addr_chn, ...)`.

After CmdQuery has run, querying the staging copy mdb shows:

| Field          | Value as Python repr | UTF-16 LE bytes         |
| -------------- | -------------------- | ----------------------- |
| `c_addr_name`  | `'﻿Wei Shi'`    | `ff fe 57 00 65 00 69 00 20 00 53 00 68 00 69 00` |
| `c_addr_chn`   | `'\t\\\x0fl'`        | `09 00 5c 00 0f 00 6c 00` |

`c_addr_name` survived intact (the BOM is preserved as a real
U+FEFF char and the ASCII letters round-trip cleanly).
**`c_addr_chn` was mangled.**  The original 3-char Unicode string
became a 4-char Unicode string with a literal TAB at position 0.

The mangling is mechanical: drop the BOM (`ff fe`), then re-interpret
the remaining 4 bytes `09 5c 0f 6c` as four single-byte chars
(TAB, `\`, SI, `l`), then promote each to a UTF-16 code unit by
zero-extending to 2 bytes.  Net result: `09 00 5c 00 0f 00 6c 00`.

This matches the consistent behaviour across all 315 dirty rows.

### Step 3 — CmdGIS writes the line, no escaping

`Form_LookAtStatus.vb:1554-1636` builds each output line as
`tStr + (column value) + tC` where `tC = Chr(9)` (line 1552).  No
sanitisation is performed on cell values.

Reading the produced `.tab` file at row 11476:

| Cell # | Header column | Actual repr           |
| ------ | ------------- | --------------------- |
| 0      | Name          | `'Ruan Fu'`           |
| 1      | NameChn       | `'阮孚'`              |
| 2      | Sex           | `'M'`                 |
| 3      | IndexYear     | `' 283'`              |
| 4      | AddrName      | `'﻿Wei Shi'`     |
| 5      | AddrChn       | `''` ← split happened here |
| 6      | X             | `'\\\x0fl'` ← what was *meant* to be in AddrChn |
| 7      | Y             | `' 114.1167357'` ← the X value, shifted left |
| 8      | xy_count      | `' 34.40078405'` ← the Y value, shifted left |
| 9      | (overflow)    | `' 1'`                |

The TAB at the start of the mangled `c_addr_chn` is consumed by
the delimiter parser as the AddrName→AddrChn separator, so the
remaining `\\\x0fl` lands in cell 6 (the X column) instead of cell
5 (AddrChn).  The user opening this file in Excel sees an empty
AddrChn, garbage in X, real X in Y, real Y in xy_count, and the
real xy_count spilled into a 10th column.

## Reproducer

```
# probe source: 315 ADDR_CODES rows with BOM in c_name and c_name_chn
python analysis/probe_status_gis_embedded_delim.py
#   wrote reports/gis_embedded_delimiter_findings.json

# probe export: actually run CmdGIS and dump every bad row's bytes
python analysis/probe_status_gis_export_bytes.py
#   wrote reports/gis_status_export_bytes_dump.json
#   wrote analysis/_status_gis_dump.tab (the actual export)
```

The export-bytes probe needs `--include-vba`-equivalent permissions
(it spawns Access via COM); the source probe is pure pyodbc and
runs in <2s.

## Why the test depth check was right to fail

The PR P depth check (`tests/test_vba_cmdgis_other_forms.py`,
`_assert_gis_export_depth`, check 8b) demands per-row cell count ==
header column count.  This is exactly the assertion a user-facing
silent-corruption bug should trip.  Weakening this check (capping
at first-N rows, lenient delta, etc.) would lose its primary
value — catching column shift bugs that produce parseable but
semantically wrong files.

PR U does **not** modify this check.  The LookAtStatus GIS test
continues to fail under `--include-vba` until either the data is
cleaned up or CmdGIS is hardened.  That's the correct posture.

## Reach analysis

| Form          | Touches ADDR_CODES through                   | Likely to surface dirty rows |
| ------------- | -------------------------------------------- | --- |
| LookAtStatus  | STATUS_DATA → BIOG_ADDR_DATA → ADDR_CODES JOIN | Yes — confirmed |
| LookAtTexts   | TEXT_BIBLDATA → BIOG_ADDR_DATA → ADDR_CODES  | Yes (untested here) |
| LookAtPlace   | reads ADDR_CODES directly via picker         | Yes (untested here) |
| LookAtKinship | KIN_DATA → BIOG_ADDR_DATA → ADDR_CODES       | Yes (untested here) |
| LookAtOffice  | POSTED_TO_ADDR_DATA / BIOG_ADDR_DATA         | Yes (untested here) |
| LookAtAssoc.  | ASSOC_DATA → BIOG_ADDR_DATA → ADDR_CODES     | Yes (untested here) |

For status_code = 40 (the PR T fixture), exactly 1 of 315 dirty
addresses is reachable.  Different fixtures will surface different
counts — running the full matrix against each form would give a
ground-truth count of how many user exports today produce
misaligned rows.  Out of scope for PR U.

## Conservative classification

Two new candidate findings, neither labelled confirmed bug:

- `candidate_silent_export_column_misalignment_BOM_addr_codes`
  (P0-class severity if confirmed: silent column shift in
  user-facing GIS exports across 6 forms).
- `candidate_jet_BOM_unicode_mangling_on_update_insert`
  (architectural; explains *why* the source-data BOM produces a
  TAB downstream).

Adding either to `reports/CBDB_Issues_Report` is left as a
separate decision — the evidence is concrete, but the canonical
issue list is curated and crossing into it is a scoped editorial
call the user should make.
