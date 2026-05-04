# Export delimiter risk audit (PR AE)

Static scan of export-bound text columns in the User MDB for characters that can break tab-/line-/comma-separated exports.  Extends Issue #20.

## Headline

- Tables scanned: 8
- Distinct (table, column, char) findings: **6**
- New findings beyond Issue #20: **4**

## Findings ranked by row count

| Table | Column | Char | Rows | Known Issue #20 |
|---|---|---|---:|---|
| `ADDR_CODES` | `c_name` | `U+FEFF` | 315 | yes |
| `ADDR_CODES` | `c_name_chn` | `U+FEFF` | 315 | yes |
| `BIOG_MAIN` | `c_notes` | `U+000A` | 193 | **no — new candidate** |
| `BIOG_MAIN` | `c_notes` | `U+000D` | 193 | **no — new candidate** |
| `OFFICE_CODES` | `c_office_chn` | `U+FEFF` | 5 | **no — new candidate** |
| `BIOG_MAIN` | `c_notes` | `U+0009` | 1 | **no — new candidate** |

## Per-table detail

### `ADDR_CODES` — 30100 rows scanned

**Problem chars:**
- `c_name` — U+FEFF×315
- `c_name_chn` — U+FEFF×315

**Comma occurrences (only relevant for comma-separated / Pajek-family exports):**
- `c_notes` — 2 rows contain `','`

**Sample values:**
- `c_name` / U+FEFF (BOM (U+FEFF) — JET mangles to TAB on UPDATE/INSERT (Issue #20)): '\ufeffPu Yang'
- `c_name_chn` / U+FEFF (BOM (U+FEFF) — JET mangles to TAB on UPDATE/INSERT (Issue #20)): '\ufeff濮陽'

### `BIOG_MAIN` — 657785 rows scanned

**Problem chars:**
- `c_notes` — U+0009×1, U+000A×193, U+000D×193

**Comma occurrences (only relevant for comma-separated / Pajek-family exports):**
- `c_notes` — 34307 rows contain `','`
- `c_fl_ly_notes` — 10 rows contain `','`
- `c_fl_ey_notes` — 11 rows contain `','`
- `c_name_proper` — 28 rows contain `','`
- `c_name` — 688 rows contain `','`

**Sample values:**
- `c_notes` / U+000A (LF (U+000A) — splits line-based exports): "Han(1) Qi(2) [631] Guohua's [651] great grandson, Ju(2)'s [642] grandson, Zhiyan's son [3321], Chao(2) Zhonghui's [7002] and Fan(2) Chunli's [560] son-in-law. 
- `c_notes` / U+000D (CR (U+000D) — splits line-based exports): "Han(1) Qi(2) [631] Guohua's [651] great grandson, Ju(2)'s [642] grandson, Zhiyan's son [3321], Chao(2) Zhonghui's [7002] and Fan(2) Chunli's [560] son-in-law. 
- `c_notes` / U+0009 (TAB (U+0009) — splits tab-separated GIS .tab rows): '2024.01.19 Notes (on the name Yan Guangyuan or Yan Guangdu):\r\nThe epitaph Gao Shi (Wife of Zhang Di) (575374) said that Yan Xiang (575372) had a son named Gu

### `ENTRY_CODES` — 272 rows scanned


**Comma occurrences (only relevant for comma-separated / Pajek-family exports):**
- `c_entry_desc` — 7 rows contain `','`
- `c_entry_desc_chn` — 4 rows contain `','`

### `ASSOC_CODES` — 498 rows scanned


**Comma occurrences (only relevant for comma-separated / Pajek-family exports):**
- `c_assoc_desc` — 10 rows contain `','`
- `c_assoc_desc_chn` — 3 rows contain `','`

### `STATUS_CODES` — 284 rows scanned


**Comma occurrences (only relevant for comma-separated / Pajek-family exports):**
- `c_status_desc` — 5 rows contain `','`

### `TEXT_CODES` — 60934 rows scanned


**Comma occurrences (only relevant for comma-separated / Pajek-family exports):**
- `c_title_chn` — 201 rows contain `','`
- `c_notes` — 186 rows contain `','`
- `c_title_alt_chn` — 1 rows contain `','`

### `TEXT_BIBLCAT_CODES` — 144 rows scanned

(no findings)

### `OFFICE_CODES` — 34049 rows scanned

**Problem chars:**
- `c_office_chn` — U+FEFF×5

**Comma occurrences (only relevant for comma-separated / Pajek-family exports):**
- `c_notes` — 50 rows contain `','`
- `c_office_pinyin` — 21 rows contain `','`

**Sample values:**
- `c_office_chn` / U+FEFF (BOM (U+FEFF) — JET mangles to TAB on UPDATE/INSERT (Issue #20)): '\ufeff某京\ufeff鹽鐵使'

## Notes

- The 315 ADDR_CODES BOM-prefixed rows from Issue #20 are flagged here for completeness; they are already documented in the maintainer report.
- Comma findings are tracked but unranked because their severity depends on which export uses comma-quoting.  Pajek node labels and GUESS exports do quote, so commas in `c_name_chn` etc. should be safe.  GIS .tab and Neo4j CSVs use tab/comma + quote depending on writer.
- This audit does NOT examine `BIOG_ADDR_DATA`, `POSTED_TO_OFFICE_DATA`, etc.  Those are fact tables with foreign-key references to the code tables here; they inherit the same risk via JOIN.  Add separately if there's value.
- No bucket is labelled a confirmed CBDB bug.