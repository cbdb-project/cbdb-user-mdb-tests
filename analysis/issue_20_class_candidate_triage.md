# Issue #20-class candidate triage (PR AP)

PR AE's static delimiter-risk audit and PR AF's reach probe
surfaced 4 candidate findings parallel to Issue #20.  Per
the morning brief: **don't add to ISSUES list directly; do a
focused review note first**.  This note documents each
candidate's user-reachability today and proposes whether to
promote, fold-into-Issue-#20, or leave latent.

## Candidate 1 — `OFFICE_CODES.c_office_chn` U+FEFF × 5

**Source**: PR AE (`reports/export_delimiter_risk_audit.json`).
**Reach probe**: PR AF
(`reports/gis_office_addr_reach.json`).

| Question | Answer |
| --- | --- |
| Same bug class as Issue #20? | Yes — BOM in Chinese-name column, JET would mangle to TAB on staging-table copy (per Form_LookAtStatus.vb:1398-1409 staging pattern; LookAtOffice's Form_LookAtOffice.vb is the parallel) |
| Reachable via LookAtOffice CmdGIS .tab? | **No** — CmdGIS reads ZZ_SCRATCH_P_OFFICE which joins ADDR_CODES (NOT OFFICE_CODES) for AddrName/AddrChn |
| Reachable via LookAtOffice CmdNeo4j OfficeCodes.csv? | **Yes** — Form_LookAtOffice.vb:1360 and :2676 emit `!c_office_chn` TAB-separated without escaping |
| User-visible incidence today? | **0** — all 5 dirty offices have 0 persons posted to them in POSTED_TO_OFFICE_DATA |
| Latent risk if any office gains a posting? | Yes — would reproduce the column-misalignment in OfficeCodes.csv |

**Verdict**: Latent data quality issue, identical bug class to
Issue #20.  Issue #20's two candidate fixes (data cleanup +
CmdGIS / CmdNeo4j escaping) cover this without modification.

**Recommendation**: **Fold into Issue #20's "Known reach" /
"Suggested fix" sections** as part of the existing maintainer
review (don't open Issue #20.b).  When the maintainer adds
the `UPDATE ADDR_CODES SET c_name = Mid(c_name, 2) WHERE
Left(c_name, 1) = ChrW(65279)` cleanup statement, the same
statement against OFFICE_CODES.c_office_chn will close this
finding.

## Candidate 2 — `BIOG_MAIN.c_notes` U+000A × 193 (LF)

**Source**: PR AE.
**Reach probe**: this PR (PR AP).

| Question | Answer |
| --- | --- |
| Same bug class as Issue #20? | Same family (delimiter-risk char in source-table column) but the export writer would have to *emit* the column for the bug to surface |
| Does any LookAt* CmdGIS write `c_notes`? | **No** — grep `!c_notes` across `analysis/dump/vba/Form_LookAt*.vb` returns **0 matches** |
| Does any LookAt* CmdNeo4j write `c_notes`? | **No** — same |
| Is `c_notes` INSERTed into scratch tables? | Yes — Form_LookAtKinship.vb (lines 932-1454), Form_LookAtGroupData.vb (lines 2313-2832), Form_LookAtOffice.vb (lines 2063-2069).  But it stays in the scratch table; no export READS it back. |
| User-visible incidence today? | **0** — column is unreachable via any documented export path |

**Verdict**: Latent data quality issue with **no current
export reachability**.

**Recommendation**: **Do not promote**.  Document in
`analysis/export_delimiter_risk_audit.md`'s "Notes" section as
a known latent c_notes finding.  If a future export feature
(e.g. a person-detail-CSV export) starts reading c_notes,
re-evaluate.

## Candidate 3 — `BIOG_MAIN.c_notes` U+000D × 193 (CR)

Same as Candidate 2.  CR usually pairs with LF (CRLF
line-endings), so it's the same 193 rows.  Same recommendation:
**do not promote**.

## Candidate 4 — `BIOG_MAIN.c_notes` U+0009 × 1 (TAB)

**Source**: PR AE.

| Question | Answer |
| --- | --- |
| Same bug class as Issue #20? | Same family — embedded TAB |
| Does any export emit `c_notes`? | No (per Candidate 2 reach analysis) |
| User-visible incidence today? | **0** |
| One row only — worth a per-row reach probe? | Marginal — single row, very low expected incidence even if c_notes ever became export-bound |

**Verdict**: Same as Candidates 2/3.  **Do not promote**.

## Summary table

| Candidate | Latent? | Reachable today? | Recommendation |
| --- | --- | --- | --- |
| 1. OFFICE_CODES BOM × 5 | yes | yes (CmdNeo4j path), but 0 affected rows today | fold into Issue #20 |
| 2. BIOG_MAIN c_notes LF × 193 | yes | **no** | do not promote |
| 3. BIOG_MAIN c_notes CR × 193 | yes | **no** | do not promote |
| 4. BIOG_MAIN c_notes TAB × 1 | yes | **no** | do not promote |

## Constraints respected

- No edits to `reports/generate_report.py`'s ISSUES list.
- No new ISSUES entries.
- No severity changes.
- All evidence is static (pure-pyodbc / file-grep); no
  Access COM.

## Re-running

```
python analysis/audit_export_delimiter_risk.py     # PR AE
python analysis/analyze_gis_office_addr_reach.py    # PR AF
# This note is hand-written; nothing to re-run for it.
```
