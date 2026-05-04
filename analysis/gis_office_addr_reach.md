# GIS reach extension — Office side + Kinship recursive (PR AF)

Addresses two caveats from PR W and incorporates the OFFICE_CODES finding from PR AE.

## Office-side reach

- OFFICE_CODES rows with BOM in c_office_chn: **5** (per PR AE)
- Of those, reachable via POSTED_TO_OFFICE_DATA (i.e. someone is actually posted to that office): **0** / 5

| c_office_id | c_office_chn (repr) | persons posted |
|---:|---|---:|
| 202976 | '\ufeff某京\ufeff鹽鐵使' | 0 |
| 202977 | '\ufeff某京戶部使事' | 0 |
| 202978 | '\ufeff某京度支副使' | 0 |
| 202979 | '\ufeff某京三司使事' | 0 |
| 202980 | '\ufeff某京轉運判官' | 0 |

### Export-path exposure

`LookAtOffice.CmdGIS_Click` does NOT emit `c_office_chn` directly — it pulls AddrName/AddrChn from the ADDR_CODES join via ZZ_SCRATCH_P_OFFICE.  OFFICE_CODES BOM rows therefore do NOT leak through the .tab GIS export.

`LookAtOffice.CmdNeo4j_Click` DOES emit `c_office_chn` into the OfficeCodes.csv file, TAB-separated and **unescaped** (Form_LookAtOffice.vb:1360, mirror at :2676).  Same architectural gap as Issue #20.  The 5 BOM rows would mangle to TAB via JET on the scratch-table staging step that precedes the write, splitting OfficeCodes.csv columns on those rows.

Severity bound: lower than Issue #20 ADDR side (5 vs 315 dirty rows), and CmdNeo4j is less commonly used than CmdGIS, but the bug class is identical.

## Kinship recursive reach

BFS from c_personid=29619 (Ruan Fu / 阮孚 — the only person whose record references a dirty ADDR_CODES row, c_addr_id=702559 / Wei Shi 尉氏) over the undirected KIN_DATA graph, up to 4 hops.

| Hop | Persons added | Cumulative |
|---:|---:|---:|
| 0 | 1 | 1 |
| 1 | 3 | 4 |
| 2 | 3 | 7 |
| 3 | 2 | 9 |
| 4 | 4 | 13 |

Any of the **12** non-seed persons in this set would, if picked in a LookAtKinship CmdRun + CmdGIS run, expose the Ruan-Fu→Wei-Shi dirty row in their kinship export (via the ZZ_SCRATCH_KIN staging that pulls Ruan Fu's BIOG_ADDR_DATA).

PR W's direct-only kin model said 3.  Recursive at 4 hops gives a much larger upper bound — but most of those pickers are unlikely to be picked in practice (distant kin), so the user-visible incidence remains negligible.

## Update to Issue #20 reach

PR W's headline was: of 315 dirty ADDR_CODES rows, only 1 reaches user exports.  PR AF adds: ALSO 5 dirty OFFICE_CODES rows reach LookAtOffice.CmdNeo4j exports via the OfficeCodes.csv file.  None labelled confirmed CBDB bug.

## Re-running

```
python analysis/audit_export_delimiter_risk.py  # PR AE
python analysis/analyze_gis_office_addr_reach.py  # this PR
```

Pure pyodbc + sqlite3.  No Access COM.