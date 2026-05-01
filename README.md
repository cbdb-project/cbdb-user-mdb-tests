# cbdb-user-mdb-tests

Automated regression tests for the **CBDB User MDB** (the Microsoft Access
front-end shipped by the [China Biographical
Database](https://projects.iq.harvard.edu/cbdb), file name
`CBDB_BJ_User.mdb` in the current edition). The test suite runs the
real VBA inside Access via COM automation, captures the resulting
scratch tables and exports, and compares them against frozen goldens —
catching regressions in queries, column backfills, and exports after
each `.mdb` data refresh.

> 中文版說明請見 [§ 中文簡介](#中文簡介-traditional-chinese)。

---

## What this is

CBDB ships a Windows-only Access database to non-technical historians.
The `.mdb` is updated periodically with new biographical data; each
update can silently break a query, drop a column from an export, or
shift a code lookup. There are 10 `LookAt*` query forms, hundreds of
saved queries, and ~50 000 lines of VBA — far too much to revalidate by
hand on every release.

This project instruments Access from Python, fires the real
`CmdQuery_Click` / `CmdRun_Click` / `CmdGIS_Click` event handlers (so
test results match what an actual user sees), and asserts:

- Result row counts vs. independent SQL replay
- Column-set equality (a dropped backfill is the #1 user-reported bug)
- Byte-level equality of GIS / Neo4j / KML exports
- Foreign-key integrity across joined lookup tables

Confirmed bugs found by the suite live in [`findings.md`](./findings.md)
([English](./findings_en.md)).

---

## Project structure

```
cbdb-user-mdb-tests/
├── data/                       # gitignored — drop CBDB MDB + HelpFiles here
│   ├── CBDB_BJ_User.mdb
│   ├── CBDB_<YYYYMMDD>_DATA.mdb
│   └── HelpFiles/
├── tests/
│   ├── cbdb_driver/            # Python ↔ Access COM driver
│   ├── cbdb_replay/            # independent SQL replay of each LookAt form
│   ├── golden/                 # frozen reference outputs
│   ├── test_vba_matrix.py
│   ├── test_vba_matrix_all_forms.py
│   ├── test_vba_export.py      # real CmdGIS export, byte-level diff
│   ├── test_vba_integrity.py   # 12-dimension data-integrity battery
│   ├── test_vba_differential.py
│   ├── test_known_bugs.py
│   ├── test_schema.py
│   ├── test_saved_views.py
│   ├── test_lookatentry.py
│   ├── test_other_lookat_forms.py
│   ├── test_exports.py
│   ├── DESIGN.md
│   ├── VALIDATION_DIMENSIONS.md
│   └── MANUAL_SMOKE.md
├── analysis/                   # one-off discovery + dump scripts
│   ├── dump_metadata.py        # writes analysis/dump/{tables,queries,…}.json
│   ├── dump_vba.py             # writes analysis/dump/vba/Form_*.vb
│   ├── discover_test_inputs.py # picks fresh fixtures from current data
│   └── dump/                   # frozen metadata snapshots (committed)
├── AGENTS.md                   # contributor / agent guide; landmines
├── findings.md                 # confirmed bugs + audit notes (中文)
├── findings_en.md              # same in English
└── FINAL_STATE.md              # snapshot of test-suite coverage
```

---

## Setup

### Requirements

| Tool | Version | Notes |
|------|---------|-------|
| Windows | 10 / 11 | Access COM is Windows-only |
| Microsoft Access | 2016+ (ACE engine) | Office 365 click-to-run works |
| Python | 3.11+ | tested on 3.12 |
| pyodbc, pywin32, pywinauto, pandas, pytest | latest | `pip install -r requirements.txt` |
| Microsoft Access Driver (`*.mdb, *.accdb`) | bundled with Office | required for ODBC access |

### Drop the MDB + HelpFiles into `data/`

The data files are not redistributed in this repo (proprietary CBDB
release). Obtain them from
[projects.iq.harvard.edu/cbdb/download-cbdb](https://projects.iq.harvard.edu/cbdb/download-cbdb)
and place them as:

```
data/CBDB_BJ_User.mdb            # the User-facing front-end
data/CBDB_<YYYYMMDD>_DATA.mdb    # the linked data backend
data/HelpFiles/                  # CBDB Users Guide + per-form HelpFile_*.pdf
```

### One-time setup of the User MDB

`CBDB_BJ_User.mdb` ships with a broken VBA reference to legacy DAO
(`dao360.dll`). On a new machine you must remove it once before tests
will run — see [`AGENTS.md` § DAO 3.6 reference](./AGENTS.md) for the
3-click fix in the VBE, or run:

```bash
python analysis/check_vba_refs.py
```

---

## Running the tests

```bash
# fast SQL-replay tests (no Access process; ~30s)
python -m pytest tests/ -W ignore --ignore=tests/test_infra_smoke.py \
  --ignore=tests/test_vba_matrix.py \
  --ignore=tests/test_vba_matrix_all_forms.py \
  --ignore=tests/test_vba_export.py

# real-VBA cross-form matrix (boots Access; ~2 min)
python -m pytest tests/test_vba_matrix_all_forms.py -v -W ignore

# real export byte-level diff (boots Access; ~10s)
python -m pytest tests/test_vba_export.py -v -W ignore -s
```

### After each CBDB `.mdb` refresh

```bash
python analysis/dump_metadata.py        # refreshes analysis/dump/*.json
python analysis/dump_vba.py             # refreshes analysis/dump/vba/
python analysis/discover_test_inputs.py # picks fresh fixtures
python -m pytest tests/ -W ignore       # full run
```

---

## Plan & status

This section is the source of truth for what's done, what's in flight,
and what's intentionally deferred. **Update it whenever a milestone
moves** — `AGENTS.md` enforces this as a contributor rule.

### Coverage as of 2026-05-01

| Form | Real-VBA matrix | Real export | Notes |
|------|-----------------|-------------|-------|
| LookAtEntry            | ✅ `test_vba_matrix.py` | ✅ CmdGIS | Bug #3 confirmed in this form only |
| LookAtStatus           | ✅ 3 fixtures | — | 17 023 + 4 931 rows |
| LookAtTexts            | ✅ biblcat 1 | — | 15 774 rows |
| LookAtAssociations     | ✅ 3 fixtures | — | 11 867 rows |
| LookAtOffice           | ✅ 2 fixtures | — | 37 429 + 35 748 rows |
| LookAtPlace            | ✅ 2 fixtures | — | 5 962 + 3 528 rows |
| LookAtKinship          | ✅ 1 fixture | — | 949 rows (Zhao Tingmei) |
| LookAtAssociationPairs | ⏭ skipped | — | `Link1stOrder` ASSOC_DATA self-join too slow |
| LookAtNetworks         | ⏭ skipped | — | recursive expansion (Zhu Xi 2 471 assocs) |
| LookAtGroupData        | ⏭ skipped | — | similar recursion |

**Latest matrix run**: `12 passed, 3 skipped in 110.22s`.

### Confirmed bugs

1. **Bug #3** (`findings.md`) — `LookAtEntry.CmdQuery_Click` backfill
   `UPDATE` silently fails on the multi-table join (only this form;
   confirmed by cross-form matrix).
2. **Bug #1** — `View_StatusData` alias swap (`c_fy_range_*` columns
   show last-year values).
3. **Bug #2** — `dao360.dll` reference broken on Office 2016+ (one-time
   manual fix).

### Roadmap

| # | Status | Item |
|---|--------|------|
| 1 | ✅ done | Build Access COM driver (`tests/cbdb_driver/vba_session.py`) — `Form_Timer` trigger pattern, `Me.Tag` chain pattern, `MsgBox Err.Description` rewrite for diagnostics |
| 2 | ✅ done | Cross-form matrix covering 7/10 LookAt forms (12 fixtures) |
| 3 | ✅ done | Real export byte-level test (`LookAtEntry.CmdGIS`) |
| 4 | ✅ done | 12-dimension data-integrity battery (`test_vba_integrity.py`) |
| 5 | ✅ done | Independent SQL replay layer (`tests/cbdb_replay/`) for differential checking |
| 6 | ✅ done | Auto-discovery of fresh fixtures (`analysis/discover_test_inputs.py`) |
| 7 | ⏳ open | Get `LookAtAssociationPairs` / `LookAtNetworks` / `LookAtGroupData` running — needs smaller-fixture search or VBA query rewrite (see `findings.md`) |
| 8 | ⏳ open | Real-export tests for the other buttons: `CmdNeo4j`, `CmdKML`, `CmdPajek`, `CmdGUESS`, `CmdGephi` (chain pattern already proven) |
| 9 | ⏳ open | Auto-run `discover_test_inputs.py` from `conftest.py` when `analysis/dump/test_inputs.json` is older than `data/CBDB_BJ_User.mdb` |
| 10 | ⏳ open | Picker-dialog tests for `frmPickEntry_multi` etc. (currently bypassed by direct `ZZ_SCRATCH_*` writes) |
| 11 | ⏳ open | Bilingual UI test for `changeDisplayLanguage` |
| 12 | ⏳ open | **Meta-tests** — see [§ Should the test project test itself?](#should-the-test-project-test-itself) |
| 13 | ⏳ open | Cross-check the `index year` and `index address` derivations in the User MDB against the equivalents produced by [`cbdb-online-main-server`](https://github.com/cbdb-project/cbdb-online-main-server), and assert per-person consistency between the two implementations |

### Should the test project test itself?

Short answer: yes, but lightly. The Access COM driver is fragile in
ways that don't show up until a real form run, and the real-VBA matrix
is too slow (~2 min) to use as a CI smoke test. Useful meta-tests:

- **Fixture-discovery smoke**: every `LookAt*` form in
  `form_specs.py` has at least one fixture in `_all_fixtures()`.
- **VBA-injection sanity**: `_inject_autodetect`'s regex round-trip
  doesn't break a control sample of `Form_*.vb` files.
- **Schema drift trip-wire**: `tests/golden/dump/schema_snapshot.json`
  vs. `analysis/dump/tables.json` — fail loud if column counts change.
- **Path / config sanity**: `data/CBDB_BJ_User.mdb` exists, the ACE
  driver is registered, `test_inputs.json` is newer than the MDB.

These run in <2 s and would catch the most common "I broke a fixture
loader" or "I forgot to refresh inputs" mistakes before the slow VBA
run. Tracked as item 12 in the roadmap.

---

## Architecture cheat-sheet

The driver pattern that makes this work is non-obvious; the full
landmine list lives in [`AGENTS.md`](./AGENTS.md). The minimum a
contributor needs to know:

1. `CmdQuery_Click` is `Private`, so `Application.Run "Form_X.CmdQuery_Click"`
   does **not** work. We trigger it via `Form_Timer`.
2. `Form_Timer` fires only **once** per `OpenForm`. To chain
   `CmdQuery → CmdGIS` we read `Me.Tag` and call the next sub from
   inside `CmdQuery_Click` itself (an autodetect-injected post-body
   block).
3. JET cache is incoherent across pyodbc and Access — write through
   pyodbc, then `DBEngine.Idle 8 + RefreshDatabaseWindow` before
   triggering VBA.
4. Use `DispatchEx` (not `Dispatch`) and skip `DoCmd.Close /
   CloseCurrentDatabase / Quit` — `taskkill` is the only reliable
   shutdown after a heavy `CmdQuery_Click`.
5. The autodetect injection rewrites `MsgBox Err.Description` into
   `INSERT INTO ZZ_TEST_DEBUG (msg) VALUES ('<form>:ERR ...')` so a
   hidden modal dialog can never silently block the COM thread again.

---

## Contributing

PRs welcome. Before submitting:

- `python -m pytest tests/ -W ignore` passes locally
- if you change anything that moves a roadmap item, **update
  the [Plan & status](#plan--status) table in this README** in the
  same PR (this is enforced by `AGENTS.md`)
- new bugs go in `findings.md` with a regression test
- new landmines (Access quirks, JET behavior, COM gotchas) go in
  `AGENTS.md`

---

## License

The Python test code in this repository is released under the MIT
License — see [`LICENSE`](./LICENSE) once added.

The CBDB Access database files (`CBDB_BJ_User.mdb`, `CBDB_*_DATA.mdb`,
the HelpFile PDFs) are **not** part of this repository; obtain them
under their own terms from
[projects.iq.harvard.edu/cbdb](https://projects.iq.harvard.edu/cbdb).

---

# 中文簡介 (Traditional Chinese)

本專案是 [中國歷代人物傳記資料庫 CBDB](https://projects.iq.harvard.edu/cbdb)
**用戶端 Access 資料庫**（即 `CBDB_BJ_User.mdb`，現行版本檔名）的
**自動化回歸測試套件**。每次 CBDB 釋出新資料時，由 Python 透過
COM 自動化觸發 Access 內真實的 VBA 事件（`CmdQuery_Click` /
`CmdRun_Click` / `CmdGIS_Click`），擷取 scratch tables 與匯出檔，
與冷凍 golden 檔比對，找出資料更新引入的 query / 欄位 / 匯出 regression。

## 為什麼

CBDB 用戶端是給歷史學家用的 Windows-only Access 介面，每次資料更新可能
默默壞掉一個 query、丟失一個欄位、或讓某個 export 變空。10 個 `LookAt*`
表單、上百個 saved query、約五萬行 VBA — 人工每次都重測不可能。

此套件在 Python 端啟動 Access、觸發真正的 VBA handler，斷言：

- 結果列數 vs. 獨立 SQL replay
- 欄位集合相等（用戶最常回報的「丟欄位」bug）
- GIS / Neo4j / KML 匯出檔位元組級相等
- 跨 lookup 表的外鍵完整性

已找到的 bug 在 [`findings.md`](./findings.md)。

## 安裝

1. Windows 10/11 + Microsoft Access 2016+ + Python 3.11+
2. `pip install pyodbc pywin32 pywinauto pandas pytest`
3. 從 [CBDB 官網](https://projects.iq.harvard.edu/cbdb/download-cbdb)
   取得 `CBDB_BJ_User.mdb`、`CBDB_<日期>_DATA.mdb`、`HelpFiles/`，
   放在 `data/` 之下（此資料夾已 gitignored，不會上傳）。
4. 第一次須在 Access VBE 中清掉斷掉的 `dao360.dll` 引用，或執行：
   ```bash
   python analysis/check_vba_refs.py
   ```

## 執行

```bash
# 快速 SQL replay 測試（不啟 Access；~30 秒）
python -m pytest tests/ -W ignore \
  --ignore=tests/test_vba_matrix.py \
  --ignore=tests/test_vba_matrix_all_forms.py \
  --ignore=tests/test_vba_export.py

# 真 VBA 跨表單 matrix（會啟 Access；~2 分鐘）
python -m pytest tests/test_vba_matrix_all_forms.py -v -W ignore

# 真匯出位元組級對比（~10 秒）
python -m pytest tests/test_vba_export.py -v -W ignore -s
```

## 專案計畫與當前狀態

完整的計畫與覆蓋狀態在英文版的
[§ Plan & status](#plan--status) 表格中（單一真相來源，請勿在中文版重複維護以免不一致）。

簡述：
- ✅ 7/10 LookAt 表單已納入真 VBA matrix（12 fixtures，110 秒跑完）
- ✅ 1 個真實匯出位元組對比（`LookAtEntry.CmdGIS`）
- ✅ 12 維度資料完整性檢查
- ✅ 已確認 3 個 bug（詳見 `findings.md`）
- ⏳ 剩 3 個表單因遞迴展開太慢暫跳過
- ⏳ 其他匯出按鈕（Neo4j/KML/Pajek/GUESS/Gephi）尚未涵蓋
- ⏳ 「測試專案自己的測試」（meta-tests）— 規劃在 roadmap 第 12 項
- ⏳ 比對 User MDB 的 `index year` / `index address` 算法與 [`cbdb-online-main-server`](https://github.com/cbdb-project/cbdb-online-main-server) 所產生結果的一致性 — roadmap 第 13 項

## 貢獻

詳見 [`AGENTS.md`](./AGENTS.md) 中的 landmine 與 driver 模式說明。
**任何改動如果讓 [Plan & status](#plan--status) 中的項目進度改變，
請在同一個 PR 中同步更新該表格** —
這是 `AGENTS.md` 強制執行的規則。
