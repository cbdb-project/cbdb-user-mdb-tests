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

### (Optional) Pull the cbdb-online-main-server SQLite snapshot

For the cross-check test (`tests/test_index_year_xcheck.py`,
roadmap item 12) you also need the upstream weekly SQLite dump:

```bash
python analysis/download_hf_sqlite.py
# downloads ~130 MB into data/cbdb_online_sqlite/, extracts ~550 MB
```

Re-run weekly when CBDB pushes a new dump.

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

### Coverage as of 2026-05-02

| Form | Real-VBA matrix | Real export | StoreID / RecallID | Import-list | Save-list | Notes |
|------|-----------------|-------------|---------------------|-------------|-----------|-------|
| LookAtEntry            | ✅ `test_vba_matrix.py` | ✅ CmdGIS (byte-diff) | ✅ Store; round-trip → Kinship ✅ | ✅ EntryCodes + Places | ✅ EntryCodes (3-col) | Bug #3 confirmed in this form only |
| LookAtStatus           | ✅ 3 fixtures | ✅ CmdGIS (struct) | ✅ Store | ✅ StatusCodes + Places | ✅ StatusCodes (3-col) | 17 023 + 4 931 rows |
| LookAtTexts            | ✅ biblcat 1 | ✅ CmdGIS (struct) | ✅ Store | ✅ TextCategories + Places | ✅ TextCategories (3-col) | 15 774 rows |
| LookAtAssociations     | ✅ 3 fixtures | ✅ CmdGIS (struct) | ✅ Store | ✅ Associations + Places | ✅ Associations (3-col) | 11 867 rows |
| LookAtOffice           | ✅ 2 fixtures | ✅ CmdGIS (struct) | ✅ Store | ✅ Offices + PlaceOffice + PlacePeople | ✅ Offices (3-col) | 37 429 + 35 748 rows |
| LookAtPlace            | ✅ 2 fixtures | ✅ CmdGIS (struct, GISFrame→CodeFrame workaround) | ✅ Store | ✅ Places | — (no save button) | 5 962 + 3 528 rows |
| LookAtKinship          | ✅ 1 fixture | ✅ CmdGIS (struct, requery shim) | ✅ Store + ✅ Recall | ✅ CmdImport | — (no save button) | 949 rows (Zhao Tingmei) |
| LookAtAssociationPairs | ⏭ skipped | — | ✅ Recall | ✅ CmdImportList | — (no save button) | `Link1stOrder` ASSOC_DATA self-join too slow |
| LookAtNetworks         | ⏭ skipped | — | ⏭ Recall (Form_Open hangs) | ⏭ ImportPeople / ImportPlaces (Form_Open) | — (no save button) | recursive expansion (Zhu Xi 2 471 assocs) |
| LookAtGroupData        | ⏭ skipped | — | ✅ Recall | ✅ CmdImport | — (no save button) | similar recursion |

**Latest matrix run**: `12 passed, 3 skipped in 110.22s`.
**Latest Store/Recall run** (`tests/test_vba_storeid_recallid.py`): `11 passed, 1 skipped in 142.82s`.
**Latest Import-list run** (`tests/test_vba_import_lists.py`): `15 passed, 2 skipped in 142.04s`.
**Latest Save-list run** (`tests/test_vba_save_lists.py`): `5 passed in 43.31s`.
**Latest Bilingual run** (`tests/test_vba_bilingual_ui.py`): `9 passed in 145.73s`.
**Latest CmdGIS-other-forms run** (`tests/test_vba_cmdgis_other_forms.py`): `4 passed, 2 skipped in 123.75s`.
**Combined regression** (all 8 test files, 71 tests): `59 passed, 8 skipped, 4 xfailed in 837.76s`.

### Confirmed bugs

1. **Bug #3** (`findings.md`) — `LookAtEntry.CmdQuery_Click` backfill
   `UPDATE` silently fails on the multi-table join (only this form;
   confirmed by cross-form matrix).
2. **Bug #1** — `View_StatusData` alias swap (`c_fy_range_*` columns
   show last-year values).
3. **Bug #2** — `dao360.dll` reference broken on Office 2016+ (one-time
   manual fix).
4. **Bug #4** (NEW, found 2026-05-02) —
   `Form_LookAtPlace.CmdGIS_Click` references a non-existent control
   `GISFrame` (the right control on this form is `CodeFrame`).
   Real users clicking the GIS button on LookAtPlace see "Object
   required" + no file written. Found by repeated `cmdgis_other_forms`
   test failures, surfaced after fixing the silent-Err-handler SQL
   bug. Test driver applies a `GISFrame.Value → CodeFrame.Value`
   rewrite as a workaround; the underlying CBDB code remains broken.

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
| 8 | 🟡 partial | First slice landed: `tests/test_vba_cmdgis_other_forms.py` extends real-`CmdGIS` coverage from LookAtEntry (`test_vba_export.py`, byte-diff) to Status / Texts / Associations / Office / Kinship (5 added). Uses **structural** assertions (file exists, header has `NameChn` column, ≥ 4 cols, ≥ 1 data row) rather than byte-diff — easier to maintain across data refreshes. **Driver fixes bundled in:** (1) chain block now injected right before `Exit Sub` (after cleanup) instead of right after `Exit_<name>:` — fixes Status/Texts/Associations whose subform rebind happens in cleanup; (2) `On Error Resume Next` injected at the start of the Exit handler so cleanup errors don't bounce into an infinite Resume loop (LookAtKinship's `OrderBy = "c_up,..."` observed); (3) per-form `_SUBFORMS_TO_REQUERY` map injects `<subform>.Form.Requery` into the chain block for forms whose CmdGIS reads a saved-query-bound subform (unblocked LookAtKinship). **Still skipped:** LookAtPlace's CmdGIS uses a different output pattern (text-stream `SaveToFile` directly, no binary CopyTo BOM-strip intermediate) and silently fails to materialise the file under our patched path. **Still open:** the other export buttons (`CmdNeo4j`, `CmdKML`, `CmdPajek`, `CmdGUESS`, `CmdGephi`, `CmdUCInet`, `CmdGISPeople`) and a fix for LookAtPlace's text-stream SaveToFile pattern |
| 9 | ✅ done | `tests/conftest.py::pytest_configure` re-runs `analysis/discover_test_inputs.py` automatically at session start when `analysis/dump/test_inputs.json` is missing or older than `data/CBDB_BJ_User.mdb`. Opt-out via `pytest --no-discover-inputs` |
| 10 | ⏳ open | Picker-dialog tests for `frmPickEntry_multi` etc. (currently bypassed by direct `ZZ_SCRATCH_*` writes) |
| 11 | ✅ done | Bilingual UI test (`tests/test_vba_bilingual_ui.py`) — for each of the 9 forms with the standard `CmdFanti` / `CmdJianti` toggle pair (Networks uses different names + Form_Open hangs), opens the form, drives one Fanti round-trip to force captions through `changeDisplayLanguage` (so the FormLabels-derived state, not design-time captions, is the baseline), then asserts each toggle changes ≥5 captions and the second toggle restores the baseline exactly. Latest: `9 passed in 145.73s`. **Caveat surfaced:** a few forms (LookAtPlace, LookAtGroupData) ship with design-time captions that don't match `FormLabels` (e.g. `LblFrom = "  From"` with leading spaces vs `FormLabels.c_fanti = "From"`); the design-time text is replaced on the first `changeDisplayLanguage` call. Pre-existing CBDB UX nit, documented in the test |
| 12 | ✅ done | `tests/test_index_year_xcheck.py` + `analysis/download_hf_sqlite.py` cross-check the User MDB's `BIOG_MAIN` against the weekly cbdb-online-main-server SQLite snapshot at <https://huggingface.co/datasets/cbdb/cbdb-sqlite/blob/main/latest.zip>. Compares 4 fields per `c_personid` — derived (`c_index_year`, `c_index_addr_id`) and source (`c_birthyear`, `c_deathyear`) — with 0.5 % / 0.1 % thresholds. **Algorithm-level agreement confirmed**: both pipelines run cbdb-online-main-server's `IndexYearRebuildService.php` Phase A/B/C; per-rule type-code distributions are nearly identical (137 vs 136 codes; per-rule counts differ by only 5-100 persons). The remaining ~575 person-level diffs (out of 657 246) are pure data-snapshot drift — the online system's source data is updated continuously and the User MDB lags. **Confirmed with maintainer: not bugs.** The test's role is to surface algorithm drift if it ever happens; per-pid data drift is expected. Full background: `findings.md` Note #1 + `AGENTS.md` "NOT a bug" callout |
| 13 | ✅ done | Import-list buttons (`tests/test_vba_import_lists.py`) — covers all 11 unique button names across 8 forms (15 of 17 tests pass; the 2 LookAtNetworks ones are skipped for the same Form_Open hang as items 7/15). Drives each via `Form_Timer`, points at a fixture file (whose delimiter / column count matches the saved `MSysIMEXSpecs`), and asserts: (a) the target `ZZ_SCRATCH_*` table contains exactly the valid IDs, (b) `InputErrorList` contains exactly the invalid IDs. The `gUse*` global side-effect is documented per spec but not asserted — an early inject-based reader caused JET re-entrancy hangs in matrix CmdQuery; the table-shape assertion is the meaningful contract. Driver gained: `patch_filedialog` now also handles the `With dlgX` block's `If .Show = -1 Then` (used in import subs) |
| 14 | ✅ done | Save-list buttons (`tests/test_vba_save_lists.py`) — covers all 5 `CmdSave*_Click` handlers. Pre-populates the source `ZZ_SCRATCH_<X>` table directly (skips CmdQuery), patches `FileDialog(msoFileDialogSaveAs)` via the existing `vba.patch_filedialog`, fires the button, and asserts the resulting tab-separated UTF-8 (BOM-stripped) file contains exactly the seeded IDs — and for the 3-column specs (Entry, Associations, Office, Status, TextCategories) also that the desc / desc_chn fields match an INNER JOIN against the codes table. Latest run: `5 passed in 43.31s`. **Important:** several Form_Open handlers wipe their picker scratch table on form load; the test seeds AFTER `open_form` to avoid this |
| 15 | ✅ done | `CmdStoreID` / `CmdRecallID` round-trip (`tests/test_vba_storeid_recallid.py`) — covers all 7 query-runnable forms for Store, 3 of 4 forms for Recall (Networks Form_Open hangs in this driver — same family as the matrix Networks skip), plus an end-to-end Entry → Kinship round-trip. Driver gained: `MsgBox "literal"` neutralizer in `_inject_autodetect`; chain+DONE block moved to *after* `Exit_<name>:` so it survives the `Resume Exit` from the form's Err handler |

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
- ✅ 已確認 4 個 bug（詳見 `findings.md`，最新一個是 LookAtPlace 的 CmdGIS 引用了不存在的控件 `GISFrame`）
- ⏳ 剩 3 個表單因遞迴展開太慢暫跳過
- ⏳ 其他匯出按鈕（Neo4j/KML/Pajek/GUESS/Gephi）尚未涵蓋
- ✅ pytest 啟動時自動偵測 `test_inputs.json` 是否過時並重跑 `discover_test_inputs.py`（`pytest --no-discover-inputs` 可關閉）— roadmap 第 9 項
- ✅ 比對 User MDB 的 `index year` / `index address` 算法與 [`cbdb-online-main-server`](https://github.com/cbdb-project/cbdb-online-main-server) — roadmap 第 12 項已成。两边跑同一个算法（PHP `IndexYearRebuildService.php`）；少量 person 级差异是线上系统持续更新源数据 vs MDB 用某个时点的快照所致，**与维护者确认这不是 bug**。测试有阈值看门，算法层一旦真飘走会立刻报警。详见 findings.md Note #1 / AGENTS.md "NOT a bug" 段
- ✅ Import-list 按鈕（`tests/test_vba_import_lists.py`，11 種按鈕跨 8 個表單，15 passed + 2 skipped；只有 LookAtNetworks 兩項因 Form_Open 卡住而 skip）— roadmap 第 13 項
- ⏳ Save-list 按鈕（`CmdSaveEntryCodes` / `CmdSaveOffices` 等）寫出清單檔的位元組級對比 — roadmap 第 14 項
- ✅ `CmdStoreID` / `CmdRecallID` 跨 form round-trip 測試（`tests/test_vba_storeid_recallid.py`，11 passed + 1 skipped；含 Entry → Kinship 完整 round-trip）— roadmap 第 15 項

## 貢獻

詳見 [`AGENTS.md`](./AGENTS.md) 中的 landmine 與 driver 模式說明。
**任何改動如果讓 [Plan & status](#plan--status) 中的項目進度改變，
請在同一個 PR 中同步更新該表格** —
這是 `AGENTS.md` 強制執行的規則。
