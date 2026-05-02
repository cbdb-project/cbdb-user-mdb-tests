# `CBDB_BJ_User.mdb` 逻辑审查报告

审查时间：2026-04-30  
审查范围：`CBDB_BJ_User.mdb`（41 MB） — 79 个表单、21 条保存的查询、66 个表单 VBA 模块、2 个类模块、共约 5 万行 VBA 代码。

## 已确认的 bug

### Bug #3 — `LookAtEntry.CmdQuery_Click` 的 backfill UPDATE 不工作（**用户报告的"丢栏位"类 bug**）

`Form_LookAtEntry.vb:1778-1789` 的 UPDATE 语句负责把 c_entry_desc / c_addr_name / c_kin_name 等 descriptive 列从 lookup 表 backfill 到 ZZ_SCRATCH_ENTRY。结果集大到 ~30k+ 行时，UPDATE 静默失败，c_entry_desc 等列保持 NULL。

**重现（自动化）**：
```bash
python -m pytest tests/test_vba_matrix.py::test_vba_full_matrix[top_entry_code_36_unfiltered]
# AssertionError: c_entry_desc backfill wrong: [(36, None, 'examination: jinshi (general)'), ...]
```

**症状**: 用户在 LookAtEntry 跑大查询（大量 jinshi、不限朝代）→ 结果表显示 c_entry_desc / c_addr_name 为空（"列丢失"）。Kaifeng addr 100658 + 900-1100 这种 ~100 行的小结果就没问题。

**关键确认（2026-05-01 cross-form matrix）**: 此 bug **仅 LookAtEntry 有**！同样规模的查询：
- LookAtStatus 17,023 行 → backfill 全正常 ✓
- LookAtTexts 15,774 行 → backfill 全正常 ✓
- LookAtAssociations 11,867 行 → backfill 全正常 ✓

所以原假设「大结果集触发」错误。真正原因是 LookAtEntry 的 UPDATE 结构特有：
- LookAtEntry UPDATE 涉及 7+ 张表的级联 JOIN（`ZZ_SCRATCH_ENTRY LEFT JOIN INDEXYEAR_TYPE_CODES INNER JOIN ENTRY_CODES LEFT JOIN ASSOC_CODES LEFT JOIN BIOG_MAIN LEFT JOIN BIOG_MAIN_1 LEFT JOIN KINSHIP_CODES LEFT JOIN SOCIAL_INSTITUTION_NAME_CODES`）
- 其他 form 的 UPDATE 都更简单（2-3 张表）
- Access JET 的 multi-table JOIN UPDATE 在 join 图过复杂时 silently fail

**修复方向**: 把 LookAtEntry 的两个大 UPDATE 拆成多个小 UPDATE（每个只 join 1-2 张 lookup 表）。

**回归测试**: `tests/test_vba_matrix.py::test_vba_full_matrix[top_entry_code_36_unfiltered]`

---

### LookAtOffice — CmdQuery 按钮在自动化测试里点不动（infra landmine, 不是 bug） — **RESOLVED 2026-05-01**

跑 cross-form matrix 时 LookAtOffice 用 top office code (80944 = 典史，37,847 个 postings) 跑出 0 行。

诊断历程（2026-05-01）:
1. ✓ ZZ_OFFICE_CODE picker 表写入正确 (1 行)
2. ✓ 控件 TxtTypeDesc='[All]' 设置正确（绕开 N/A 分支）
3. ✓ Form_Open 会清空 ZZ_OFFICE_CODE（其他 form 不会）→ 已修：populate AFTER open
4. ✓ JET 缓存一致性问题（pyodbc 写、Access 内 SQL 读看不到）→ 已修：set_picker_codes 后调 `DBEngine.Idle 8 + RefreshDatabaseWindow`
5. ✓ autodetect 注入正确（marker v4 + gUseOfficeID 行都在）
6. ✓ 同样的 SQL pyodbc 直跑 → 37,429 行（确认 SQL 本身对的）
7. ✗ 但 VBA click 完成后 ZZ_SCRATCH_OFFICE = 0 行
8. ✗ ZZ_TEST_DEBUG 表里没有 "ENTERED CmdQuery_Click" → CmdQuery_Click 根本没被触发
9. ✗ Application.Run 一个 Public wrapper sub 也没真正执行 (form module sub 不可达)

**根本原因**: Form_LookAtOffice 的 CmdQuery 在 form-designer 里默认 `Enabled=False`，必须由用户先点 CmdAllOffices 或 CmdPickOffice 才会启用。pywinauto.click_input 对 disabled 控件被 Windows 静默丢弃。force_enable (COM 设 `Enabled=True`) 在 COM 端确认生效，但 UIA tree 仍然 cache 为 disabled，所以 pywinauto 看到的还是 disabled。

**修复（2026-05-01）**: 注入 Form_Timer 触发器到 form module，从 Python 设 `Forms("LookAtOffice").TimerInterval = 100` 触发 — Access 自己 fire Form_Timer，绕开 click + 绕开 Application.Run-can't-reach-form-module-sub 两个坑。`vba_session.click_via_timer` + `_inject_timer_trigger` 实现。

另两个相关坑也一起修了：
- **Done 信号**: backfill UPDATE 不改 row_count，单凭 row_count poll 不知道 CmdQuery_Click 真正完成。注入了 `INSERT INTO ZZ_TEST_DEBUG VALUES ('<form>:DONE')` 在 Exit_<sub>: 之前，Python poll 那个 marker。
- **pd.read_sql 死锁**: 大表 (37k 行) full select * via pandas read_sql 与 Access 内部 recordset binding 死锁。换成 raw pyodbc cursor + fetchall 解决。

**当前状态**: `tests/test_vba_matrix_all_forms.py::test_cross_form_matrix[office_80944_unfiltered]` 通过，37,429 行 + 完整 integrity check。

注：source-SQL differential check 对 Office 暂时跳过（Access 还持有 POSTED_TO_OFFICE_DATA linked-table read lock，pyodbc 死锁等）。其他 6 维度 (row count / column structure / FK integrity 等) 都跑了。

---



### Bug #1 — `View_StatusData` 别名错位（影响显示，不影响选择）

`View_StatusData` 是 `STATUS_DATA_2 Subform` 的 `RecordSource`，每次用户点开人物详情查看「身份/状态」面板时都会用到。

**症状**：返回行的 `c_fy_range_desc` / `c_fy_range_chn`（首年范围描述）显示的是 last-year 的范围值，而不是 first-year 的。

**根因**（`analysis/dump/queries.json`，`View_StatusData`）：

```sql
... LEFT JOIN YEAR_RANGE_CODES ON STATUS_DATA.c_fy_range = YEAR_RANGE_CODES.c_range_code
... LEFT JOIN YEAR_RANGE_CODES AS YEAR_RANGE_CODES_1 ON STATUS_DATA.c_ly_range = YEAR_RANGE_CODES_1.c_range_code
```

但 SELECT 列里 4 个范围别名都从 `YEAR_RANGE_CODES_1` 取：

```
YEAR_RANGE_CODES_1.c_range     AS c_fy_range_desc   ← 应该用 YEAR_RANGE_CODES
YEAR_RANGE_CODES_1.c_range_chn AS c_fy_range_chn    ← 应该用 YEAR_RANGE_CODES
YEAR_RANGE_CODES_1.c_range     AS c_ly_range_desc   ← 正确
YEAR_RANGE_CODES_1.c_range_chn AS c_ly_range_chn    ← 正确
```

**修复**：改为

```
YEAR_RANGE_CODES.c_range     AS c_fy_range_desc
YEAR_RANGE_CODES.c_range_chn AS c_fy_range_chn
```

**回归测试**：`tests/test_known_bugs.py::test_bug_view_statusdata_fy_alias_swap`、`test_bug_view_statusdata_fy_value_equals_ly_value`。修复后这两个测试会失败 —— 是修复成功的信号，按提示更新断言即可。

---

### Bug #2 — VBA 引用 `dao360.dll` 已断（影响表单自动化与新机器打开）

`CBDB_BJ_User.mdb` 的 VBA 项目中一条 references 指向 `C:\Program Files\Common Files\Microsoft Shared\DAO\dao360.dll`，这是 Access 2003 时代的 DAO 3.6 路径。新版 Office (2016+) 用的是 `ACEDAO.DLL`。

**症状**：
- 在没有装 legacy DAO 的机器上，所有表单的 `Form_Open` 报「Can't find project or library」。
- 自动化脚本（包括本测试套件本来想做的 COM-driven UI 测试）打不开任何表单。

**修复**：
- 用 Access 打开 .mdb，`Alt+F11` 进 VBE。
- `Tools → References`，去掉勾选的 `MISSING: dao360.dll`。
- 勾上 `Microsoft Office 16.0 Access Database Engine Object Library`（即 ACEDAO.DLL）。
- 保存。

**已自动化**：`analysis/check_vba_refs.py` 会检查并修复（在工作副本上）。

---

### Bug #4 — `LookAtPlace.CmdGIS_Click` 引用了不存在的控件 `GISFrame`（**用户点 GIS 必报错**）

`Form_LookAtPlace.vb:1539` 写：

```vba
If GISFrame.Value = 1 Then
    tStream.Charset = "utf-8"
    ...
Else
    tStream.Charset = "gb18030"
End If
```

但 LookAtPlace 表单上**没有名为 `GISFrame` 的控件** —— 只有 `CodeFrame` (OptionGroup, ControlType=107) 和 `GephiFrame`。整个 `Form_LookAtPlace.vb` 中 `GISFrame` 只出现这一次。

显然是开发者从其他 form（Status/Texts/Associations 用 `GISFrame`）拷贝 `CmdGIS_Click` 时漏改了控件名 —— 同一个文件的 `CmdNeo4j_Click` / `CmdGephi_Click` / `CmdPajek_Click` 都用 `CodeFrame.Value`，证明 Place 的 encoding 选择器就是 `CodeFrame`。

**症状**：用户在 LookAtPlace 跑完查询，点 **GIS** 按钮 → VBA 抛 "Object required" → MsgBox 弹出 "Object required" → 用户按确定 → `Resume Exit_CmdGIS_Click`（不写文件）。**LookAtPlace 的 GIS 导出功能在所有用户机器上都是坏的**。

发现路径：`tests/test_vba_cmdgis_other_forms.py::test_cmd_gis_produces_file[place_addr_7213]` 反复失败 → 注入 step markers + 修好 Err handler 的 SQL 注入 bug（之前 Err handler 自身因 SQL 语法错误而沉默） → ZZ_TEST_DEBUG 显示 `LookAtPlace:STEP1 → STEP2 → STEP3 → ERR Object required`，对应到代码就是 `Set tStream = New ADODB.Stream` 和 `Set dlgSaveAs = ...` 之间唯一的语句 `If GISFrame.Value = 1 Then`。

**修复方向**：把 `GISFrame.Value` 改成 `CodeFrame.Value`（与 Place 上其他 export 子保持一致）。然后 GIS 输出会按 CodeFrame 的 1=UTF-8 / 2=BIG5 / 3=GB2312 切换编码（注意 Place 的 CodeFrame 用 BIG5/GB2312，而其他 form 的 GISFrame 用 GB18030 —— 修的时候要保证用户预期）。

**测试 workaround**：`tests/cbdb_driver/vba_session.py` 的 inject 在 Place 加载时把 CmdGIS_Click 里的 `GISFrame.Value` 替换成 `CodeFrame.Value`。打勾 `tests/test_vba_cmdgis_other_forms.py::test_cmd_gis_produces_file[place_addr_7213]`。

---

## 待分类的发现（差异性测试）

### Open Question #1 — User MDB vs cbdb-online-main-server 的 `c_index_year` / `c_index_addr_id` 不一致

`tests/test_index_year_xcheck.py` 在最新一次全量跑（657246 共有 person id）发现：
- **83 個 `c_index_year` 差异**（0.013%）
- **492 個 `c_index_addr_id` 差异**（0.075%）

样例（`(c_personid, MDB(year, addr_id), SQLite(year, addr_id))`）：

```
c_index_year 差异：
  (1455, (1001, 12887), (1008, 12887))   ← 7 年差
  (3501, (1018, 100658), (1028, 100658)) ← 10 年差
  (16266, (1004, 12723), (992, 12723))   ← 12 年差
  (31573, (585, 15294), (1056, 15294))   ← 471 年差！可能 Tang vs Song
  (40258, (1157, 12780), (None, None))   ← MDB 有，SQLite 全空
  (40848, (985, None), (None, None))
  (46642, (999, 100409), (None, None))

c_index_addr_id 差异：
  (1, (1042, 101117), (1042, None))      ← MDB 有 addr，SQLite 空
  (470, (1012, 12879), (1012, 12785))    ← addr 不同
  (481, (1043, 100416), (1043, 12785))   ← 100416 → 12785（同一模式）
  (485, (1046, 100416), (1046, 12785))   ← 100416 → 12785
  (562, (1067, 100658), (1067, 13292))   ← 开封 → 别处
  (927, (None, 100658), (None, 12449))
  (1005, (1128, 12793), (1128, 11232))
```

**待分类**：每条差异可能是 (a) 派生规则差异 — 两边对 "index year" 的优先级定义不同（生年 vs 进士年 vs 卒年）；(b) 数据时序差异 — 两个快照不同时间生成，中间有源数据更新；(c) 一边有 bug。第一步：抽样 10-20 个差异手工验证 BIOG_MAIN 和派生源（YEAR_RANGE_CODES、ENTRY_DATA、KIN_DATA 等），把多数差异归到 (a) 或 (b)，剩下的标记为可能的 bug。

**复现**：
```bash
python analysis/download_hf_sqlite.py
CBDB_FULL_XCHECK=1 python -m pytest tests/test_index_year_xcheck.py -v -s
```

---

## 启发式扫描结论（无问题）

| 扫描器 | 结果 |
|---|---|
| `analysis/audit_view_aliases.py` — YEAR_RANGE_CODES 别名错位 | View_PeopleData 的 `c_da_*` 别名是 death-age 的简写，假阳；其余唯一一处真阳即 Bug #1 |
| `analysis/audit_duplicate_aliases.py` — SELECT 列重复别名 | 无 |

---

## 架构观察（非 bug，但写测试时需要知道）

1. **链接表**：`BIOG_MAIN`、`*_DATA`、`*_CODES`、`ADDR_CODES` 等 ~120 张表都是从 `CBDB_*_DATA.mdb` 链接的（`tables.json` 里 `record_count == -1` 是这个标识）。`CBDB_BJ_User.mdb` 自己只持有 ~63 张本地的 `ZZ_*` / `Z_*` 工作表。

2. **查询入口**：每个 `LookAtXxx` 表单都有一个 `Private Sub CmdQuery_Click`。这个事件：
   - 清空 `ZZ_SCRATCH_<XXX>` 输出表
   - 根据用户输入（控件值 + picker 写入的 `ZZ_SCRATCH_<CODE>` 表 + 公共全局变量 `gUseADDRID` 等）拼一段 `INSERT INTO ZZ_SCRATCH_<XXX> SELECT ...`
   - 跑若干 `UPDATE` 把描述性字段从 `*_CODES` 表回填到结果表
   - 把表单的 `RecordSource` 重新指向 `ZZ_SCRATCH_<XXX>`

3. **picker 表单约定**：所有 `frmPickXxx_multi` 都把用户选中的 ID 写入对应的 `ZZ_SCRATCH_<XXX>` 表（如 `ZZ_SCRATCH_ENTRY_CODE`、`ZZ_SCRATCH_ADDR`）。测试时可以直接 `INSERT INTO` 这些表绕过 picker UI。

4. **私有事件不能 `Application.Run`**：`CmdQuery_Click` 默认是 `Private`，无法通过 `Application.Run "Form_X.CmdQuery_Click"` 从外部触发。要么改 Public、要么用 SendKeys、要么走 Python SQL replay 路径（本测试套件采用后者）。

5. **HelpFile 数字会随 CBDB 数据更新而漂移**：`HelpFile_LookAtEntry.pdf` 给的「凯封 yin general 900-1100 = 104 人」例子，在当前数据上是 103 人（漂移 -1）；entry-years 版本是 12 人（HelpFile 写 11，漂移 +1）。这印证了 (a) 我们的 SQL replay 是正确的，(b) HelpFile 数字应该作为「软参照」（5-20% 容差）而非硬断言。

---

## 未审查 / 待补全

- 复杂 picker 的边界条件（多选、子单位展开、XY 半径展开）—— LookAtEntry 实现了 sub-units 和 XY 展开但没有为它们写 fixture。
- KIN 网络递归（`LookAtKinship` 1-hop 已 covered，但 `LookAtNetworks` / `LookAtAssociationPairs` / `LookAtGroupData` 的多跳递归 (`clsTreeView` 深度遍历) 还没跑通 — 当前 fixture (Zhu Xi 2471 assocs / Wang Anshi×Sima Guang 47-edge) 在 120s timeout 内回不来。`LookAtAssociationPairs` 已确认根因：`Form_LookAtAssociationPairs.Link1stOrder` 的 ASSOC_DATA 自联 (`ZABA INNER JOIN ZABA_1 ON ZABA.c_assoc_id = ZABA_1.c_assoc_id`) 在 JET 优化器里 *先 join 再 filter*，对任何高 assoc 数的 person 都会 materialize 巨大中间结果。设 `Chk2Nodes=0`/`ChkKinship=0` 也无效。要想跑通，需要找一对 < 10 个 association 的人，或者重写 VBA 用 saved query 预过滤。
- 双语切换（`changeDisplayLanguage` 改了大量 Label 的 `Caption`，没有测试覆盖）。
- 其他 export 按钮：`tests/test_vba_export.py` 当前只 cover LookAtEntry CmdGIS。CmdNeo4j / CmdKML / CmdPajek / CmdGUESS / CmdGephi 待补 (chain pattern 已经验证，加 fixture 即可).

---

## 自动化覆盖状态（2026-05-01）

| 表单 | 真 VBA matrix | 真 export | 备注 |
|---|---|---|---|
| LookAtEntry | ✅ `test_vba_matrix.py` | ✅ CmdGIS | Bug #3 已确认在此 form |
| LookAtStatus | ✅ 3 fixtures | — | 17k+4.9k 行 |
| LookAtTexts | ✅ biblcat 1 | — | 15.7k 行 |
| LookAtAssociations | ✅ 3 fixtures | — | 11.8k 行 |
| LookAtOffice | ✅ 2 fixtures | — | 37k+35k 行；Form_Timer landmine 已解 |
| LookAtPlace | ✅ 2 fixtures | — | 5.9k+3.5k 行 |
| LookAtKinship | ✅ 1 fixture | — | 949 行（赵廷美 1-hop）|
| LookAtAssociationPairs | ⏭ skip | — | CmdQuery 90s timeout |
| LookAtNetworks | ⏭ skip | — | recursive expansion 太重 |
| LookAtGroupData | ⏭ skip | — | 同上 |

**当前总计**: 7/10 forms 真 VBA matrix-tested + 1 form 真 export-tested。Matrix run 12 passed + 3 skipped in 114s。

**更新（2026-05-02）**: README "Plan & status" 表格是单一真相来源；findings.md 这一节滞后了。最新覆盖请见
[`README.md` § Plan & status](README.md#plan--status)。
本会话期间发现并已确认 Bug #4（LookAtPlace 的 CmdGIS 引用 GISFrame 笔误），见上文。
