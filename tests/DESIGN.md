# 全控件自动化测试设计

**目标**：覆盖 `CBDB_BJ_User.mdb` 中**每一个**有事件处理的按钮 / 复选框 / 选项框 / 文本框，验证（a）能跑通不报错（b）副作用（结果表 / 文件 / 控件状态）符合预期。

## 1. 控件清单（来自 `analysis/dump/control_inventory.json`）

| 维度 | 数 |
|---|---:|
| 总 form 数 | 79 |
| 交互控件总数（CommandButton/CheckBox/OptionGroup/TextBox/ListBox/ComboBox/Subform） | 1744 |
| 有 VBA 事件处理的控件 | 370 |
| 10 个主 LookAt form 中的有事件控件 | 230 |

**按行为聚类（10 个 LookAt form, 230 控件）**：

| 桶 | 数 | 代表 | 测试代价 |
|---|---:|---|---|
| `export` | 57 | CmdGIS / CmdNeo4j / CmdGephi / CmdPajek / CmdUCINet / CmdGUESS / writeKML / CmdSaveXxx | 高（需要拦截 SaveAs dialog） |
| `checkbox` | 49 | ChkSubUnits / ChkXYRef / Chk* (LookAtNetworks 28 个) | 中（联动逻辑多） |
| `picker` | 35 | CmdPickEntry / CmdSelectPlace / CmdFromDynasty / CmdToDynasty | 高（modal 子表单） |
| `unknown` | 26 | CmdStoreID / CmdRecallID / CmdRun / CmdClose / TxtFromYear_LostFocus | 低-中 |
| `display` | 20 | CmdFanti / CmdJianti / changeDisplayLanguage | 低 |
| `state` | 18 | CmdAllDynasties / CmdAllPlaces / CmdClearList | 低 |
| `nav` | 10 | CmdExit / CmdHelp | 极低 |
| `frame` | 8 | FrameYears / FrameFilterYears (OptionGroup) | 中 |
| `query` | 7 | CmdQuery / CmdRun | 中（已部分覆盖） |

加上 picker 子 form 自身（`frmPickEntry_multi` 等 8 个 picker，每个 ~5-6 个 control）和子表单 `*_2 Subform`（11 个）的事件，整体覆盖目标约 **400 个 (form, control) 测试点**。

---

## 2. 三个底层技术问题及解决方案

### P1: VBA 引用 `dao360.dll` 已断
- 解决：每次开 .mdb 副本前，自动移除断引用 + 加 ACEDAO.DLL（已在 `analysis/check_vba_refs.py` 实现）。
- 写进 `tests/cbdb_driver/access_app.py` 的 `open()` 里。

### P2: `Private Sub Cmd*_Click` 不能从外部 `Application.Run`
**两条路（推荐路 B）**：

- 路 A — 改 access modifier：用 VBE 把所有 `Private Sub Cmd*_Click` 改写成 `Public`。`AddFromString` 不行（会重复 sub），需要用 `CodeModule.Lines` 读出来 + `ReplaceLine` 改一行。脆弱。
- 路 B — 注入 dispatcher：在 std module 里加一个公共 sub `TestInvoke(formName, eventName)`，内部用 `Application.VBE.CodeModule.Run`（VBE 的 Run 能跨 access 调）。或者更简单：用 `Application.Run "VBA.Forms!" & formName & "." & eventName`（Access 实际上接受这种语法访问 form module 的 Public 成员）。

**实测可行的方案**：在每次测试前给指定 form 的指定事件**注入一个 Public 包装**（如 `Public Sub TestRun_CmdQuery() : Call CmdQuery_Click : End Sub`），然后用 `Forms!FormName.TestRun_CmdQuery` 触发（这个语法实测不行，但 `Forms("FormName").TestRun_CmdQuery` 借助 form 实例可以；最稳的还是改成 Public）。

最终选择：**把所有 Cmd*_Click 自动改 Public 一次性**，写一个 `analysis/make_handlers_public.py`，只在测试副本上做。

### P3: `Application.FileDialog(msoFileDialogSaveAs)` 是 modal，会卡住测试
**两条路（推荐路 B）**：

- 路 A — 用 SendKeys "{ENTER}" 提前回车确认。脆弱。
- 路 B — **重写 export handler 把 FileDialog 替换成 fixed path**。每个 form 的 export sub 都遵循同一个模式：
  ```vba
  Set dlgSaveAs = Application.FileDialog(msoFileDialogSaveAs)
  dlgSaveAs.InitialFileName = "entry_gis_" + tCodeStr + ".tab"
  If dlgSaveAs.Show = -1 Then
      tFileName = dlgSaveAs.SelectedItems(1)
      ' ... write file
  End If
  ```
  我们注入一个 `g_TestExportPath` 变量，再写一段 monkey-patch 把这段替换成「if g_TestExportPath <> "" then tFileName = g_TestExportPath else use dialog」。一次性脚本搞定所有 export handler。

---

## 3. 9 类控件的测试设计

### 3.1 `query` 桶（7 个，已覆盖 1）

**模式**：CmdQuery_Click 把所有当前输入拼成 `INSERT INTO ZZ_SCRATCH_<XXX> SELECT ...`，写入工作表。

**测试方案**：
```python
def run_query_test(form, fixture):
    drv.open_form(form)
    fixture.populate_pickers(drv)         # INSERT INTO ZZ_SCRATCH_<CODE> ...
    fixture.set_controls(drv)             # 设 TxtFromYear / FrameYears / ChkXxx ...
    drv.invoke_event(form, "CmdQuery_Click")
    # 1. 没有 raise
    # 2. ZZ_TEST_ERRORS 为空（捕获 VBA OnError MsgBox）
    # 3. ZZ_SCRATCH_<XXX> 行数与 golden 一致
    # 4. ZZ_SCRATCH_<XXX> 内容（按 PK 排序）与 golden CSV 一致
```

**fixture 数量**：每 form 至少 5 个（覆盖主要分支：无 picker × 单 picker × 多 picker × 子单位展开 × XY 展开）→ 35 个测试。
**已实现**：LookAtEntry 5 个 fixture（用 Python SQL replay，未驱动 VBA）。
**待补**：9 form × 5 fixture = 45 个，全部走 COM driver。

### 3.2 `picker` 桶（35 个）

**模式**：弹出一个 modal 子 form (`frmPickEntry_multi` 等)，用户选 N 项，子 form 关闭后把选择写到对应的 `ZZ_SCRATCH_<CODE>` 表 + 父 form 的 `TxtXxx` 显示选中项的描述。

**子模式**：
- `CmdPick<X>` → 多选 picker，写 `ZZ_SCRATCH_<CODE>`
- `CmdSelectPlace` / `CmdSelectPerson` → 选地址 / 人，写 `ZZ_SCRATCH_ADDR` / `ZZ_STORE_PERSON_ID`
- `CmdFromDynasty` / `CmdToDynasty` → 选朝代，写父 form 全局 `gFromDynasty`、`gFromDynastyBegin`、`gFromDynastyEnd`

**测试方案**：
```python
# 验证 1: 子 form 能被打开 (smoke)
def test_picker_opens(form, picker_button):
    drv.open_form(form)
    drv.invoke_event(form, picker_button + "_Click")
    # 子 form 在 acDialog 模式下打开会阻塞，所以我们用「窗口列表」检查
    # 通过 mock：注入测试模式，让子 form 直接关闭并写入 mock 选择
    ...

# 验证 2: 选择写入正确的目标表
def test_picker_writes_correct_table(form, picker_button, expected_table):
    setup_picker_mock(picker_button, mock_selection=[1, 2, 3])
    drv.invoke_event(form, picker_button + "_Click")
    rows = drv.read_table(expected_table)
    assert set(rows.iloc[:, 0]) == {1, 2, 3}
```

**fixture 数量**：35 个 picker × 2 个测试（开 + 写）= 70 个。
**关键技巧**：注入一个 stub 子 form 替换 `frmPickEntry_multi` 等，stub 直接写入选中项后关闭。这是 8 个 stub 子 form 的工作量。

### 3.3 `export` 桶（57 个）

**模式**：从已填充的 `ZZ_SCRATCH_<XXX>` 读数据，写到磁盘文件（.tab / .kml / .csv / .gephi / .pajek / .ucinet）。

**测试方案**：
```python
def test_export(form, query_fixture, export_button, expected_format):
    # 1. 先跑 query 让 ZZ_SCRATCH_<XXX> 有数据
    run_query_test(form, query_fixture)
    # 2. 设 g_TestExportPath
    drv.set_global("TestHelpers", "g_TestExportPath", str(tmp_path / "out"))
    # 3. 触发 export
    drv.invoke_event(form, export_button + "_Click")
    # 4. 文件应该存在
    out = tmp_path / "out" + extension_for(expected_format)
    assert out.exists() and out.stat().st_size > 0
    # 5. 行数 = ZZ_SCRATCH_<XXX>.RecordCount + 1 (header)
    # 6. 表头与 golden 一致
    # 7. 每行字段数与 golden 一致（catches "丢栏位"）
    # 8. 全文与 golden diff（最严格，catches "栏位错位 / 丢数据"）
```

**fixture 数量**：每 form 平均 6 个 export × 2-3 个数据规模 = ~150 个测试。每个有一个 golden 文件。

**特别处理**：
- **GIS (.tab)**：用 tab 分隔，UTF-8 / GB18030 两种编码（GISFrame 控件控制），golden 双份。
- **KML**：XML 格式，golden 也是 XML，diff 时按行比较。
- **Neo4j / Pajek / UCINet / Gephi**：网络文件，golden 是节点 + 边两种文件，需检查节点数 = 边的端点 ID 集合。

### 3.4 `state` 桶（18 个）

**模式**：清掉某项约束。例如 `CmdAllDynasties_Click` 把 `gFromDynasty` 设为 -2（"all"），并清空 `TxtFromDynasty.Value`。

**测试方案**：
```python
def test_state_reset(form, button, target_global, target_textbox):
    # pre-populate
    drv.set_global(form, target_global, 15)  # 假设 = 宋
    drv.set_control(form, target_textbox, "宋")
    # action
    drv.invoke_event(form, button + "_Click")
    # assert
    assert drv.get_global(form, target_global) == -2
    assert drv.get_control(form, target_textbox) in ("", None)
```

**fixture 数量**：18 × 1 = 18 个（每个 button 一个 setUp/assert）。

### 3.5 `checkbox` 桶（49 个）

**模式**：复选框影响 `CmdQuery` 的 SQL 拼接。LookAtEntry 的 `ChkSubUnits` / `ChkXYRef` 是简单标志；LookAtNetworks 的 49 个复选框组成层级（`ChkPoliticsAll` 控制 4 个子）。

**测试方案（双层）**：
```python
# Layer A: 联动测试 — 父复选框 toggle 时子复选框是否同步
def test_checkbox_cascade(form, parent_chk, child_chks):
    drv.set_control(form, parent_chk, True)
    for c in child_chks:
        assert drv.get_control(form, c) == True
    drv.set_control(form, parent_chk, False)
    for c in child_chks:
        assert drv.get_control(form, c) == False

# Layer B: 行为测试 — 复选框影响 query 结果
def test_checkbox_affects_query(form, chk):
    base = run_with(chk=False)
    toggled = run_with(chk=True)
    assert set(base.c_personid) != set(toggled.c_personid), \
        f"{chk} doesn't affect query result -- dead code?"
```

**fixture 数量**：49 × 1.5 = ~75 个。
**特别**：LookAtNetworks 的 28 个 ChkSch* / ChkPol* / ChkWri* / ChkMilitary* 是测试重灾区，需要专门 fixture。

### 3.6 `frame` 桶（8 个）

**模式**：OptionGroup（FrameYears / FrameFilterYears），值 1/2/3 决定年限解释方式（entry / index / dynasty）。

**测试方案**：
```python
@pytest.mark.parametrize("mode", [1, 2, 3])
def test_frame_year_mode(form, mode):
    drv.set_control(form, "TxtFromYear", 900)
    drv.set_control(form, "TxtToYear", 1100)
    drv.set_control(form, "FrameYears", mode)
    drv.invoke_event(form, "CmdQuery_Click")
    # 不同 mode 下的 distinct person 数应不同
```

**fixture 数量**：8 × 3 = 24 个。

### 3.7 `display` 桶（20 个）

**模式**：CmdFanti/CmdJianti 调用 `changeDisplayLanguage`，把所有 Label / 按钮的 `Caption` 在繁/简/英三种之间切换。Caption 的繁简对照表存在 `FormLabels` 表（643 行）。

**测试方案**：
```python
def test_display_language_switch(form):
    base_caps = {c.Name: c.Caption for c in drv.iter_labels(form)}
    drv.invoke_event(form, "CmdJianti_Click")
    after_caps = {c.Name: c.Caption for c in drv.iter_labels(form)}
    # 至少 N 个 label 的 caption 改变
    changed = sum(1 for n in base_caps if base_caps[n] != after_caps[n])
    assert changed > 5, "language switch did nothing"
    # 切回 fanti 后应还原
    drv.invoke_event(form, "CmdFanti_Click")
    assert {c.Name: c.Caption for c in drv.iter_labels(form)} == base_caps
```

**fixture 数量**：10 form × 1 = 10 个。

### 3.8 `nav` 桶（10 个）

**CmdExit_Click**：仅 `DoCmd.Close`。测试：调用后 form 不在 `Forms` 集合中。
**CmdHelp_Click**：`Application.FollowHyperlink "HelpFile_<X>.pdf"`。测试：捕获 hyperlink 调用（用 mock）或直接断言文件存在。

**fixture 数量**：10 个。

### 3.9 `unknown` 桶（26 个，需细化）

| Sub-bucket | 控件 | 测试 |
|---|---|---|
| `CmdRun_Click`（LookAtKinship/Networks/GroupData 实际查询入口） | 3 | 当作 `query` 处理 |
| `CmdStoreID_Click`（保存当前结果 ID 到 `ZZ_STORE_PERSON_ID`） | 10 | 跑完 query 后 click，断言 `ZZ_STORE_PERSON_ID` 的内容 = `ZZ_SCRATCH_<XXX>.c_personid` |
| `CmdRecallID_Click`（从 `ZZ_STORE_PERSON_ID` 回填） | 3 | 预填 store, click, 断言下游表 |
| `CmdClose_Click`（同 CmdExit） | 3 | nav |
| `Txt*_LostFocus`（输入校验） | 7 | 设非法值，触发 LostFocus, 断言 MsgBox 出现或 control reset |

**fixture 数量**：26 个。

---

## 4. 总测试规模

| 桶 | 测试数 |
|---|---:|
| query | 50（10 form × 5 fixture） |
| picker | 70（35 picker × 2） |
| export | 150（10 form × 5 fixture × 3 export） |
| state | 18 |
| checkbox | 75 |
| frame | 24 |
| display | 10 |
| nav | 10 |
| unknown | 26 |
| **小计** | **~430** |
| picker stub form 集成测试 | 16 |
| sub-form 数据绑定测试 | 22 |
| schema / saved view（已实现） | 50 |
| **总计** | **~520** |

每测试在我机器上估算 200-500 ms（一个 form 操作），全套约 2-5 分钟。

---

## 5. 基础设施实现清单

### 5.1 `tests/cbdb_driver/`（新模块，与 `cbdb_replay/` 并列）

```
tests/cbdb_driver/
├── __init__.py
├── access_app.py          # AccessApp: 管 Access COM 进程 + 工作副本 + 自动修引用
├── form_driver.py         # FormDriver: open/close/set_control/get_control/invoke_event
├── vba_inject.py          # 一次性把所有 Private Cmd*_Click 改 Public + 注入 ZZ_TEST_ERRORS 钩子
├── picker_stub.py         # 注入 stub 替代 frmPick* 子 form
├── export_capture.py      # 把所有 export sub 的 FileDialog 替换成固定路径
└── error_catcher.py       # 全局 OnError → 写入 ZZ_TEST_ERRORS
```

### 5.2 mdb 改造（一次性、写入测试副本）

`vba_inject.py` 的逻辑：
```vba
' 1. 在每个 form 模块开头加：
On Error GoTo TestErrHandler

' 2. 在每个 form 的 Form_Open 末尾加：
Call TestHelpers.SuppressMsgBox

' 3. 把所有 Private Sub Cmd*_Click → Public Sub Cmd*_Click

' 4. 在新建的 TestHelpers 标准模块里加：
Public Sub SuppressMsgBox()
    ' MsgBox 替换：写入 ZZ_TEST_ERRORS 而不弹框
End Sub

Public Sub LogError(formName As String, eventName As String, errDesc As String)
    DoCmd.SetWarnings False
    CurrentDb.Execute "INSERT INTO ZZ_TEST_ERRORS (form_name, event_name, err_desc, ts) " & _
        "VALUES ('" & formName & "', '" & eventName & "', '" & Replace(errDesc, "'", "''") & "', Now())"
End Sub

Public g_TestExportPath As String

' 5. 新建表 ZZ_TEST_ERRORS (form_name TEXT, event_name TEXT, err_desc MEMO, ts DATE)
```

### 5.3 测试 fixture 组织

```
tests/
├── cbdb_driver/             # 新
├── cbdb_replay/             # 已有
├── golden/
│   ├── lookatentry/         # 已有
│   ├── exports/             # 新：每个 form/export 的 golden 文件
│   │   ├── lookatentry_gis_kaifeng_yin.tab
│   │   ├── lookatentry_kml_kaifeng_yin.kml
│   │   ├── lookatentry_neo4j_kaifeng_yin.cypher
│   │   └── ...
│   └── controls/            # 新：control state golden（如 caption 切换前后对照）
├── fixtures/
│   ├── pickers/             # picker stub 子 form (.bas)
│   └── inputs/              # 标准化的输入 fixture (yaml/json 形式)
├── test_query_*.py          # 10 form 各一个文件
├── test_export_*.py
├── test_picker_*.py
├── test_state_*.py
├── test_checkbox_*.py
├── test_frame_*.py
├── test_display_*.py
├── test_nav_*.py
└── test_unknown_*.py
```

### 5.4 共用 fixture（pytest）

```python
@pytest.fixture(scope="session")
def driver():
    """Session-wide AccessApp instance against a fresh working copy."""
    with working_copy(USER_MDB, TEST_COPY) as copy:
        app = AccessApp(copy).open()
        VbaInjector(app).fix_refs().make_handlers_public().inject_helpers()
        yield app
        app.close()


@pytest.fixture(scope="function")
def clean_state(driver):
    """Reset all ZZ_SCRATCH_* + ZZ_TEST_ERRORS between tests."""
    driver.exec_sql("DELETE FROM ZZ_TEST_ERRORS")
    for tbl in SCRATCH_TABLES:
        driver.exec_sql(f"DELETE FROM [{tbl}]")
    yield
```

### 5.5 通用断言

```python
def assert_no_vba_errors(driver):
    rows = driver.read_table("ZZ_TEST_ERRORS")
    assert len(rows) == 0, (
        "VBA error(s) raised:\n" + rows.to_string()
    )

def assert_export_matches_golden(actual_path, golden_path,
                                 *, normalize_whitespace=True):
    a = actual_path.read_text(encoding="utf-8")
    g = golden_path.read_text(encoding="utf-8")
    if normalize_whitespace:
        a = "\n".join(line.rstrip() for line in a.splitlines())
        g = "\n".join(line.rstrip() for line in g.splitlines())
    assert a == g, ...

def assert_table_matches_golden(driver, table, golden_csv, *, sort_by):
    df = driver.read_table(table)
    assert_matches_golden(df, golden_csv, sort_by=sort_by, ...)
```

### 5.6 控件覆盖追踪

新增 `tests/test_coverage.py`：
```python
def test_every_control_with_handler_has_a_test():
    """枚举 control_inventory.json 中每个 (form, control)，
    验证至少有一个测试 case 标记了它（用 pytest mark）。"""
    inv = load_inventory()
    expected = {(f, c["name"]) for f, info in inv.items()
                for c in info["controls"] if c["events"]}
    tested = collect_tested_controls()  # 扫描 pytest mark
    missing = expected - tested
    if missing:
        pytest.xfail(f"{len(missing)} controls untested:\n" +
                     "\n".join(f"  {f}.{c}" for f, c in sorted(missing)))
```

---

## 6. 分阶段实现路线

### Phase 1（约 2-3 天工作量）— 基础设施
- [ ] `tests/cbdb_driver/access_app.py` + `form_driver.py`（~200 行）
- [ ] `vba_inject.py`（含 Public 化、ZZ_TEST_ERRORS 表、错误钩子，~300 行）
- [ ] `export_capture.py`（FileDialog → fixed-path patch，~150 行）
- [ ] `picker_stub.py`（8 个 stub 子 form，~200 行）
- [ ] 验收测试：`test_infrastructure.py` 跑通 LookAtEntry 的 1 个 query + 1 个 export + 1 个 picker

### Phase 2（约 2-3 天）— 主 form 全覆盖
- [ ] `test_query_*.py` × 10（用 driver 重写 LookAtEntry，迁移其 5 个 fixture；其他 9 form 各加 5 个）
- [ ] `test_export_*.py` × 10（每 form 5-7 个 export × 2 个数据）
- [ ] `test_picker_*.py` × 10（每 form 4-6 个 picker × 2 个测试）
- [ ] `test_state_*.py` × 10
- [ ] `test_nav_*.py` × 10

### Phase 3（约 1-2 天）— 复杂控件
- [ ] `test_checkbox_*.py`（重点 LookAtNetworks 28 个）
- [ ] `test_frame_*.py`
- [ ] `test_display_*.py`
- [ ] `test_unknown_*.py`（StoreID/RecallID/LostFocus）

### Phase 4（约 1 天）— 子 form / 集成
- [ ] picker 子 form 的独立测试
- [ ] `*_2 Subform` RecordSource 绑定测试
- [ ] CBDB_Browser_2 端到端流程测试
- [ ] 控件覆盖追踪测试

### Phase 5（持续）— Golden 维护
- [ ] CI hook：跑 `pytest --regenerate-goldens` 后 `git diff --stat tests/golden/` 输出到 PR
- [ ] 文档：每次 .mdb 更新前后该如何 review goldens

---

## 7. 不在覆盖范围内（明确）

- **手动 GIS 可视化**：导出后 .tab 文件需用 ArcGIS / QGIS 打开，自动化测不到地图正确性
- **PDF Help 链接**：CmdHelp 打开 PDF，PDF 内容质量不测
- **Access UI 渲染**：按钮位置 / 颜色 / 字体不测
- **跨 Access 版本兼容性**：只在 Office 16 (2016+) 上测
- **多用户并发 / 锁**：CBDB User 是单用户应用，不测
- **Recordset Bookmark / Find Next**：每个 form 的 datasheet 内置「查找」功能用 Access 自带的，不测

---

## 8. 风险与开放问题

1. **改 Public 后能否还原**：测试副本即扔，原 mdb 不动，无影响。但若用户的 .mdb 本身依赖 Private 语义（如同名 sub），就要更细粒度。**实测前先 grep**：
   ```bash
   grep -r "Private Sub Cmd.*_Click" analysis/dump/vba/
   ```
   若有同名 sub 被多处调用，需要保留命名空间。
2. **picker stub 维护**：8 个 stub 子 form 的 .bas 源码需与对应 picker 的接口同步；picker 改了我们也要改。
3. **导出 golden 编码稳定性**：UTF-8 BOM、行尾 CR/LF、中文 punctuation 不同输入法可能不一致 → 比较前 normalize。
4. **测试时长**：~520 测试 × ~300 ms = ~3 分钟，可接受。但如果每测试都要重开 Access (~5s)，就 ~45 分钟，需 session-scoped driver。
5. **picker `acDialog` 阻塞**：`DoCmd.OpenForm name, , , , , acDialog` 是同步阻塞调用。stub form 必须立即关闭，否则测试卡死。

---

## 9. 前置条件 / 状态机测试

### 9.1 现状统计

`analysis/dump/precondition_graph.md` 自动抽取的结果：

- 全部 form 加起来共 **814 处 `.Enabled = True/False`** 调用
- LookAtNetworks 一个 form 就有 **152 处**
- 几乎每个 LookAt form 都遵循同一个流水线：
  ```
  Form_Open
    └─> 大部分按钮初始 Disabled
        └─> picker / import 触发后 → CmdQuery 才 Enabled
            └─> CmdQuery 成功后 → CmdGIS / CmdNeo4j / CmdStoreID / ... 才 Enabled
                └─> CmdImportXxx 加载后 → CmdRecallID 才 Enabled
  ```

### 9.2 三类前置条件

| 类型 | 含义 | 例子 |
|---|---|---|
| **Pre-Enabled** | 控件本身的 `Enabled` 属性必须为 True 才能点 | CmdGIS 在 `RecordCount = 0` 时 Disabled |
| **Pre-State** | 数据状态必须满足，否则点了会报错或无效 | CmdSaveEntryCodes 需要 `ZZ_SCRATCH_ENTRY_CODE` 非空 |
| **Pre-Sequence** | 必须按顺序完成上游步骤，否则下游用旧数据 | CmdGIS 必须在 CmdQuery 之后；中间不能改 picker 输入 |

### 9.3 测试模式 — 每个有前置条件的按钮 3 种 case

```python
# Case A: precondition NOT met → button should be Disabled OR raise
def test_export_disabled_before_query(form, export_btn):
    drv.open_form(form)
    # 不跑 CmdQuery
    assert drv.get_control_property(form, export_btn, "Enabled") == False
    # 强制触发应该报 VBA 错（用 driver bypass enable check）
    drv.invoke_event_unchecked(form, export_btn + "_Click")
    errs = drv.read_table("ZZ_TEST_ERRORS")
    assert len(errs) >= 1, "expected an error when triggering disabled button"

# Case B: precondition met (with empty result) → button should still be Disabled
def test_export_disabled_after_empty_query(form, export_btn):
    drv.open_form(form)
    # setup 一个返回空结果的 fixture（unmatchable filter）
    fixture_empty.populate_pickers(drv); fixture_empty.set_controls(drv)
    drv.invoke_event(form, "CmdQuery_Click")
    assert drv.row_count("ZZ_SCRATCH_<XXX>") == 0
    assert drv.get_control_property(form, export_btn, "Enabled") == False

# Case C: precondition met (with non-empty result) → button works
def test_export_works_after_query(form, export_btn, fixture):
    drv.open_form(form)
    fixture.populate_pickers(drv); fixture.set_controls(drv)
    drv.invoke_event(form, "CmdQuery_Click")
    assert drv.row_count("ZZ_SCRATCH_<XXX>") > 0
    assert drv.get_control_property(form, export_btn, "Enabled") == True
    drv.set_global("TestHelpers", "g_TestExportPath", str(tmp_out))
    drv.invoke_event(form, export_btn + "_Click")
    assert tmp_out.with_suffix(".tab").exists()
```

每个 export button 都需要这 3 个 case。10 form × 平均 6 export = 60 button × 3 case = **180 个 enable-state 测试**（在原 export 计划之外追加）。

### 9.4 Fixture Stack 抽象

为了避免每个测试都手动重复「先 open form → 再设 picker → 再设 control → 再跑 query」，引入 fixture 栈：

```python
# tests/cbdb_driver/fixture_stack.py

class Step:
    """A single state-mutating step."""
    name: str
    def apply(self, drv): ...
    def assert_post_state(self, drv): ...   # optional

class OpenForm(Step):
    def __init__(self, form): self.form = form
    def apply(self, drv): drv.open_form(self.form)

class SetPickerCodes(Step):
    def __init__(self, table, ids): self.table, self.ids = table, ids
    def apply(self, drv):
        drv.exec_sql(f"DELETE FROM [{self.table}]")
        for i in self.ids:
            drv.exec_sql(f"INSERT INTO [{self.table}] VALUES ({i})")

class SetControl(Step):
    def __init__(self, form, ctl, val):
        self.form, self.ctl, self.val = form, ctl, val
    def apply(self, drv): drv.set_control(self.form, self.ctl, self.val)

class InvokeEvent(Step):
    def __init__(self, form, event):
        self.form, self.event = form, event
    def apply(self, drv):
        drv.invoke_event(self.form, self.event)
        # auto-assert no VBA errors after every event invocation
        errs = drv.read_table("ZZ_TEST_ERRORS")
        assert len(errs) == 0, f"VBA error in {self.event}:\n{errs}"

class AssertEnabled(Step):
    def __init__(self, form, ctl, expected):
        self.form, self.ctl, self.expected = form, ctl, expected
    def apply(self, drv):
        actual = drv.get_control_property(self.form, self.ctl, "Enabled")
        assert actual == self.expected, \
            f"{self.ctl}.Enabled = {actual}, expected {self.expected}"


class Stack:
    def __init__(self, steps): self.steps = steps
    def run(self, drv):
        for s in self.steps:
            s.apply(drv)
```

测试声明变得可读：
```python
# 这个 fixture 在 LookAtEntry 上跑「凯封 yin general 900-1100」全流程
KAIFENG_YIN_FIXTURE = Stack([
    OpenForm("LookAtEntry"),
    AssertEnabled("LookAtEntry", "CmdGIS", False),       # 初始 disabled
    AssertEnabled("LookAtEntry", "CmdQuery", False),     # 初始 disabled (无 picker 输入)
    SetPickerCodes("ZZ_SCRATCH_ENTRY_CODE", [118]),
    SetPickerCodes("ZZ_SCRATCH_ADDR", [100658]),
    SetControl("LookAtEntry", "TxtFromYear", 900),
    SetControl("LookAtEntry", "TxtToYear", 1100),
    SetControl("LookAtEntry", "FrameYears", 2),
    AssertEnabled("LookAtEntry", "CmdQuery", True),      # 此时已可点
    InvokeEvent("LookAtEntry", "CmdQuery_Click"),
    AssertEnabled("LookAtEntry", "CmdGIS", True),        # query 后 enabled
    AssertEnabled("LookAtEntry", "CmdNeo4j", True),
    AssertEnabled("LookAtEntry", "CmdStoreID", True),
])

def test_kaifeng_yin_full_workflow(driver):
    KAIFENG_YIN_FIXTURE.run(driver)

def test_kaifeng_yin_export_to_gis(driver, tmp_path):
    # 把 fixture 当 prerequisite 跑掉
    KAIFENG_YIN_FIXTURE.run(driver)
    # 然后追加导出步骤
    Stack([
        SetGlobal("TestHelpers", "g_TestExportPath", str(tmp_path / "out")),
        InvokeEvent("LookAtEntry", "CmdGIS_Click"),
    ]).run(driver)
    assert (tmp_path / "out.tab").exists()
```

这种栈化写法让我们：
- 复用 setup 步骤，不必每次重写
- 每个 InvokeEvent 自动断言无 VBA 错
- 每个 AssertEnabled 验证 enable 状态符合精确预期
- pytest fixture 可以提供一个 "已查询完毕" 的 driver state，后续每个 export 测试基于它

### 9.5 Pre-Sequence 测试（顺序敏感）

部分按钮**先后顺序**会影响结果。例如：
- 跑 CmdQuery → 改 picker 输入 → 点 CmdGIS：导出的还是旧 query 结果（因为 ZZ_SCRATCH 没刷新）。这是「stale data」类 bug。
- 跑 CmdQuery → 点 CmdAllPlaces（清地点）→ 点 CmdGIS：CmdGIS 应该用 query 时的地点，还是当前 form 状态的地点？

测试方案：
```python
def test_export_uses_query_time_state(driver):
    """CmdGIS 输出的地点应是 CmdQuery 时的，不是当前控件的。"""
    Stack([
        OpenForm("LookAtEntry"),
        SetPickerCodes("ZZ_SCRATCH_ADDR", [100658]),  # Kaifeng
        SetPickerCodes("ZZ_SCRATCH_ENTRY_CODE", [118]),
        SetControl("LookAtEntry", "TxtFromYear", 900),
        SetControl("LookAtEntry", "TxtToYear", 1100),
        SetControl("LookAtEntry", "FrameYears", 2),
        InvokeEvent("LookAtEntry", "CmdQuery_Click"),
    ]).run(driver)
    n_kaifeng = driver.row_count("ZZ_SCRATCH_ENTRY")
    # 现在改 picker (不重新 query)
    Stack([
        SetPickerCodes("ZZ_SCRATCH_ADDR", [100660]),  # 中牟
        InvokeEvent("LookAtEntry", "CmdAllPlaces_Click"),  # 清掉
        SetGlobal("TestHelpers", "g_TestExportPath", str(tmp / "out")),
        InvokeEvent("LookAtEntry", "CmdGIS_Click"),
    ]).run(driver)
    out = (tmp / "out.tab").read_text(encoding="utf-8")
    # GIS 文件应仍包含 Kaifeng 数据（即 ZZ_SCRATCH_ENTRY 时的地点）
    assert "Kaifeng" in out or n_kaifeng == out.count("\n") - 1  # header
```

每个 form 至少 2-3 个 sequence 测试 → 25 个新增测试。

### 9.6 自动从 precondition_graph.md 生成测试

为了避免手写 180 个 enable-state 测试，自动生成：

```python
# tests/test_enable_states_generated.py
import json, pytest
from pathlib import Path

GRAPH = json.loads(Path("analysis/dump/precondition_graph.json").read_text())

@pytest.mark.parametrize("form,handler,enables", [
    (form, handler, ctls)
    for form, handlers in GRAPH.items()
    for handler, ctls in handlers.items()
    if ctls.get("enables")
], ids=lambda x: f"{x[0]}.{x[1]}")
def test_handler_enables_targets(driver, form, handler, enables):
    """After running <handler>, the listed control(s) should become Enabled."""
    drv = driver
    Stack([OpenForm(form)]).run(drv)
    # 先确认初始 disabled
    for c in enables:
        assert drv.get_control_property(form, c, "Enabled") == False, \
            f"{form}.{c} unexpectedly Enabled at form open"
    # 跑 handler 的最小前置 fixture（picker/textbox 任意值）
    fixture = MIN_FIXTURE_FOR.get((form, handler))
    if fixture: fixture.apply(drv)
    drv.invoke_event(form, handler)
    for c in enables:
        assert drv.get_control_property(form, c, "Enabled") == True, \
            f"{form}.{c} should be Enabled after {handler}"
```

需要维护一个 `MIN_FIXTURE_FOR` 字典 — 每个 handler 的最小输入。这是一次性 ~30 行。

### 9.7 修订后的总测试规模

| 桶 | 原计划 | + 前置条件追加 | 总计 |
|---|---:|---:|---:|
| query | 50 | + 0 | 50 |
| picker | 70 | + 35 (enable 状态) | 105 |
| export | 150 | + 180 (3 case × 60 button) | 330 |
| state | 18 | + 0 | 18 |
| checkbox | 75 | + 0 | 75 |
| frame | 24 | + 24 (enable 状态) | 48 |
| display | 10 | + 0 | 10 |
| nav | 10 | + 0 | 10 |
| unknown | 26 | + 0 | 26 |
| **sequence** | — | + 25 | 25 |
| **enable-state generated** | — | + 100 | 100 |
| schema / saved view | 50 | + 0 | 50 |
| **总计** | ~520 | ~ +364 | **~880** |

约 880 个测试，按 300ms/测试，全套约 4-5 分钟（session-scoped driver）。

---

## 10. 参考代码骨架

`tests/cbdb_driver/access_app.py` 起手：

```python
from __future__ import annotations
import shutil, subprocess
from pathlib import Path
import win32com.client
import pyodbc

class AccessApp:
    def __init__(self, mdb_path: Path):
        self.mdb_path = mdb_path
        self._app = None
        self._conn = None

    def open(self):
        self._app = win32com.client.Dispatch("Access.Application")
        self._app.Visible = False
        self._app.OpenCurrentDatabase(str(self.mdb_path))
        # 修引用
        proj = self._app.VBE.VBProjects(1)
        for r in list(proj.References):
            if r.IsBroken:
                full = getattr(r, "FullPath", "") or ""
                proj.References.Remove(r)
                if "dao" in full.lower():
                    for c in (
                        r"C:\Program Files\Microsoft Office\root\Office16\ACEDAO.DLL",
                        r"C:\Program Files (x86)\Microsoft Office\root\Office16\ACEDAO.DLL",
                    ):
                        if Path(c).exists():
                            proj.References.AddFromFile(c); break
        # ODBC for table I/O
        cs = ("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
              f"DBQ={self.mdb_path};")
        self._conn = pyodbc.connect(cs, autocommit=True)
        return self

    def close(self):
        try: self._conn.close()
        except: pass
        try: self._app.CloseCurrentDatabase()
        except: pass
        try: self._app.Quit()
        except: pass
        subprocess.run(["taskkill", "/F", "/IM", "MSACCESS.EXE"],
                       capture_output=True)

    @property
    def app(self): return self._app
    @property
    def conn(self): return self._conn

    # --- precondition helpers ---
    def get_control_property(self, form, ctl, prop):
        return getattr(self._app.Forms(form).Controls(ctl), prop)

    def invoke_event(self, form, event):
        """Trigger a (now-Public) form-module Sub. Auto-checks ZZ_TEST_ERRORS."""
        self._app.Run(f"Form_{form}.{event}")
        cur = self._conn.cursor()
        cur.execute("SELECT COUNT(*) FROM ZZ_TEST_ERRORS")
        n = cur.fetchone()[0]
        cur.close()
        if n > 0:
            errs = self.read_table("ZZ_TEST_ERRORS")
            raise AssertionError(f"VBA error during {form}.{event}:\n{errs}")

    def invoke_event_unchecked(self, form, event):
        """Trigger handler regardless of Enabled state — for negative tests."""
        try:
            self._app.Run(f"Form_{form}.{event}")
        except Exception:
            pass  # we'll inspect ZZ_TEST_ERRORS separately
```

---

## 11. 实现顺序优先级（从用户痛点出发）

按用户反馈的痛点排序：

1. **「有些导出功能不 work」** → Phase 1 + Section 5.2 export tests（**最优先**）
2. **「导出栏位不对/丢数据」** → Section 3.3 + golden files 对比（**次优先**）
3. **「触发错误提示」** → ZZ_TEST_ERRORS 钩子（infra 必备，每个测试都要）
4. **「按钮要满足前提才能触发」** → Section 9 enable-state + sequence tests（**第三优先**）
5. 其他控件类型 → Phase 3+

按这个顺序，第 1 周内可以做出针对前 3 项痛点的可运行覆盖（约 200 个测试），后续 1 周补完剩余 680 个。


`tests/cbdb_driver/vba_inject.py` 起手（核心 Public 化部分）：

```python
import re
PRIVATE_SUB = re.compile(r"^(\s*)Private Sub (Cmd\w+_(Click|AfterUpdate|Change))",
                         re.MULTILINE)

def make_handlers_public(app):
    proj = app.VBE.VBProjects(1)
    for comp in proj.VBComponents:
        if not comp.Name.startswith("Form_"):
            continue
        cm = comp.CodeModule
        n = cm.CountOfLines
        if n == 0: continue
        body = cm.Lines(1, n)
        new_body = PRIVATE_SUB.sub(r"\1Public Sub \2", body)
        if new_body != body:
            cm.DeleteLines(1, n)
            cm.AddFromString(new_body)

def inject_test_helpers(app):
    proj = app.VBE.VBProjects(1)
    if "TestHelpers" in [c.Name for c in proj.VBComponents]:
        return
    comp = proj.VBComponents.Add(1)  # vbext_ct_StdModule
    comp.Name = "TestHelpers"
    comp.CodeModule.AddFromString(TEST_HELPERS_VBA)

TEST_HELPERS_VBA = """\
Option Compare Database
Option Explicit

Public g_TestExportPath As String
Public g_SuppressMsgBox As Boolean

Public Sub LogError(formName As String, evName As String, desc As String)
    Dim sql As String
    sql = "INSERT INTO ZZ_TEST_ERRORS (form_name, event_name, err_desc, ts) " & _
          "VALUES (""" & formName & """, """ & evName & """, """ & _
          Replace(desc, """", """""") & """, Now())"
    CurrentDb.Execute sql, dbFailOnError
End Sub

Public Sub Invoke(formName As String, eventName As String)
    Application.Run "Form_" & formName & "." & eventName
End Sub
"""
```

