# CBDB_BJ_User.mdb 自动化回归测试

本测试套件用于在 `CBDB_BJ_User.mdb` 更新后，自动验证其查询逻辑仍产生
正确结果。

## 设计

`CBDB_BJ_User.mdb` 把用户的查询行为分散在三层：

1. **链接表**（linked tables）— `BIOG_MAIN`、`*_DATA`、`*_CODES` 等都
   是从 `CBDB_*_DATA.mdb` 链接过来的，包含真实的人物 / 事件数据。
2. **保存的查询**（`View_*`）— 18 个去规范化的连接视图，作为子表单
   `RecordSource` 使用。
3. **VBA 表单事件**（`Form_LookAt*.CmdQuery_Click`）— 10 个主查询表单
   各有几千行 VBA，根据用户输入动态拼 SQL，把结果写入 `ZZ_SCRATCH_*`
   工作表。

针对这三层，本套件分三类测试：

| 测试文件 | 覆盖层 | 数量 |
|---|---|---|
| `test_schema.py` | 表 / 列存在性 | 9 |
| `test_saved_views.py` | 18 个 `View_*` 查询 | 25 |
| `test_lookatentry.py` | LookAtEntry 表单的查询逻辑（黄金范本） | 5 |
| `test_exports.py` | LookAtEntry GIS 导出格式（列集 / NULL / 数字） | 8 |
| `test_other_forms_skeletons.py` | 其余 9 个 LookAt 表单（待补全） | 9 + 9 skip |
| `test_known_bugs.py` | 已知 bug 回归 | 3 |
| `test_infra_smoke.py` | COM/pywinauto driver smoke (大部分 skip — 见 PHASE1_BREAKTHROUGH.md) | 2 + 1 skip |

**Auto-test 不能覆盖的部分**（需要 5 分钟手动 smoke）：见 [`MANUAL_SMOKE.md`](MANUAL_SMOKE.md)。

## 当前已知 bug（详见 `reports/CBDB_Issues_Report_EN.md`）

1. **`View_StatusData` 别名错位**：`c_fy_range_desc` / `c_fy_range_chn`
   被错误地从 `YEAR_RANGE_CODES_1`（last-year 那个 join）取值，
   首年范围实际显示的是末年范围值。
2. **DAO 3.6 引用断裂**：发布的 .mdb 引用了旧版 DAO
   (`C:\Program Files\Common Files\Microsoft Shared\DAO\dao360.dll`)，
   在新版 Office 上自动化打开表单时会报「Can't find project or library」。

## 运行

```bash
# 跑全部测试
python -m pytest tests/ -v

# 只跑某一组
python -m pytest tests/test_lookatentry.py -v
python -m pytest tests/test_saved_views.py -v
python -m pytest tests/test_schema.py -v

# 把 HelpFile / 文档相关的 print 信息显示出来
python -m pytest tests/ -v -s

# 抑制 pandas 关于 SQLAlchemy 的警告
python -m pytest tests/ -W ignore
```

## .mdb 更新工作流

更新 `CBDB_BJ_User.mdb` 或 `CBDB_*_DATA.mdb` 后：

```bash
# 1. 重新提取 mdb 元数据（表 / 列 / 查询 / VBA）
python analysis/dump_metadata.py
python analysis/dump_vba.py

# 2. 跑全套回归测试
python -m pytest tests/ -v

# 3. 如果有 golden CSV 失败，但你确认是数据真的改了
#    (而不是逻辑出错), 重新生成 golden:
python -m pytest tests/ --regenerate-goldens

# 4. 用 git diff tests/golden/ 检查每条 golden 的变化
#    确认无误后 commit
```

## 给其余 9 个 LookAt 表单补测试

⛔ **严禁翻译 VBA 成 Python。** `cbdb_replay/` 是历史遗留，不应扩展。

正确的新测试方法是通过 Access COM 驱动真实 VBA（`--include-vba`）：

1. 在 `tests/test_vba_matrix_all_forms.py` 中为目标 form 添加 fixture
2. 使用 `VbaSession` + `click_via_timer` 调用真实的 `CmdQuery_Click`
3. 从 `ZZ_SCRATCH_*` 结果表断言行数和列集合
4. 参考 `tests/test_vba_matrix.py` 作为完整范本

详见 `docs/skills/access-vba-probe.md` 中的 "pure-SQL-first principle"。

## 故障排查

- `Microsoft Access Driver` 未找到：装 [Access Database Engine 2016 Redistributable](https://www.microsoft.com/en-us/download/details.aspx?id=54920)（x64 的）。
- `Can't find project or library`：用 Access 打开 .mdb，`Alt+F11` 进 VBE，
  `Tools → References`，去掉打钩的 `MISSING:` 项，加回 `Microsoft Office
  16.0 Access Database Engine Object Library` 中的 ACEDAO.dll。
- `pandas only supports SQLAlchemy`：警告，无影响。`pytest -W ignore` 抑制。

## 文件树

```
tests/
├── README.md                          # 本文件
├── conftest.py                        # pytest fixtures
├── golden_helpers.py                  # CSV diff 工具
├── golden/                            # 黄金快照 CSV
│   └── lookatentry/
│       └── *.csv
├── cbdb_replay/
│   ├── __init__.py
│   ├── driver.py                      # ODBC 连接帮助
│   ├── lookatentry.py                 # CmdQuery_Click → Python 重写
│   └── TEMPLATE_lookat.py             # 其余 9 个表单的模板
├── test_schema.py
├── test_saved_views.py
├── test_lookatentry.py
├── test_other_forms_skeletons.py
└── test_known_bugs.py
```
