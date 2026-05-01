# 查询结果验证维度全菜单

每一维都对应一类潜在 bug。打勾的是当前已实现，方框是待加。

## ✅ 已实现（in `test_vba_integrity.py`）

- ✅ **列结构** — INSERT/UPDATE 目标列都存在（catches schema drift）
- ✅ **数据类型合理** — c_personid / c_index_year / c_year 是 numeric
- ✅ **Source 数据保真** — c_name / c_name_chn / c_index_year / c_dy ≡ BIOG_MAIN（catches copy-paste corruption）
- ✅ **Entry 事件可追溯** — 每行 (personid, entry_code) 在 ENTRY_DATA 找得到
- ✅ **Backfill 正确性** — c_entry_desc ≡ ENTRY_CODES.c_entry_desc；c_addr_name ≡ ADDR_CODES.c_name
- ✅ **FK 完整性** — c_addr_id ∈ ADDR_CODES；c_entry_code ∈ ENTRY_CODES
- ✅ **Backfill 完整性** — non-NULL FK 必须有 non-NULL desc（catches silent JOIN failure）
- ✅ **信息无丢失** — 独立 SQL 算的 personid 集合 ≡ VBA 结果（catches 复杂 JOIN 漏行）
- ✅ **每人事件数对** — 每个 personid 的行数 ≡ ENTRY_DATA COUNT（catches dedup bug）

## ⬜ 可加（高 ROI）

### 输入约束的事后验证 — 抓「filter 没生效」类
- ⬜ **年限范围严守** — 每行 c_index_year ∈ [from_year, to_year]（如果 mode=index）
- ⬜ **朝代过滤生效** — 每行 c_dy = 选定朝代代码（如果 mode=dynasty）
- ⬜ **地址过滤生效** — 每行 c_addr_id ∈ 选定 addr 集合（含 sub-units 展开）
- ⬜ **入仕代码过滤生效** — 每行 c_entry_code ∈ picker 选定集合

### 跨行一致性 — 抓「数据不自洽」类
- ⬜ **同 personid 名字一致** — 所有行同 personid 的 c_name / c_name_chn 应相同
- ⬜ **同 addr_id 地名一致** — 所有行同 c_addr_id 的 c_addr_name / c_addr_chn 相同
- ⬜ **xy_count 自洽** — c_xy_count = COUNT GROUP BY (x_coord, y_coord)
- ⬜ **同 entry_code 描述一致** — 所有行同 c_entry_code 的 c_entry_desc 相同

### 数值合理性 — 抓「脏数据」类
- ⬜ **c_index_year 合理** — 应在 [-2000, 2026]（CBDB 涵盖范围）
- ⬜ **x_coord 合理** — 应在 [60, 145]（中国大致经度范围）
- ⬜ **y_coord 合理** — 应在 [15, 55]（中国大致纬度范围）
- ⬜ **c_year 合理** — 应在 [-2000, 2026]
- ⬜ **c_personid > 0**

### 编码 / 字符集 — 抓「乱码」类
- ⬜ **c_name 仅 ASCII / 拼音** — 不应含中文字符
- ⬜ **c_name_chn 是合法 CJK** — 字符在 CJK Unified Ideographs 范围
- ⬜ **无替换字符** — '?' / '�' / 'mojibake' 检测
- ⬜ **无过长字段** — c_name < 100、c_name_chn < 50 字符
- ⬜ **无 leading/trailing 空白**

### 已知文档值对照 — 抓「regression vs 已知答案」
- ⬜ **HelpFile 数字** — 凯封 yin general 900-1100 = 104 ± 漂移容差
- ⬜ **王安石 (1762) 入仕方法** — 已知是 yin（恩荫）
- ⬜ **苏轼 (7097) 入仕年** — 已知 1057 jinshi
- ⬜ **欧阳修 (4017) 朝代** — 北宋

### 跨 Form 一致性 — 抓「同人物在不同 form 结果不一致」
- ⬜ **LookAtEntry 中的人 ⊂ LookAtPlace 同一地点的人**
- ⬜ **LookAtKinship 出现的 kin_id ∈ BIOG_MAIN**
- ⬜ **LookAtNetworks 边端点 ⊂ LookAtEntry 候选**

### 边界 / 负面测试 — 抓「不该出结果时出了」
- ⬜ **不存在的 entry_code (-1)** → empty
- ⬜ **未来年份 (3000-3100)** → empty
- ⬜ **不存在的 addr_id (-1)** → empty
- ⬜ **空 picker** → empty
- ⬜ **超大范围 (year -2000-9999)** → 不崩溃 + 返回大量行

### 性能 / 大小约束 — 抓「O(n²) 退化」
- ⬜ **大 fixture 在 < N 秒内返回**
- ⬜ **结果不超过 X 行**（防止笛卡尔积）

### Sequence / 顺序保留 — 抓「无序化导致丢顺序信息」
- ⬜ **同 personid 多事件时 c_sequence 单调** —不一定，但可检查不全 NULL
- ⬜ **xy_count 在 group 内一致**

### 真实 export 端到端 — 抓「export 端数据丢失」
- ⬜ **VBA 真触发 CmdGIS_Click → 文件行数 = ZZ_SCRATCH_ENTRY 行数 + 1 (header)**
- ⬜ **导出文件每列对应数据库列**（按 column index 验）
- ⬜ **导出文件 NULL 占位符正确**

---

要哪几类？我建议：
1. **「输入约束事后验证」** —— 5 个测试，立刻能抓「filter 没生效」类 bug
2. **「跨行一致性」** —— 4 个测试，抓数据不自洽
3. **「数值合理性」** —— 5 个测试，抓脏数据
4. **「真实 export 端到端」** —— 3 个测试，最后补上，抓 export bug

总计约 17 个新测试 × ~12s/test = ~3 分钟。
