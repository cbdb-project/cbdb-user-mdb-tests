# CBDB 使用者版 .mdb — 問題彙報

_測試過程中發現的問題彙總，謹呈維護團隊斧正。_

尊敬的維護者：

下面是我們在為 CBDB 使用者版 .mdb 編寫自動化迴歸測試套件過程中，陸續整理出來的一些問題清單。我們希望這份報告能在您繼續主持這份寶貴資料集時有所幫助；同時，對您多年來在這套資料上的辛勤付出，我們由衷地表示感謝和敬意。

問題按嚴重程度排序（P0 最高）。每一條都包括：簡明描述、使用者端一步一步的復現步驟、（在介面上能看到時）相關截圖，以及一份建議的修復方案。這些問題並不緊急，整理在此只是為了方便您在合適的時候逐一處理。

## 覆蓋矩陣 —— 表單 × 按鈕測試結果

| Form | CmdQuery | CmdGIS | CmdNeo4j | CmdPajek | CmdGephi | CmdUCINet | CmdKML | CmdGUESS | CmdRun | CmdUTF8Pajek |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| LookAtEntry | ✗ FAIL | ✓ | ✓ | — | — | — | — | — | — | — |
| LookAtStatus | ✓ | ✓ | ✓ | ✗ FAIL | ~ SKIP | ✓ | — | — | — | — |
| LookAtTexts | ✓ | ✓ | ✓ | — | — | — | — | — | — | — |
| LookAtPlace | ✓ | ✗ FAIL | ✗ FAIL | ✓ | ✓ | — | — | — | — | — |
| LookAtAssociations | ✓ | ✓ | ✓ | ✗ FAIL | ✓ | ✗ FAIL | — | — | — | — |
| LookAtOffice | ✓ | ✗ FAIL | ✓ | — | — | — | — | ✓ | — | — |
| LookAtKinship | — | ✓ | ✓ | ✓ | — | ✗ FAIL | — | ✗ FAIL | ✓ | ✗ FAIL |
| LookAtNetworks | — | — | ~ SKIP | — | — | — | — | — | ~ SKIP | — |
| LookAtGroupData | — | ✓ | ✗ FAIL | — | — | — | — | — | ✗ FAIL | — |
| LookAtAssocPairs | ~ SKIP | ✓ | ✓ | ✓ | — | ✓ | — | — | — | — |

_PASS: 28 · FAIL: 12 · ERROR: 0 · SKIP: 4 · NOT RUN: 0 · N/A: 56_

## 目錄

- [P0 — 靜默資料錯誤](#p0--靜默資料錯誤)
  - [Issue #21 — LookAtOffice：CmdGIS 輸出的 IndexYear 欄幾乎為空（填充率 0.2%）——疑似靜默欄位繫結退化](#issue-21--lookatofficecmdgis-輸出的-indexyear-欄幾乎為空填充率-02疑似靜默欄位繫結退化)
  - [Issue #26 — User MDB 與 cbdb-online 快照的 c_index_addr_id 不一致率超過 0.5% 閾值](#issue-26--user-mdb-與-cbdb-online-快照的-c_index_addr_id-不一致率超過-05-閾值)
  - [Issue #23 — LookAtAssociations：CmdPajek 頂點區段數量錯誤——標頭宣告 501 個頂點，但實際匯出 8,093 列](#issue-23--lookatassociationscmdpajek-頂點區段數量錯誤標頭宣告-501-個頂點但實際匯出-8093-列)
  - [Issue #24 — LookAtKinship：CmdGUESS Gephi 輸出每個節點列的欄位數錯誤（nodedef 宣告 15 欄）](#issue-24--lookatkinshipcmdguess-gephi-輸出每個節點列的欄位數錯誤nodedef-宣告-15-欄)
- [P1 — 可見的執行時報錯](#p1--可見的執行時報錯)
  - [Issue #22 — LookAtAssociations / LookAtKinship：當 c_name 含有 CJK 漢字時，CmdUCINet 因「Invalid procedure call or argument」崩潰](#issue-22--lookatassociations--lookatkinship當-c_name-含有-cjk-漢字時cmducinet-因invalid-procedure-call-or-argument崩潰)
  - [Issue #6 — LookAtGroupData.queryEntry 崩潰「找不到欄位」——ENTRY_DATA.c_parental_status 缺少 _code 字尾](#issue-6--lookatgroupdataqueryentry-崩潰找不到欄位entry_datac_parental_status-缺少-_code-字尾)
  - [Issue #7 — LookAtPlace.CmdNeo4j：People-CSV 迴圈讀取未在 SELECT 中投影的 c_dynasty / c_dynasty_chn / c_female——觸發 JET 3265「找不到專案」](#issue-7--lookatplacecmdneo4jpeople-csv-迴圈讀取未在-select-中投影的-c_dynasty--c_dynasty_chn--c_female觸發-jet-3265找不到專案)
  - [Issue #25 — LookAtKinship / LookAtGroupData / LookAtAssociationPairs：CmdImport 往返失敗——ZZ_SCRATCH_IMPORT_PEOPLE 保持空白](#issue-25--lookatkinship--lookatgroupdata--lookatassociationpairscmdimport-往返失敗zz_scratch_import_people-保持空白)
- [P2 — 靜默顯示問題](#p2--靜默顯示問題)
  - [Issue #2 — LookAtGroupData：CmdRun 未從 BIOG_MAIN 回填 c_name](#issue-2--lookatgroupdatacmdrun-未從-biog_main-回填-c_name)
  - [Issue #3 — LookAtEntry：entry_code = 36（進士及第）時，c_entry_desc 回填全部為 NULL](#issue-3--lookatentryentry_code--36進士及第時c_entry_desc-回填全部為-null)
  - [Issue #10 — EVENT_ADDR_2 子表單：TxtAddrCHN / TxtAddrPY 繫結至 View_EventAddrData 中不存在的未別名欄位——顯示空白](#issue-10--event_addr_2-子表單txtaddrchn--txtaddrpy-繫結至-view_eventaddrdata-中不存在的未別名欄位顯示空白)
- [P3 — 缺失介面](#p3--缺失介面)
  - [Issue #15 — LookAtPlace 缺少 CmdGIS 按鈕——處理程式存在但無 UI 控制項](#issue-15--lookatplace-缺少-cmdgis-按鈕處理程式存在但無-ui-控制項)
- [P5 — 潛伏 / 不可達 / 當前無法復現](#p5--潛伏--不可達--當前無法復現)
  - [Issue #20 — BOM 字首地址名稱在 GIS 匯出中會產生嵌入的 TAB 分隔符——當前資料集靜止（BOM 資料已在上游清理，本 dump 中受影響列數為 0）](#issue-20--bom-字首地址名稱在-gis-匯出中會產生嵌入的-tab-分隔符當前資料集靜止bom-資料已在上游清理本-dump-中受影響列數為-0)
  - [Issue #1 — View_StatusData：c_fy_range_desc / c_fy_range_chn 引用錯誤的 YEAR_RANGE_CODES 別名——當前資料集靜止](#issue-1--view_statusdatac_fy_range_desc--c_fy_range_chn-引用錯誤的-year_range_codes-別名當前資料集靜止)
  - [Issue #4 — LookAtPlace.CmdGIS_Click 引用不存在的控制項 GISFrame——潛伏，被缺少按鈕（Issue #15）遮蔽](#issue-4--lookatplacecmdgis_click-引用不存在的控制項-gisframe潛伏被缺少按鈕issue-15遮蔽)
  - [Issue #5 — LookAtStatus.CmdPajek_Click 引用缺少的控制項 ChkIDs 及三個不存在的欄位——潛伏，被缺少按鈕（Issue #16）遮蔽](#issue-5--lookatstatuscmdpajek_click-引用缺少的控制項-chkids-及三個不存在的欄位潛伏被缺少按鈕issue-16遮蔽)
  - [Issue #9 — LookAtEntry.CmdNeo4j Institutions 區塊使用錯誤的記錄集變數 tRstAssocCodes——潛伏（當前無 ENTRY_DATA 列有 c_inst_code > 0）](#issue-9--lookatentrycmdneo4j-institutions-區塊使用錯誤的記錄集變數-trstassoccodes潛伏當前無-entry_data-列有-c_inst_code--0)
  - [Issue #11 — EVENTS_DATA_2 子表單：c_event_record_id 控制項繫結至不存在的欄位——已隱藏，因此為潛伏](#issue-11--events_data_2-子表單c_event_record_id-控制項繫結至不存在的欄位已隱藏因此為潛伏)
  - [Issue #12 — POSTED_TO_OFFICE_DATA_2 子表單：c_appt_type_code 控制項繫結至未投影的欄位——已隱藏，因此為潛伏](#issue-12--posted_to_office_data_2-子表單c_appt_type_code-控制項繫結至未投影的欄位已隱藏因此為潛伏)
- [嚴重等級說明](#嚴重等級說明)
- [附錄 A —— c_index_year / c_index_addr_id 與 cbdb-online-main-server 快照之間的偏差（差異需要逐筆分類後才能判定是否為缺陷）](#附錄-a--c_index_year--c_index_addr_id-與-cbdb-online-main-server-快照之間的偏差差異需要逐筆分類後才能判定是否為缺陷)
- [附錄 B —— TablesFields：文件表與實際資料庫結構對比](#附錄-b--tablesfields文件表與實際資料庫結構對比)
- [附錄 C —— ForeignKeys：文件表與實際資料庫結構對比](#附錄-c--foreignkeys文件表與實際資料庫結構對比)
- [結語](#結語)

## 嚴重等級說明

- P0 — 靜默資料錯誤：資料錯或缺失，但沒有任何報錯提示。
- P1 — 可見的執行時報錯：彈出錯誤對話方塊，操作中斷。
- P2 — 靜默顯示問題：表單欄位本應有資料，卻顯示為空。
- P3 — 缺失介面：程式碼裡實現了某功能，但介面上沒有按鈕去觸發它。
- P4 — 安裝設定：每臺新機器需要一次性處理。
- P5 — 潛伏 / 不可達 / 當前無法復現：保留作為歷史記錄；我們在當前 dump 上重新驗證過，無法再觸發症狀。

## P0 — 靜默資料錯誤

### Issue #21 — LookAtOffice：CmdGIS 輸出的 IndexYear 欄幾乎為空（填充率 0.2%）——疑似靜默欄位繫結退化

**涉及位置:** `Form_LookAtOffice.CmdGIS_Click`

**嚴重等級:** P0 — 靜默資料損毀：GIS 匯出看似成功，但 99.8% 列的 IndexYear 資料遺失。依賴年份篩選的下游 GIS 工作流程將靜默地收到空白年份。

#### 問題描述

以人物 80944（無篩選）執行 LookAtOffice CmdGIS 時，GIS 輸出檔案雖已產生，但 IndexYear 欄僅在 36,602 列中的 64 列有非空值（0.2%），遠低於正確 GIS 輸出預期的 80% 閾值。此模式與 Bug #10、#11、#12 記錄的靜默欄位繫結退化一致——CmdGIS SELECT 中的欄位名稱與 ZZ_SCRATCH 表格的實際 schema 不符。

由 test_cmd_gis_produces_file[office_80944_unfiltered] 偵測到，斷言 [LookAtOffice] CmdGIS 欄 IndexYear 僅 64/36602 列非空（0.2%），低於 80% 閾值。

#### 復現步驟

1. 以 Microsoft Access 開啟 CBDB_BJ_User.mdb。
2. 開啟 LookAtOffice 表單。
3. 將人物 ID 設為 80944，其餘篩選器留空。
4. 點選 CmdGIS。檔案產生時不會出現錯誤彈出視窗。
5. 開啟 GIS 輸出檔並檢查 IndexYear 欄：絕大多數列將為空白。

#### 建議修復方案

檢查 Form_LookAtOffice.CmdGIS_Click：找到填充 GIS 輸出 IndexYear 欄的 SELECT 語句，確認來源欄位名稱與實際 schema 一致（檢查 ZZ_SCRATCH_OFFICE 或對應資料表）。

### Issue #26 — User MDB 與 cbdb-online 快照的 c_index_addr_id 不一致率超過 0.5% 閾值

**涉及位置:** `BIOG_MAIN (c_index_addr_id)`

**嚴重等級:** P0 — 靜默資料漂移：約 25 名人物的主要地址 ID 與線上系統不同。使用 c_index_addr_id 的地理分析和 GIS 匯出將靜默地將這些人物置於錯誤位置。

#### 問題描述

對 User MDB 與 cbdb-online-main-server SQLite 快照之間的 BIOG_MAIN.c_index_addr_id 進行交叉核查，發現不一致率為 0.500%，恰好達到最大可接受閾值。以預設 5,000 列樣本計，約有 25 名人物在兩個系統中的 c_index_addr_id 不同，表明 User MDB 可能尚未完全套用近期上游地址指定，或快照領先於當前資料匯出。

由 test_index_year_addr_xcheck_sample 偵測到，斷言 c_index_addr_id 不一致率 0.500% 超過 0.5% 閾值。

#### 復現步驟

1. 執行：python reports/collect_index_year_diffs.py
2. 檢查 reports/index_drift_examples.json 中 bucket 為 'addr_only' 的列——這些是 User MDB 與線上快照之間 c_index_addr_id 不同的人物。
3. 對每個不一致的人物，查詢 BIOG_MAIN.c_index_addr_id 並與線上伺服器比較，確認哪個值為權威值。

#### 建議修復方案

將 cbdb-online 伺服器最新的 c_index_addr_id 指定套用至 User MDB 中的 BIOG_MAIN。不一致的列已列舉於 reports/index_drift_examples.json（bucket: 'addr_only'）中。

### Issue #23 — LookAtAssociations：CmdPajek 頂點區段數量錯誤——標頭宣告 501 個頂點，但實際匯出 8,093 列

**涉及位置:** `Form_LookAtAssociations.CmdPajek_Click`

**嚴重等級:** P0 — 靜默資料損毀：匯出的 Pajek 檔案在結構上無效。讀取該檔案的網路分析將在截斷的頂點集上執行，在沒有任何警告的情況下產生不正確的中心性/社群偵測結果。

#### 問題描述

Form_LookAtAssociations.CmdPajek_Click 產生的 Pajek .net 檔案在標頭宣告 '*Vertices 501'，但實際頂點區段在下一個 `*` 標記之前包含 8,093 列。Pajek 及其他依賴頂點計數標頭的網路分析工具將在 501 列後截斷頂點列表或觸發解析錯誤，靜默地丟棄其餘約 7,592 個頂點。

由 test_export_button_produces_file[LookAtAssociations_CmdPajek] 偵測到，斷言標頭宣告 501 個頂點但在下一個 `*` 區段前找到 8,093 列頂點。

#### 復現步驟

1. 以 Microsoft Access 開啟 CBDB_BJ_User.mdb。
2. 開啟 LookAtAssociations，選擇一個回傳大型關聯網路的查詢。
3. 點選 CmdPajek。.net 檔案寫入時不出現錯誤視窗。
4. 開啟 .net 檔案並計算 *Vertices 區段的列數：實際數量超過 '*Vertices N' 標頭中宣告的數字。

#### 建議修復方案

在 Form_LookAtAssociations.CmdPajek_Click 中，找到寫入 '*Vertices N' 標頭的位置，確保 N 來自實際寫入的頂點列數，而非預先計算的估計值或單獨的查詢結果。

### Issue #24 — LookAtKinship：CmdGUESS Gephi 輸出每個節點列的欄位數錯誤（nodedef 宣告 15 欄）

**涉及位置:** `Form_LookAtKinship.CmdGUESS_Click`

**嚴重等級:** P0 — 靜默資料損毀：Gephi 檔案在結構上無效。節點屬性靜默錯位，導致所有匯入的節點元資料不可靠。

#### 問題描述

Form_LookAtKinship.CmdGUESS_Click 產生的 Gephi .gdf 檔案在 nodedef 標頭宣告 15 個欄位，但實際節點資料列包含不同數量的欄位（欄位/值錯位）。Gephi 及下游工具將無法載入該檔案，或靜默地將節點屬性對應到錯誤的欄位。

由 test_cmd_guess_produces_file[kinship_person_3211] 偵測到，斷言 [LookAtKinship] Gephi: node rows with bad field count (nodedef has 15 cols)。

#### 復現步驟

1. 以 Microsoft Access 開啟 CBDB_BJ_User.mdb。
2. 開啟 LookAtKinship，設定一個可回傳親屬網路的人物 ID。
3. 點選 CmdGUESS。.gdf 檔案寫入時不出現錯誤。
4. 開啟 .gdf 檔案：計算 'nodedef>' 標頭宣告的欄數，與第一個節點資料列中逗號分隔值的數量進行比較。

#### 建議修復方案

在 Form_LookAtKinship.CmdGUESS_Click 中，確保 nodedef 標頭欄位列表與每列值列表來自同一個有序欄位定義。不符情況通常發生在某欄位被加入一個列表但未加入另一個時。

## P1 — 可見的執行時報錯

### Issue #22 — LookAtAssociations / LookAtKinship：當 c_name 含有 CJK 漢字時，CmdUCINet 因「Invalid procedure call or argument」崩潰

**涉及位置:** `Form_LookAtAssociations.CmdUCINet_Click / Form_LookAtKinship.CmdUCINet_Click`

**嚴重等級:** P1 — 可見的執行期崩潰：彈出視窗中止匯出。任何包含漢字姓名人物的關聯網路 UCINet 工作流程均會失敗。絕大多數 CBDB 人物具有漢字姓名，使其在實際的 LookAtAssociations / LookAtKinship → UCINet 使用中幾乎全面失敗。

#### 問題描述

Form_LookAtAssociations.CmdUCINet_Click 和 Form_LookAtKinship.CmdUCINet_Click 均以 2 引數形式呼叫 CreateTextFile（filename, overwrite），未傳入 Unicode 旗標。當輸出路徑或 c_name 值包含 CJK 漢字時，VBA 的 CreateTextFile 觸發「Invalid procedure call or argument」（執行期錯誤 5），因系統 ANSI 字碼頁無法編碼漢字。錯誤以彈出視窗形式出現並中止匯出。使用關聯程式碼 c_assoc_code = 437（'Presented literary composition as gift to' / '贈詩、文'）的韌體能可靠觸發此問題，因相關人物的 c_name 包含漢字。

由 test_bug22_associations_cmducinet_fires_invalid_procedure_call 及 test_bug22_kinship_cmducinet_sibling_form_fires_invalid_procedure_call 偵測到。

#### 復現步驟

1. 以 Microsoft Access 開啟 CBDB_BJ_User.mdb。
2. 開啟 LookAtAssociations 表單。
3. 選取關聯程式碼 c_assoc_code = 437（'Presented literary composition as gift to'）。
4. 點選 CmdUCINet。出現彈出視窗：'Invalid procedure call or argument'。UCINet 匯出檔案未建立。
5. 在 LookAtKinship.CmdUCINet 中，當親屬網路包含 c_name 含 CJK 漢字的人物時，同樣會出現此錯誤。

#### 建議修復方案

將 Form_LookAtAssociations.CmdUCINet_Click 和 Form_LookAtKinship.CmdUCINet_Click 中的 CreateTextFile 呼叫改為 3 引數形式：CreateTextFile(filename, True, True)——第三個引數啟用 Unicode 輸出。以 c_name 含 CJK 漢字的人物韌體進行測試。

### Issue #6 — LookAtGroupData.queryEntry 崩潰「找不到欄位」——ENTRY_DATA.c_parental_status 缺少 _code 字尾

**涉及位置:** `Form_LookAtGroupData.queryEntry`

**嚴重等級:** P1 — 常見操作路徑上的明顯崩潰：任何在 LookAtGroupData 中勾選 Entry 核取方塊的使用者都會遇到此錯誤。ZZ_SCRATCH_ENTRY 保持 0 列，後續匯出步驟（GIS、Neo4j 等）無法取得任何 Entry 資料。

#### 問題描述

Form_LookAtGroupData.vb 第 ~2621 行的 INSERT INTO，其目標欄位列表末尾為 c_parental_status_code，但 SELECT 投影末尾卻是 ENTRY_DATA.c_parental_status（缺少 _code 字尾）。ENTRY_DATA 上的實際欄位為 c_parental_status_code，因此 SQL 在使用者勾選 Entry 核取方塊並點選 Run 時立即崩潰，觸發 JET 錯誤 3061。ZZ_SCRATCH_ENTRY 保持 0 列。Form_LookAtEntry.vb 中完全相同的查詢使用了正確名稱；這是兩個表單之間的單字元漂移。

由 test_bug6_lookat_groupdata_query_entry_fires_no_such_field 及靜態原始碼斷言偵測到。

#### 復現步驟

1. 以 Microsoft Access 開啟 CBDB_BJ_User.mdb。
2. 開啟 **LookAtGroupData** 表單。
3. 在匯入人員清單中，輸入 person ID **1**（安惇 An Dun；有 2 筆 ENTRY_DATA）。
4. 僅勾選 **Entry** 核取方塊；其餘 Status / Office / Text / Addr 保持未勾選。
5. 點選 **Run**。
6. 彈出 JET 錯誤 3061 彈窗（「未提供一個或多個必要引數的值」或「找不到欄位」）。ZZ_SCRATCH_ENTRY 保持空白。

#### 建議修復方案

將 Form_LookAtGroupData.vb 第 ~2621 行的 ENTRY_DATA.c_parental_status 改為 ENTRY_DATA.c_parental_status_code。一字元修復；Form_LookAtEntry.vb 中已有正確寫法可參照。

### Issue #7 — LookAtPlace.CmdNeo4j：People-CSV 迴圈讀取未在 SELECT 中投影的 c_dynasty / c_dynasty_chn / c_female——觸發 JET 3265「找不到專案」

**涉及位置:** `Form_LookAtPlace.CmdNeo4j_Click`

**嚴重等級:** P1 — 正常使用者操作中的明顯崩潰。任何非空的 LookAtPlace → Neo4j 匯出都會確定性地觸發此問題。儘管 SaveAs 對話方塊已觸發，仍產生 0 個 CSV 檔案。

#### 問題描述

Form_LookAtPlace.CmdNeo4j_Click 的 People-CSV 區段透過僅投影四個 ZZ_SCRATCH_P_TEXT 欄位的 SELECT 開啟記錄集，但列寫入迴圈從該記錄集讀取 !c_dynasty、!c_dynasty_chn 與 !c_female。DAO 的 Recordset.Fields 集合僅包含 SELECT 投影的欄位；JOIN 只將 DYNASTIES 和 BIOG_MAIN 帶入範圍作為篩選，並不暴露其欄位。JET 在第一次 !c_dynasty 讀取時觸發 3265「找不到專案」，錯誤處理器在任何檔案寫入前退出 Sub，使用者看到彈窗且匯出產生 0 個 CSV 檔案。

由 test_bug7_lookat_place_cmdneo4j_fires_item_not_found 及靜態原始碼斷言偵測到。

#### 復現步驟

1. 以 Microsoft Access 開啟 CBDB_BJ_User.mdb。
2. 開啟 **LookAtPlace** 表單。
3. 使用地址選擇器選取一個有充足資料的地址（例如 c_addr_id = 100658，開封）。點選 **Run Query**。
4. 點選 **Neo4j** 匯出按鈕並選擇儲存位置。
5. 彈出執行時期錯誤 3265「找不到專案」。所選資料夾中沒有任何 Neo4j 匯出檔案。

#### 建議修復方案

將 Form_LookAtPlace.CmdNeo4j_Click（第 ~643-647 行）的 SELECT 子句擴充套件，加入 DYNASTIES.c_dynasty、DYNASTIES.c_dynasty_chn 及 BIOG_MAIN.c_female 的投影。FROM / JOIN 結構已將這些來源表帶入範圍；在 SELECT 中新增三個欄位引用即為完整修復。

### Issue #25 — LookAtKinship / LookAtGroupData / LookAtAssociationPairs：CmdImport 往返失敗——ZZ_SCRATCH_IMPORT_PEOPLE 保持空白

**涉及位置:** `Form_LookAtKinship.CmdImport_Click / Form_LookAtGroupData.CmdImport_Click / Form_LookAtAssociationPairs.CmdImportList_Click`

**嚴重等級:** P1 — 靜默匯入失敗：匯入看似成功完成，但目標資料表為空。依賴匯入人物列表的後續查詢或匯出將在沒有警告的情況下操作空資料集。

#### 問題描述

在填入人物 ID [1, 2, 3] 並點選 CmdImport（或 LookAtAssociationPairs 的 CmdImportList）後，handler 應將填入的 ID 寫入 ZZ_SCRATCH_IMPORT_PEOPLE。在全部三個表單中，匯入完成後目標表格仍為空（c_person_id = []）。不顯示錯誤彈出視窗——匯入看似成功，但未寫入任何資料。

由 test_cmd_import_round_trip[LookAtKinship.CmdImport]、[LookAtGroupData.CmdImport]、[LookAtAssociationPairs.CmdImportList] 偵測到，均斷言 ZZ_SCRATCH_IMPORT_PEOPLE.c_person_id = []；預期 [1, 2, 3]。

#### 復現步驟

1. 以 Microsoft Access 開啟 CBDB_BJ_User.mdb。
2. 開啟 LookAtKinship（或 LookAtGroupData / LookAtAssociationPairs）。
3. 在匯入欄位中輸入人物 ID 1, 2, 3。
4. 點選 CmdImport。不出現錯誤彈出視窗。
5. 查詢 ZZ_SCRATCH_IMPORT_PEOPLE：SELECT c_person_id FROM ZZ_SCRATCH_IMPORT_PEOPLE——資料表為空。

#### 建議修復方案

檢查各受影響表單的 CmdImport_Click / CmdImportList_Click handler。驗證 INSERT INTO ZZ_SCRATCH_IMPORT_PEOPLE 語句：確認來源控制項（文字方塊或列表方塊）被正確讀取，且 INSERT 在已提交的活躍交易中執行。

## P2 — 靜默顯示問題

### Issue #2 — LookAtGroupData：CmdRun 未從 BIOG_MAIN 回填 c_name

**涉及位置:** `Form_LookAtGroupData.CmdRun_Click`

**嚴重等級:** P2 — 靜默顯示問題：CmdRun 完成時沒有任何錯誤訊息，但結果中的 c_name 欄位為空白。使用者無從得知回填已失敗。

#### 問題描述

當使用者在 LookAtGroupData 中填入一個 person ID 並點選 CmdRun 時，handler 應執行 UPDATE 查詢，將 ZZ_SCRATCH_IMPORT_PEOPLE JOIN BIOG_MAIN，並為每一筆填入 c_name（及 c_dynasty）。在此版本中，UPDATE 未成功執行：CmdRun 完成後，ZZ_SCRATCH_IMPORT_PEOPLE 的 c_name 仍為 NULL。

結果是群組資料匯入畫面顯示空白的姓名欄位，且使用者不會看到任何錯誤訊息，CmdRun 靜默地失敗了。

由 test_hard_form_query_small_fixture[groupdata_person_1_small] 偵測到：斷言 'CmdRun didn't backfill c_name for c_person_id=1'，CmdRun 完成後 c_name 仍為 None。

#### 復現步驟

1. 以 Microsoft Access 開啟 CBDB_BJ_User.mdb。
2. 從導覽窗格開啟 **LookAtGroupData** 表單。
3. 在匯入人員清單中，輸入一個有效的 person ID（例如 **1**）。
4. 點選 **Run**（CmdRun 按鈕）。
5. CmdRun 完成後，檢視結果：姓名欄為空白。
6. SQL 驗證：`SELECT c_person_id, c_name FROM ZZ_SCRATCH_IMPORT_PEOPLE` 回傳 (1, NULL)——c_name 未從 BIOG_MAIN 回填。

#### 建議修復方案

在 Form_LookAtGroupData.CmdRun_Click 中找到將 ZZ_SCRATCH_IMPORT_PEOPLE JOIN BIOG_MAIN 並設定 c_name 的 UPDATE 語句，確認 JOIN 條件使用了正確的主鍵欄位，且 UPDATE 目標欄位名稱拼寫正確。修復後，以任意有效 person ID 執行 CmdRun，c_name 應能在 ZZ_SCRATCH_IMPORT_PEOPLE 中被填入。

### Issue #3 — LookAtEntry：entry_code = 36（進士及第）時，c_entry_desc 回填全部為 NULL

**涉及位置:** `Form_LookAtEntry.CmdQuery_Click`

**嚴重等級:** P2 — 靜默顯示問題：92,545 筆受影響。使用者可在結果格中看到空白的 c_entry_desc 欄，但 Access 不顯示錯誤——容易被忽略。參照此欄的匯出（GIS、Neo4j、KML）也會包含空白值。

#### 問題描述

當使用者在 LookAtEntry 以 entry_code = 36（進士及第）執行查詢時，結果表 ZZ_SCRATCH_ENTRY 雖然產生了 92,545 筆資料，但 c_entry_desc 欄位對每一筆都是 NULL。預期值應為 'examination: jinshi (general)'。

CmdQuery_Click 成功地從 ENTRY_DATA JOIN ENTRY_CODES 插入了資料，但 c_entry_desc 的回填步驟對此 entry code 並未寫入說明文字。其他欄位看起來都正常填充。因此，使用者在螢幕上看到的查詢結果中，每一筆記錄的入仕方式欄位都是空白，難以判斷是何種考試型別。

由 test_vba_full_matrix[top_entry_code_36_unfiltered] 等偵測到，斷言 'c_entry_desc backfill wrong'，影響 92,545 筆。

#### 復現步驟

1. 以 Microsoft Access 開啟 CBDB_BJ_User.mdb。
2. 從導覽窗格開啟 **LookAtEntry** 表單。
3. 在 Entry Code 選擇器中，選取 entry code **36**（標籤：'examination: jinshi (general)'）。
4. 朝代、地址、年份篩選器留空。
5. 點選 **Run Query**（CmdQuery 按鈕）。
6. 查詢完成後，檢視結果格：每一筆記錄的入仕方式說明欄（c_entry_desc）皆為空白。
7. SQL 驗證：`SELECT TOP 5 c_entry_code, c_entry_desc FROM ZZ_SCRATCH_ENTRY` 對所有列回傳 (36, NULL)。

#### 建議修復方案

在 Form_LookAtEntry.CmdQuery_Click 中找到對 ZZ_SCRATCH_ENTRY 設定 c_entry_desc 的回填步驟，確認 JOIN ENTRY_CODES 的條件（c_entry_code = 36）沒有被意外篩除，且 UPDATE / 回填 SQL 使用了正確的欄位名稱。修復後，`SELECT c_entry_desc FROM ZZ_SCRATCH_ENTRY WHERE c_entry_code = 36 LIMIT 1` 應回傳 'examination: jinshi (general)'。

### Issue #10 — EVENT_ADDR_2 子表單：TxtAddrCHN / TxtAddrPY 繫結至 View_EventAddrData 中不存在的未別名欄位——顯示空白

**涉及位置:** `EVENT_ADDR_2 Subform`

**嚴重等級:** P2 — 靜默顯示：EVENT_ADDR_2 子表單的兩個地址控制項對每一列均顯示空白，且不彈出任何錯誤訊息。使用者無從得知地址名稱有資料；父列的地址顯示不受影響。

#### 問題描述

EVENT_ADDR_2 子表單的 TxtAddrCHN 控制項 ControlSource 為 c_name_chn，TxtAddrPY 為 c_name，但表單的 RecordSource 是 View_EventAddrData，該查詢將 ADDR_CODES.c_name_chn 別名為 c_event_addr_chn，將 ADDR_CODES.c_name 別名為 c_event_addr_name。由於投影中不包含 c_name 與 c_name_chn，兩個控制項在每一列的「事件含地址」子資料表中均靜默地顯示空白。SQL 探測確認：SELECT c_name_chn FROM View_EventAddrData 丟擲「引數太少，預期 2 個」——JET 將未知識別字視為引數。

由 test_subform_control_source_unresolved[bug10_TxtAddrCHN] 偵測到。

#### 復現步驟

1. 以 Microsoft Access 開啟 CBDB_BJ_User.mdb。
2. 開啟 CBDB_Browser_2 並導航至 c_personid = 44872（孫才 Sun Cai）——此人有 EVENT_ADDR 列指向 c_addr_id = 12603（安豐）。
3. 切換至 **Events** 子分頁。
4. 觀察事件列中巢狀的 EVENT_ADDR_2 子表單：TxtAddrCHN 與 TxtAddrPY 顯示空白，即使父列的地址控制項（繫結至 View_EventsData 中的 c_addr_chn / c_addr_name）顯示「安豐」。

#### 建議修復方案

在 EVENT_ADDR_2 Subform 的表單設計師中，將 TxtAddrCHN.ControlSource 從 c_name_chn 改為 c_event_addr_chn，將 TxtAddrPY.ControlSource 從 c_name 改為 c_event_addr_name——即 View_EventAddrData 中的實際別名。

## P3 — 缺失介面

### Issue #15 — LookAtPlace 缺少 CmdGIS 按鈕——處理程式存在但無 UI 控制項

**涉及位置:** `LookAtPlace`

**嚴重等級:** P3 — 缺少 UI：即使底層處理程式功能完整（套用 Issue #4 修復後），GIS 匯出功能對 LookAtPlace 使用者完全不可用。

#### 問題描述

Form_LookAtPlace.vb 定義了功能完整的 CmdGIS_Click 處理程式——它建立並寫入與 Status / Texts / Associations / Office / Kinship 上 GIS 按鈕形狀相同的GIS .tab 匯出。但 LookAtPlace 的表單設計沒有 CmdGIS 按鈕。Place 的使用者可以使用 Pajek / Gephi / Neo4j 匯出，但無法使用 GIS 匯出；處理程式存在，只是無法從 UI 觸及。注意：若新增按鈕，Issue #4（同一處理程式中 GISFrame 與 CodeFrame 的錯字）必須同時修復。

由 test_orphan_export_button_truly_missing[bug15_LookAtPlace_CmdGIS] 偵測到。

#### 復現步驟

1. 以 Microsoft Access 開啟 CBDB_BJ_User.mdb。
2. 開啟 **LookAtPlace** 表單。
3. 檢視右下角的匯出按鈕列。沒有 GIS 按鈕。
4. 與 LookAtStatus / LookAtAssociations / LookAtOffice 等比較——這些表單均有 GIS 按鈕。

#### 建議修復方案

在 LookAtPlace 的表單設計中，在現有 CmdPajek / CmdGephi 按鈕旁新增 CmdGIS 按鈕，並設定 OnClick = [事件程式]。同時在同一補丁中修復 Issue #4 （GISFrame → CodeFrame 錯字）。

## P5 — 潛伏 / 不可達 / 當前無法復現

_本層的條目作為歷史 / 潛伏記錄保留。可分為三類：(a) DORMANT 潛伏 — 已驗證當前源資料無法觸發該症狀；(b) 當前無法復現 — 症狀不再出現，但可疑程式碼仍在（我們**沒有**確認上游有原始碼層面的修復；原因可能是 JET / Office 的行為改變、可能是我們這邊 fixture/driver 改變，也可能原本的診斷就是 false positive）；(c) LATENT 被遮蔽 — 原始碼缺陷確實存在，但因為另一個 issue（例如某個 UI 按鈕缺失）擋住了使用路徑，使用者目前碰不到。本層條目當下都不是使用者會遇到的問題，**也沒有任何一條被確認上游修復**；若要當成緊急或已關閉處理，請先諮詢。_

### Issue #20 — BOM 字首地址名稱在 GIS 匯出中會產生嵌入的 TAB 分隔符——當前資料集靜止（BOM 資料已在上游清理，本 dump 中受影響列數為 0）

**涉及位置:** `ADDR_CODES + Form_LookAt*.CmdGIS_Click`

**嚴重等級:** P5 — 在此 dump 中為靜止（BOM 資料已在上游清理，data-20260602 中受影響列數為 0）。若未來任何 ADDR_CODES 列在 c_name 或 c_name_chn 重新引入 U+FEFF 字首，將重新啟動並提升為 P0 靜默資料損毀。未轉義的 CmdGIS 寫入程式碼是尚未解決的第二部分修復。

#### 問題描述

ADDR_CODES 表先前有一些 c_name 和 c_name_chn 帶有前導 U+FEFF（BOM）字首的列（幾乎可確定是資料匯入時以帶 BOM 的 UTF-8 貼上所致），經 JET 處理後會在 CmdGIS 輸出中產生嵌入的 TAB 字元。最主要的可達範例是 c_addr_id = 702559（尉氏 Wei Shi），可從 c_status_code = 40（[為官者：文] / civil office）的人物到達。

在此版本（data-20260602）中，BOM 資料已在上游完成清理：test_addr_codes_has_known_bom_dirty_rows 現在找到 0 個帶 BOM 字首的列，test_known_reachable_dirty_addr_present 也回傳 0 列。CmdGIS 中未轉義的寫入模式仍存在於程式碼中（LookAtTexts / LookAtPlace / LookAtAssociations / LookAtOffice / LookAtKinship 的 CmdGIS 仍未進行 TAB 轉義），因此若未來資料引入新的 BOM 列，結構性風險將再次啟動。

GOLDEN_STALE：BOM golden 測試現在預期 0 列；請更新 goldens。此 Issue 因未轉義寫入的程式碼缺陷仍然存在，故以 P5_dormant_or_latent 保留。

#### 復現步驟

1. 在此版本中，Bug 無法觸發——SELECT COUNT(*) FROM ADDR_CODES WHERE Left(c_name, 1) = ChrW(65279) 回傳 0。
2. 結構性風險：開啟 Form_LookAtOffice（或任意 LookAt 表單），以狀態碼 **40**（[為官者：文] / civil office）執行 CmdGIS。若 c_addr_id = 702559（尉氏 Wei Shi）存在 BOM 列，輸出檔案在第 11476 列附近會出現多餘的 TAB 欄。
3. （靜止驗證）確認：SELECT COUNT(*) FROM ADDR_CODES WHERE Left(c_name, 1) = ChrW(65279) 在此 dump 中回傳 0。

#### 建議修復方案

資料端修復已在上游套用（0 個 BOM 列殘留）。程式碼端修復仍需進行：在所有 LookAt 表單 CmdGIS 主體的每個 tStr = tStr + value + tC 追加之前，將 value 中嵌入的 Chr(9)、Chr(10)、Chr(13) 或 U+FEFF 替換為空格，以防未來匯入帶 BOM 字首列時再次發生。

### Issue #1 — View_StatusData：c_fy_range_desc / c_fy_range_chn 引用錯誤的 YEAR_RANGE_CODES 別名——當前資料集靜止

**涉及位置:** `View_StatusData`

**嚴重等級:** P5 — 在當前資料集為靜止（若有任何 STATUS_DATA 列同時填入不同的 c_fy_range 與 c_ly_range，即提升為 P2 靜默顯示）。SQL 別名錯誤為確認的原始碼層級缺陷；當前只是缺少觸發列。

#### 問題描述

已儲存查詢 View_StatusData 將 YEAR_RANGE_CODES 聯結兩次，將第二個副本別名為 YEAR_RANGE_CODES_1 並以 STATUS_DATA.c_ly_range（末年範圍）聯結。但 SELECT 子句從 YEAR_RANGE_CODES_1 中取出 c_fy_range_desc 與 c_fy_range_chn——使用了錯誤的別名——導致每筆狀態記錄在「首年範圍」欄顯示的實為末年範圍文字。目前資料集中，沒有任何 STATUS_DATA 列同時填入不同的 c_fy_range 與 c_ly_range，因此症狀在 UI 上目前不可見，但一旦未來資料中出現此類列即會浮現。

由 test_bug_view_statusdata_fy_alias_swap 及 test_bug_view_statusdata_fy_value_equals_ly_value 偵測到。

#### 復現步驟

1. 以 Microsoft Access 開啟 CBDB_BJ_User.mdb。
2. 按 F11 顯示導覽窗格，然後雙擊查詢 **View_StatusData**。
3. 檢視 SELECT 子句：c_fy_range_desc 與 c_fy_range_chn 皆引用 YEAR_RANGE_CODES_1，但 FROM 子句是以 STATUS_DATA.c_ly_range（而非 c_fy_range）聯結該別名。
4. （靜止驗證）執行：SELECT c_personid, c_fy_range_desc, c_ly_range_desc FROM View_StatusData WHERE c_fy_range > 0 OR c_ly_range > 0。在當前資料集中結果為空，確認此 Bug 目前為靜止狀態。

#### 建議修復方案

在 View_StatusData 中，將 YEAR_RANGE_CODES_1.c_range AS c_fy_range_desc 及 YEAR_RANGE_CODES_1.c_range_chn AS c_fy_range_chn 改為引用未別名的 YEAR_RANGE_CODES（FROM 子句已以 STATUS_DATA.c_fy_range 聯結該副本）。每欄一行修復。

### Issue #4 — LookAtPlace.CmdGIS_Click 引用不存在的控制項 GISFrame——潛伏，被缺少按鈕（Issue #15）遮蔽

**涉及位置:** `Form_LookAtPlace.CmdGIS_Click`

**嚴重等級:** P5 — 潛伏（若 Issue #15 被修復但本行未一併修正，則提升為 P1 明顯崩潰）。測試驅動程式透過 GISFrame→CodeFrame 替換補丁使整合測試透過；CBDB 原始碼缺陷仍存在。

#### 問題描述

Form_LookAtPlace.CmdGIS_Click 在第 ~1539 行讀取 GISFrame.Value，但 LookAtPlace 沒有名為 GISFrame 的控制項——實際的編碼選擇器名為 CodeFrame。若在未修正此行的情況下新增了按鈕（修復 Issue #15），每次點選都會丟擲執行時期錯誤 424「需要物件」，GIS 匯出將無法執行。目前此 Bug 因表單上不存在 CmdGIS 按鈕（Issue #15）而被遮蔽。

由 test_bug4_lookat_place_cmdgis_fires_object_required 及 test_bug4_lookat_place_cmdgis_references_nonexistent_gisframe 偵測到。

#### 復現步驟

1. （假設情境，需先修復 Issue #15。）開啟 LookAtPlace。
2. 執行任意查詢使暫存表有資料。
3. 點選 GIS 按鈕。
4. 彈出執行時期錯誤 424「需要物件」；匯出不產生任何檔案。

#### 建議修復方案

將 Form_LookAtPlace.vb 第 ~1539 行的 GISFrame.Value 改為 CodeFrame.Value。同一表單的 CmdNeo4j_Click、CmdGephi_Click、CmdPajek_Click 已正確使用 CodeFrame ——這是單一識別字的漂移。應與 Issue #15（新增 CmdGIS 按鈕）在同一補丁中修復。

### Issue #5 — LookAtStatus.CmdPajek_Click 引用缺少的控制項 ChkIDs 及三個不存在的欄位——潛伏，被缺少按鈕（Issue #16）遮蔽

**涉及位置:** `Form_LookAtStatus.CmdPajek_Click`

**嚴重等級:** P5 — 潛伏（若 Issue #16 被修復但本缺陷未一併修正，則提升為 P1 明顯崩潰）。整個 Sub 看起來是從 LookAtAssociations.CmdPajek_Click 複製而未更新欄位名稱。

#### 問題描述

Form_LookAtStatus.CmdPajek_Click 包含兩個從 LookAtAssociations 複製但未更新名稱的相關缺陷：(a) 第 ~2308 行讀取 ChkIDs.Value，但 LookAtStatus 沒有 ChkIDs 控制項；(b) CmdPajek_Click 內的 SELECT 引用 ZZ_SCRATCH_STATUS.c_person_id、c_status_id 及 c_status_count——這些欄位均不存在（實際欄位為 c_personid、c_status_code，且無計數欄）。由於 LookAtStatus 沒有 CmdPajek 按鈕（Issue #16），兩個缺陷目前均無法觸發，但新增按鈕時若未一併修復，將使兩個錯誤暴露給使用者。

由 test_bug5_lookat_status_cmdpajek_sql_fires_field_error 及靜態原始碼斷言偵測到。

#### 復現步驟

1. （假設情境，需先修復 Issue #16。）開啟 LookAtStatus。
2. 執行任意查詢使 ZZ_SCRATCH_STATUS 有資料。
3. 點選 Pajek 按鈕。
4. 首先：彈出執行時期錯誤 424「需要物件」（ChkIDs.Value）。
5. 若繞過前者：因 SELECT 引用 c_person_id / c_status_id / c_status_count，觸發「找不到欄位」錯誤。

#### 建議修復方案

需要兩項修復：(a) 將 ChkIDs.Value 替換為 False（或在 LookAtStatus 新增真正的 ChkIDs 控制項）；(b) 將 SELECT 改寫為使用 ZZ_SCRATCH_STATUS.c_personid 及 ZZ_SCRATCH_STATUS.c_status_code，並視情況刪除或重新計算 c_status_count。實際上整個 Sub 可能需要徹底改寫而非區域性修補。

### Issue #9 — LookAtEntry.CmdNeo4j Institutions 區塊使用錯誤的記錄集變數 tRstAssocCodes——潛伏（當前無 ENTRY_DATA 列有 c_inst_code > 0）

**涉及位置:** `Form_LookAtEntry.CmdNeo4j_Click`

**嚴重等級:** P5 — 潛伏的原始碼層級錯字（若未來任何 ENTRY_DATA 列有 c_inst_code > 0，則重新提升為 P1）。test_bug9_lookat_entry_cmdneo4j 在此版本中 PASSED（閘道關閉）。當前缺少 InstitutionCodes CSV 並非使用者可見的 Bug。

#### 問題描述

Form_LookAtEntry.vb 第 ~1415 行以 Set tRstInstitutions = CurrentDb.OpenRecordset(tQueryStr) 開啟機構記錄集，但第 ~1425 行卻寫 With tRstAssocCodes——引用了已在上游 AssocCodes 區塊中 Close 的記錄集。若執行，.MoveFirst 將觸發 DAO 3021「沒有當前記錄」。整個區塊位於第 ~1389 行的 If tRecDeleted > 0 Then 之內，其中 tRecDeleted 為 c_inst_code > 0 的 ENTRY_DATA 列數。當前資料集中，263,454 筆 ENTRY_DATA 中有 0 筆 c_inst_code > 0，因此閘道評估為 false，test_bug9_lookat_entry_cmdneo4j PASSED（閘道關閉，確認潛伏狀態）。

代表性韌體：c_entry_code = 36（進士 jinshi general）和 c_entry_code = 101（薦舉 recommendation）。兩者的 ENTRY_DATA 列 c_inst_code = 0，確認閘道在當前資料集中保持關閉。

依 MANIFEST 要求保留（test_report_code_labels_audit_clean）。

#### 復現步驟

1. 在當前資料集中，此 Bug 無法透過 UI 觸發——Form_LookAtEntry.vb 第 ~1389 行的 If tRecDeleted > 0 Then 對所有可能的 LookAtEntry 韌體均評估為 false（263,454 筆 ENTRY_DATA 中有 0 筆 c_inst_code > 0）。
2. 靜態驗證原始碼層級錯字：開啟 analysis/dump/vba/Form_LookAtEntry.vb 並檢視第 ~1415-1425 行。第 ~1415 行：Set tRstInstitutions = CurrentDb.OpenRecordset(tQueryStr)。第 ~1425 行：With tRstAssocCodes（應為：With tRstInstitutions）。
3. （可選）確認閘道條件：SELECT COUNT(*) FROM ENTRY_DATA WHERE c_inst_code > 0 返回 0。
4. 代表性韌體：c_entry_code = 36（進士 jinshi general）和 c_entry_code = 101（薦舉 recommendation）。兩者均確認 c_inst_code = 0。

#### 建議修復方案

將第 ~1425 行的 With tRstAssocCodes 改為 With tRstInstitutions。單一識別字修復；正確的變數就在上方幾行剛被開啟。

### Issue #11 — EVENTS_DATA_2 子表單：c_event_record_id 控制項繫結至不存在的欄位——已隱藏，因此為潛伏

**涉及位置:** `EVENTS_DATA_2 Subform`

**嚴重等級:** P5 — 潛伏（若控制項被設為 Visible=True 或加寬，則提升為 P2 靜默顯示）。僅為程式碼衛生問題；目前對使用者無可見影響。

#### 問題描述

EVENTS_DATA_2 子表單有一個名為 c_event_record_id 的控制項，其 ControlSource 同樣為 c_event_record_id。EVENTS_DATA 和 View_EventsData 均未投影該名稱的欄位。若控制項可見，將顯示空白。即時 COM 探測確認控制項 Visible=False，寬度 240 twips（約 4 mm）——一個隱藏的內部控制項，幾乎可確定是從未打算顯示的遺留聯結鍵欄位。實際使用者不會看到空白欄，因為控制項未顯示。

由 test_subform_control_source_unresolved[bug11_c_event_record_id] 偵測到。

#### 復現步驟

1. 驗證僅限靜態與 COM 探測——無 UI 症狀。
2. 靜態證據：SELECT c_event_record_id FROM View_EventsData 丟擲「引數太少，預期 1 個」——確認該欄位不在投影中。
3. 可見性證據：COM 探測確認 EVENTS_DATA_2 Subform 上的 c_event_record_id 控制項 Visible=False，寬度 240 twips。

#### 建議修復方案

刪除隱藏的 c_event_record_id 控制項，或將其 ControlSource 改為真實欄位（如 c_event_code），使其不帶有過時的繫結。任何一種修改對使用者均不可見；僅為程式碼衛生。

### Issue #12 — POSTED_TO_OFFICE_DATA_2 子表單：c_appt_type_code 控制項繫結至未投影的欄位——已隱藏，因此為潛伏

**涉及位置:** `POSTED_TO_OFFICE_DATA_2 Subform`

**嚴重等級:** P5 — 潛伏（若控制項被設為可見，則提升為 P2 靜默顯示）。表單上面向使用者的任職型別控制項工作正常。僅為程式碼衛生。

#### 問題描述

POSTED_TO_OFFICE_DATA_2 子表單有一個 c_appt_type_code 控制項，其 ControlSource 為 c_appt_type_code，但 View_PostingOfficeData 投影的是 c_appt_code（無 _type 中綴）——而非 c_appt_type_code。即時 COM 探測確認控制項 Visible=False，因此空白渲染目前對使用者不可見。同一表單上面向使用者的任職型別控制項工作正常。此為純粹的程式碼衛生問題。

由 test_subform_control_source_unresolved[bug12_c_appt_type_code] 偵測到。

#### 復現步驟

1. 驗證僅限靜態與 COM 探測——無 UI 症狀。
2. 靜態證據：在 control_inventory.json 中，POSTED_TO_OFFICE_DATA_2 Subform 有一個 control_source = 'c_appt_type_code' 的控制項，但 View_PostingOfficeData 投影的是 c_appt_code。
3. 可見性證據：COM 探測確認 c_appt_type_code 控制項 Visible=False。

#### 建議修復方案

刪除隱藏的 c_appt_type_code 控制項，或將其 ControlSource 改為 c_appt_code（View_PostingOfficeData 實際投影的欄位）。

## 附錄 A —— c_index_year / c_index_addr_id 與 cbdb-online-main-server 快照之間的偏差（差異需要逐筆分類後才能判定是否為缺陷）

我們把本 .mdb 的 BIOG_MAIN 與 cbdb-online-main-server 每週釋出的 SQLite 快照在 `c_index_year`、`c_index_addr_id` 兩個欄位上做比對，可以看到一小部分人物對不齊。

**兩邊是兩套獨立的實作。**SQLite 快照中的 `c_index_year` 是 cbdb-online-main-server 的 PHP `IndexYearRebuildService.php` 算出來的，`c_index_addr_id` 則是 `IndexAddressRebuildService.php` 算出來的（程式碼都在 <https://github.com/cbdb-project/cbdb-online-main-server>）；User MDB 上對應的這兩個User MDB 那一邊：`c_index_addr_id` 由前端 mdb 裡的 `Form_frmIndexAddr` VBA 重建；`c_index_year` 由連結表後端 `data/CBDB_<YYYYMMDD>_DATA.mdb` 裡 **37 條 `BM IY Rule …` 的 QueryDef** 重建，由 `frmBaseMaintenance` 驅動。兩邊演算法已抽取到 `analysis/dump_data/querydefs_index/*.sql`；form / module 驅動 VBA 仍需 Access SaveAsText 互動式提取。PHP **意圖**映象 VBA，但兩者是兩條獨立的程式路徑。每一行差異**可能**來自下列至少四個原因，光看差異本身分不出來：(1) 源資料快照漂移；(2) PHP 與 VBA 之間的演演算法 / 移植差異；(3) 優先序 / 平手規則不同；(4) null / 預設值處理不同。

**我們並沒有對目前看到的 ~575 / 657 246 筆差異做完整分類。**下方列舉的樣本（目前共 13 筆、3 種分桶，來自 `reports/index_drift_examples.json`）只是**示範**這些差異**長什麼樣**，並非統計上有代表性，是後續逐筆分類的起點，不是結論。

### 僅 c_index_year 不一致的樣例

**`c_personid = 3501` — 李孝稱 (Li Xiaocheng)**

| 欄位 | 本 .mdb (User MDB) | cbdb-online-main-server 快照 |
|---|---|---|
| `c_index_year` | 1018 | 1028 |
| `c_index_addr_id` | 100658 | 100658 |
| `c_birthyear` | 0 | 0 |
| `c_deathyear` | 0 | 0 |
| `c_index_year_type_code` | 19 | 1912 |
| `c_index_year_source_id` | 19149 | 3479 |

**`c_personid = 15971` — 郭世隆 (Guo Shilong)**

| 欄位 | 本 .mdb (User MDB) | cbdb-online-main-server 快照 |
|---|---|---|
| `c_index_year` | 960 |  |
| `c_index_addr_id` | 100658 | 100658 |
| `c_birthyear` | 0 | 0 |
| `c_deathyear` | 0 | 0 |
| `c_index_year_type_code` | 14 |  |
| `c_index_year_source_id` | 24426 |  |

**`c_personid = 16266` — 錢孟回 (Qian Menghui)**

| 欄位 | 本 .mdb (User MDB) | cbdb-online-main-server 快照 |
|---|---|---|
| `c_index_year` | 1004 | 992 |
| `c_index_addr_id` | 12723 | 12723 |
| `c_birthyear` | 0 | 0 |
| `c_deathyear` | 0 | 0 |
| `c_index_year_type_code` | 13 | 0512 |
| `c_index_year_source_id` | 700103 | 3035 |

**`c_personid = 16267` — 錢知雄 (Qian Zhixiong)**

| 欄位 | 本 .mdb (User MDB) | cbdb-online-main-server 快照 |
|---|---|---|
| `c_index_year` | 1034 | 1071 |
| `c_index_addr_id` | 12723 | 12723 |
| `c_birthyear` | 0 | 0 |
| `c_deathyear` | 0 | 0 |
| `c_index_year_type_code` | 1312 | 14 |
| `c_index_year_source_id` | 16266 | 16269 |

**`c_personid = 19771` — 李彭 (Li Peng)**

| 欄位 | 本 .mdb (User MDB) | cbdb-online-main-server 快照 |
|---|---|---|
| `c_index_year` | 1253 |  |
| `c_index_addr_id` | 100185 | 100185 |
| `c_birthyear` | 0 | 0 |
| `c_deathyear` | 0 | 0 |
| `c_index_year_type_code` | 11 |  |
| `c_index_year_source_id` | 40822 |  |

### 僅 c_index_addr_id 不一致的樣例

**`c_personid = 1` — 安惇 (An Dun)**

| 欄位 | 本 .mdb (User MDB) | cbdb-online-main-server 快照 |
|---|---|---|
| `c_index_year` | 1042 | 1042 |
| `c_index_addr_id` | 101117 |  |
| `c_birthyear` | 1042 | 1042 |
| `c_deathyear` | 1104 | 1104 |
| `c_index_year_type_code` | 01 | 01 |
| `c_index_year_source_id` |  |  |

**`c_personid = 470` — 金君卿 (Jin Junqing)**

| 欄位 | 本 .mdb (User MDB) | cbdb-online-main-server 快照 |
|---|---|---|
| `c_index_year` | 1012 | 1012 |
| `c_index_addr_id` | 12879 | 12785 |
| `c_birthyear` | 0 | 0 |
| `c_deathyear` | 0 | 0 |
| `c_index_year_type_code` | 05 | 05 |
| `c_index_year_source_id` |  |  |

**`c_personid = 481` — 周秩 (Zhou Zhi)**

| 欄位 | 本 .mdb (User MDB) | cbdb-online-main-server 快照 |
|---|---|---|
| `c_index_year` | 1043 | 1043 |
| `c_index_addr_id` | 100416 | 12785 |
| `c_birthyear` | 0 | 0 |
| `c_deathyear` | 0 | 0 |
| `c_index_year_type_code` | 05 | 05 |
| `c_index_year_source_id` |  |  |

**`c_personid = 485` — 周穜 (Zhou Tong)**

| 欄位 | 本 .mdb (User MDB) | cbdb-online-main-server 快照 |
|---|---|---|
| `c_index_year` | 1046 | 1046 |
| `c_index_addr_id` | 100416 | 12785 |
| `c_birthyear` | 0 | 0 |
| `c_deathyear` | 0 | 0 |
| `c_index_year_type_code` | 05 | 05 |
| `c_index_year_source_id` |  |  |

**`c_personid = 562` — 範沖 (Fan Chong)**

| 欄位 | 本 .mdb (User MDB) | cbdb-online-main-server 快照 |
|---|---|---|
| `c_index_year` | 1067 | 1067 |
| `c_index_addr_id` | 100658 | 13292 |
| `c_birthyear` | 1067 | 1067 |
| `c_deathyear` | 1141 | 1141 |
| `c_index_year_type_code` | 01 | 01 |
| `c_index_year_source_id` |  |  |

### 底層 SOURCE 資料本身不同（生年 / 卒年）的樣例

**`c_personid = 263` — 張穆之 (Zhang Muzhi)**

| 欄位 | 本 .mdb (User MDB) | cbdb-online-main-server 快照 |
|---|---|---|
| `c_index_year` | 1016 |  |
| `c_index_addr_id` |  |  |
| `c_birthyear` | 1016 | 0 |
| `c_deathyear` | 1079 | 0 |
| `c_index_year_type_code` | 01 |  |
| `c_index_year_source_id` |  |  |

**`c_personid = 1455` — 沈邈 (Shen Miao)**

| 欄位 | 本 .mdb (User MDB) | cbdb-online-main-server 快照 |
|---|---|---|
| `c_index_year` | 1001 | 1008 |
| `c_index_addr_id` | 12887 | 12887 |
| `c_birthyear` | 1001 | 0 |
| `c_deathyear` | 1047 | 0 |
| `c_index_year_type_code` | 01 | 05 |
| `c_index_year_source_id` |  |  |

**`c_personid = 19149` — 李孝基 (Li Xiaoji)**

| 欄位 | 本 .mdb (User MDB) | cbdb-online-main-server 快照 |
|---|---|---|
| `c_index_year` | 1016 | 1026 |
| `c_index_addr_id` | 100658 | 100658 |
| `c_birthyear` | 1016 | 0 |
| `c_deathyear` | 1076 | 0 |
| `c_index_year_type_code` | 01 | 11 |
| `c_index_year_source_id` |  | 41030 |

## 附錄 B —— TablesFields：文件表與實際資料庫結構對比

本節將 `CBDB_20260602_DATA.mdb` 中 `TablesFields` 表的內容與 `reports/collect_schema_diffs.py` 透過 Access DAO（TableDefs）重建的資料庫結構進行比對。若存在差異，表示文件表可能已過時。

TablesFields 共 875 筆。從資料庫重建：997 筆。

重建結果: [tables_fields_regen.csv](tables_fields_regen.csv)

### TablesFields 中有但實際資料庫中不存在的記錄（過時）

| AccessTblNm | AccessFldNm |
|---|---|
| ADMIN_CAT_CODE_TYPE_REL | c_admin_type_code |
| ADMIN_CAT_TYPES | c_admin_type_code |
| ADMIN_CAT_TYPES | c_admin_type_hz |
| ADMIN_CAT_TYPES | c_admin_type_trans |
| ENTRY_DATA | c_addr_id |
| ENTRY_DATA | c_posting_id |
| MERGED_PERSON_DATA | c_merged_to_personid |
| PersonIDSource | LineNum |
| PersonIDSource | SourceTable |
| TMP_ADDR_C | Max_c_belongs_first_year |

### 實際資料庫中有但 TablesFields 未記錄的欄位

| AccessTblNm | AccessFldNm | DataFormat | NULL_allowed |
|---|---|---|---|
| ADDRESSES | belongs1_ID | Long | True |
| ADDRESSES | belongs1_Name | Text | True |
| ADDRESSES | belongs2_ID | Long | True |
| ADDRESSES | belongs2_Name | Text | True |
| ADDRESSES | belongs3_ID | Long | True |
| ADDRESSES | belongs3_Name | Text | True |
| ADDRESSES | belongs4_ID | Long | True |
| ADDRESSES | belongs4_Name | Text | True |
| ADDRESSES | belongs5_ID | Long | True |
| ADDRESSES | belongs5_Name | Text | True |
| ADDRESSES | c_addr_cbd | Text | True |
| ADDRESSES | c_addr_id | Long | True |
| ADDRESSES | c_admin_type | Text | True |
| ADDRESSES | c_firstyear | Integer | True |
| ADDRESSES | c_lastyear | Integer | True |
| ADDRESSES | c_name | Text | True |
| ADDRESSES | c_name_chn | Text | True |
| ADDRESSES | x_coord | Double | True |
| ADDRESSES | y_coord | Double | True |
| ADMIN_CAT_CODE_TYPE_REL | c_admin_cat_type_code | Text | False |
| ADMIN_CAT_TYPES | c_admin_cat_type_code | Text | False |
| ADMIN_CAT_TYPES | c_admin_cat_type_hz | Text | False |
| ADMIN_CAT_TYPES | c_admin_cat_type_trans | Text | False |
| ASSOC_DATA | c_tertiary_type_notes | Text | True |
| BIOG_ADDR_DATA | c_delete | Integer | True |
| BIOG_MAIN | c_name_fixed | Text | True |
| CopyTables | NotProcessed | Yes/No | True |
| CopyTables | TableName | Text | False |
| CopyTablesDefault | ID | Long | True |
| CopyTablesDefault | TableName | Text | True |
| ENTRY_DATA | c_entry_addr_id | Long | True |
| ETHNICITY_TRIBE_CODES | c_sortorder | Integer | True |
| ForeignKeys | AccessFldNm | Text | True |
| ForeignKeys | AccessTblNm | Text | True |
| ForeignKeys | DataFormat | Text | True |
| ForeignKeys | FKName | Text | True |
| ForeignKeys | FKString | Text | True |
| ForeignKeys | ForeignKey | Text | True |
| ForeignKeys | ForeignKeyBaseField | Text | True |
| ForeignKeys | IndexOnField | Text | True |
| ForeignKeys | NULL_allowed | Yes/No | True |
| ForeignKeys | skip | Integer | True |
| FormLabels | c_english | Text | True |
| FormLabels | c_fanti | Text | True |
| FormLabels | c_form | Text | True |
| FormLabels | c_jianti | Text | True |
| FormLabels | c_label_id | Integer | True |
| MERGED_PERSON_DATA | c_merged_from_personid | Long | False |
| OFFICE_CODES_CONVERSION | c_office_chn | Text | True |
| OFFICE_CODES_CONVERSION | c_office_chn_backup | Text | True |
| OFFICE_CODES_CONVERSION | c_office_id | Long | True |
| OFFICE_CODES_CONVERSION | c_office_id_backup | Long | True |
| OFFICE_TYPE_TREE_backup | c_office_type_desc | Text | True |
| OFFICE_TYPE_TREE_backup | c_office_type_desc_chn | Text | True |
| OFFICE_TYPE_TREE_backup | c_office_type_node_id | Text | True |
| OFFICE_TYPE_TREE_backup | c_parent_id | Text | True |
| OFFICE_TYPE_TREE_backup | c_tts_node_id | Text | True |
| Paste Errors | c_bibl_cat_code | Long | True |
| Paste Errors | c_created_by | Text | True |
| Paste Errors | c_created_date | Date/Time | True |
| Paste Errors | c_extant | Long | True |
| Paste Errors | c_modified_by | Text | True |
| Paste Errors | c_modified_date | Date/Time | True |
| Paste Errors | c_notes | Memo | True |
| Paste Errors | c_pages | Text | True |
| Paste Errors | c_source | Long | True |
| Paste Errors | c_textid | Long | True |
| Paste Errors | c_text_country | Long | True |
| Paste Errors | c_text_dy | Long | True |
| Paste Errors | c_text_nh_code | Long | True |
| Paste Errors | c_text_nh_year | Long | True |
| Paste Errors | c_text_range_code | Long | True |
| Paste Errors | c_text_type_id | Text | True |
| Paste Errors | c_text_year | Long | True |
| Paste Errors | c_title | Text | True |
| Paste Errors | c_title_alt_chn | Text | True |
| Paste Errors | c_title_chn | Text | True |
| Paste Errors | c_title_trans | Text | True |
| Paste Errors | c_url_api | Text | True |
| Paste Errors | c_url_api_coda | Text | True |
| Paste Errors | c_url_homepage | Text | True |
| POSTED_TO_OFFICE_DATA | c_posting_id_old | Long | True |
| SOCIAL_INSTITUTION_ALTNAME_CODES | c_inst_altname_chn | Text | True |
| SOCIAL_INSTITUTION_ALTNAME_CODES | c_inst_altname_desc | Text | True |
| SOCIAL_INSTITUTION_ALTNAME_CODES | c_inst_altname_type | Integer | True |
| SOCIAL_INSTITUTION_ALTNAME_CODES | c_notes | Text | True |
| SOCIAL_INSTITUTION_ALTNAME_DATA | c_inst_altname_hz | Text | False |
| SOCIAL_INSTITUTION_ALTNAME_DATA | c_inst_altname_py | Text | True |
| SOCIAL_INSTITUTION_ALTNAME_DATA | c_inst_altname_type | Integer | False |
| SOCIAL_INSTITUTION_ALTNAME_DATA | c_inst_code | Integer | False |
| SOCIAL_INSTITUTION_ALTNAME_DATA | c_inst_name_code | Integer | False |
| SOCIAL_INSTITUTION_ALTNAME_DATA | c_notes | Memo | True |
| SOCIAL_INSTITUTION_ALTNAME_DATA | c_pages | Text | True |
| SOCIAL_INSTITUTION_ALTNAME_DATA | c_source | Long | True |
| SOCIAL_INSTITUTION_CODES | c_inst_end_dy | Integer | True |
| SOCIAL_INSTITUTION_CODES | c_inst_end_year | Integer | True |
| SOCIAL_INSTITUTION_CODES_CONVERSION | c_inst_code | Long | True |
| SOCIAL_INSTITUTION_CODES_CONVERSION | c_inst_code_new | Integer | True |
| SOCIAL_INSTITUTION_CODES_CONVERSION | c_inst_name_code | Integer | True |
| SOCIAL_INSTITUTION_CODES_CONVERSION | c_new_new_code | Long | True |
| STATUS_TYPES | c_status_type_parent_code | Text | True |
| TablesFields | AccessFldNm | Text | False |
| TablesFields | AccessTblNm | Text | False |
| TablesFields | DataFormat | Text | True |
| TablesFields | DumpFldNm | Text | True |
| TablesFields | DumpTblNm | Text | True |
| TablesFields | ForeignKey | Text | True |
| TablesFields | ForeignKeyBaseField | Text | True |
| TablesFields | IndexOnField | Text | True |
| TablesFields | NULL_allowed | Yes/No | True |
| TablesFields | RowNum | Long | True |
| TablesFieldsChanges | Change | Text | True |
| TablesFieldsChanges | ChangeDate | Text | True |
| TablesFieldsChanges | ChangeNotes | Text | True |
| TablesFieldsChanges | FieldName | Text | True |
| TablesFieldsChanges | TableName | Text | True |
| TEXT_BIBLCAT_CODES | c_text_cat_level | Text | True |
| TEXT_BIBLCAT_CODES | c_text_cat_parent_id | Text | True |
| TEXT_CODES | c_text_type_id | Text | True |
| TMP_ADDR_C | Min_c_belongs_first_year | Integer | True |
| TMP_ADDR_D | c_addr_cbd | Text | True |
| TMP_ADDR_E | c_addr_cbd | Text | True |
| TMP_DISTANCE_DATA | assoc_xcoord | Double | True |
| TMP_DISTANCE_DATA | assoc_ycoord | Double | True |
| TMP_DISTANCE_DATA | c_assoc_id | Long | False |
| TMP_DISTANCE_DATA | c_distance | Double | True |
| TMP_DISTANCE_DATA | c_personid | Long | False |
| TMP_DISTANCE_DATA | c_t_dist | Double | True |
| TMP_DISTANCE_DATA | x_coord | Double | True |
| TMP_DISTANCE_DATA | y_coord | Double | True |
| ZZZ_DY_DATA | c_dy | Integer | False |
| ZZZ_DY_DATA | c_personid | Long | False |

### 屬性不一致

完整清單：`reports/schema_diff_tables_fields_mismatches.csv`（143 筆）

## 附錄 C —— ForeignKeys：文件表與實際資料庫結構對比

本節涵蓋 `ForeignKeys` 表及其所記錄的外部索引鍵關係。

ForeignKeys 共 188 筆。從資料庫重建（透過 Access.Application DAO）：223 筆。

重建結果: [foreign_keys_regen.csv](foreign_keys_regen.csv)

### ForeignKeys 中有但實際資料庫中不存在的記錄（過時）

| AccessTblNm | AccessFldNm | ForeignKey | ForeignKeyBaseField |
|---|---|---|---|
| ADDR_BELONGS_DATA | c_source | TEXT_CODES | c_textid |
| assoc_data | c_assoc_day_gz | GANZHI_CODES | c_ganzhi_code |
| assoc_data | c_assoc_nh_code | nian_hao | c_nianhao_id |
| assoc_data | c_assoc_range | year_range_codes | c_range_code |
| assoc_data | c_inst_name_code | SOCIAL_INSTITUTION_NAME_CODES | c_inst_name_code |
| assoc_data | c_inst_name_code,c_inst_code | SOCIAL_INSTITUTION_CODES | c_inst_code |
| biog_addr_data | c_addr_id | ADDR_CODES | c_addr_id |
| biog_addr_data | c_fy_day_gz | GANZHI_CODES | c_ganzhi_code |
| biog_addr_data | c_fy_nh_code | nian_hao | c_nianhao_id |
| biog_addr_data | c_fy_range | year_range_codes | c_range_code |
| biog_addr_data | c_ly_day_gz | GANZHI_CODES | c_ganzhi_code |
| biog_addr_data | c_ly_nh_code | nian_hao | c_nianhao_id |
| biog_addr_data | c_ly_range | year_range_codes | c_range_code |
| biog_addr_data | c_personid | BIOG_MAIN | c_personid |
| biog_addr_data | c_source | TEXT_CODES | c_textid |
| BIOG_INST_DATA | c_inst_name_code | SOCIAL_INSTITUTION_NAME_CODES | c_inst_name_code |
| BIOG_INST_DATA | c_inst_name_code,c_inst_code | SOCIAL_INSTITUTION_CODES | c_inst_code |
| biog_main | c_death_age_range | year_range_codes | c_range_code |
| biog_main | c_index_year_source_id | BIOG_MAIN | c_personid |
| biog_main | c_index_year_type_code | INDEXYEAR_TYPE_CODES | c_index_year_type_code |
| ENTRY_DATA | c_entry_dy | DYNASTIES | c_dy |
| ENTRY_DATA | c_inst_name_code | SOCIAL_INSTITUTION_NAME_CODES | c_inst_name_code |
| ENTRY_DATA | c_inst_name_code,c_inst_code | SOCIAL_INSTITUTION_CODES | c_inst_code |
| EVENTS_ADDR | c_event_code | EVENT_CODES | c_event_code |
| EVENTS_ADDR | c_personid | BIOG_MAIN | c_personid |
| EVENTS_ADDR | c_personid,c_sequence,c_event_code | EVENTS_DATA | c_event_code |
| POSTED_TO_OFFICE_DATA | c_inst_name_code,c_inst_code | SOCIAL_INSTITUTION_CODES | c_inst_code |
| SOCIAL_INSTITUTION_ADDR | c_inst_name_code,c_inst_code | SOCIAL_INSTITUTION_CODES | c_inst_code |

### 實際資料庫中有但 ForeignKeys 未記錄的欄位

| AccessTblNm | AccessFldNm | ForeignKey | ForeignKeyBaseField |
|---|---|---|---|
| ADDRESSES | c_addr_id | ADDR_CODES | c_addr_id |
| ASSOC_CODES | c_assoc_pair | ASSOC_CODES | c_assoc_code |
| Assoc_data | c_assoc_fy_day_gz | GANZHI_CODES | c_ganzhi_code |
| Assoc_data | c_assoc_fy_nh_code | NIAN_HAO | c_nianhao_id |
| Assoc_data | c_assoc_fy_range | YEAR_RANGE_CODES | c_range_code |
| ASSOC_TYPES | c_assoc_type_parent_id | ASSOC_TYPES | c_assoc_type_code |
| ENTRY_DATA | c_entry_addr_id | ADDR_CODES | c_addr_id |
| EVENTS_DATA | c_event_code | EVENTS_ADDR | c_event_code |
| EVENTS_DATA | c_personid | EVENTS_ADDR | c_personid |
| EVENTS_DATA | c_sequence | EVENTS_ADDR | c_sequence |
| POSTED_TO_OFFICE_DATA | c_inst_code | SOCIAL_INSTITUTION_CODES | c_inst_code |
| POSTED_TO_OFFICE_DATA | c_inst_name_code | SOCIAL_INSTITUTION_CODES | c_inst_name_code |
| SOCIAL_INSTITUTION_ADDR | c_inst_code | SOCIAL_INSTITUTION_CODES | c_inst_code |
| SOCIAL_INSTITUTION_ADDR | c_inst_name_code | SOCIAL_INSTITUTION_CODES | c_inst_name_code |
| SOCIAL_INSTITUTION_CODES | c_inst_end_dy | DYNASTIES | c_dy |

## 結語

感謝您抽時間讀完這份報告。以上各條都不緊急，我們把它們集中整理在一起，只是希望方便您在合適的時候逐一處理。

如果對其中任何一條的描述或建議有疑問，歡迎隨時一同討論。本倉庫裡對應的迴歸測試，會在任何一個迴歸標記不再復現時自動從 PASS 翻成 FAIL —— 這是「請調查一下」的訊號，而不是「問題已修復」的自動確認（因為標記不再復現也可能是 fixture / driver 變了，或者是我們當初的分類有誤）。
