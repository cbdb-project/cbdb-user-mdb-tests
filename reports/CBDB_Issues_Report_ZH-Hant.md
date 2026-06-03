# CBDB 使用者版 .mdb — 問題彙報

_測試過程中發現的問題彙總，謹呈維護團隊斧正。_

尊敬的維護者：

下面是我們在為 CBDB 使用者版 .mdb 編寫自動化迴歸測試套件過程中，陸續整理出來的一些問題清單。我們希望這份報告能在您繼續主持這份寶貴資料集時有所幫助；同時，對您多年來在這套資料上的辛勤付出，我們由衷地表示感謝和敬意。

問題按嚴重程度排序（P0 最高）。每一條都包括：簡明描述、使用者端一步一步的復現步驟、（在介面上能看到時）相關截圖，以及一份建議的修復方案。這些問題並不緊急，整理在此只是為了方便您在合適的時候逐一處理。

## 目錄

- [P2 — 靜默顯示問題](#p2--靜默顯示問題)
  - [Issue #1 — LookAtEntry：entry_code = 36（進士及第）時，c_entry_desc 回填全部為 NULL](#issue-1--lookatentryentry_code--36進士及第時c_entry_desc-回填全部為-null)
  - [Issue #2 — LookAtGroupData：CmdRun 未從 BIOG_MAIN 回填 c_name](#issue-2--lookatgroupdatacmdrun-未從-biog_main-回填-c_name)
- [嚴重等級說明](#嚴重等級說明)
- [附錄 —— c_index_year / c_index_addr_id 與 cbdb-online-main-server 快照之間的偏差（差異需要逐筆分類後才能判定是否為缺陷）](#附錄--c_index_year--c_index_addr_id-與-cbdb-online-main-server-快照之間的偏差差異需要逐筆分類後才能判定是否為缺陷)
- [結語](#結語)

## 嚴重等級說明

- P0 — 靜默資料錯誤：資料錯或缺失，但沒有任何報錯提示。
- P1 — 可見的執行時報錯：彈出錯誤對話方塊，操作中斷。
- P2 — 靜默顯示問題：表單欄位本應有資料，卻顯示為空。
- P3 — 缺失介面：程式碼裡實現了某功能，但介面上沒有按鈕去觸發它。
- P4 — 安裝設定：每臺新機器需要一次性處理。
- P5 — 潛伏 / 不可達 / 當前無法復現：保留作為歷史記錄；我們在當前 dump 上重新驗證過，無法再觸發症狀。

## P2 — 靜默顯示問題

### Issue #1 — LookAtEntry：entry_code = 36（進士及第）時，c_entry_desc 回填全部為 NULL

**涉及位置:** `Form_LookAtEntry.CmdQuery_Click`

**嚴重等級:** P2 — 靜默顯示問題：92,545 筆受影響。使用者可在結果格中看到空白的 c_entry_desc 欄，但 Access 不顯示錯誤——容易被忽略。參照此欄的匯出（GIS、Neo4j、KML）也會包含空白值。

#### 問題描述

當使用者在 LookAtEntry 以 entry_code = 36（進士及第）執行查詢時，結果表 ZZ_SCRATCH_ENTRY 雖然產生了 92,545 筆資料，但 c_entry_desc 欄位對每一筆都是 NULL。預期值應為 'examination: jinshi (general)'。

CmdQuery_Click 成功地從 ENTRY_DATA JOIN ENTRY_CODES 插入了資料，但 c_entry_desc 的回填步驟對此 entry code 並未寫入說明文字。其他欄位看起來都正常填充。因此，使用者在螢幕上看到的查詢結果中，每一筆記錄的入仕方式欄位都是空白，難以判斷是何種考試型別。

由 test_vba_full_matrix[top_entry_code_36_unfiltered] 偵測到，斷言 'c_entry_desc backfill wrong'，影響 92,545 筆。

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

## 結語

感謝您抽時間讀完這份報告。以上各條都不緊急，我們把它們集中整理在一起，只是希望方便您在合適的時候逐一處理。

如果對其中任何一條的描述或建議有疑問，歡迎隨時一同討論。本倉庫裡對應的迴歸測試，會在任何一個迴歸標記不再復現時自動從 PASS 翻成 FAIL —— 這是「請調查一下」的訊號，而不是「問題已修復」的自動確認（因為標記不再復現也可能是 fixture / driver 變了，或者是我們當初的分類有誤）。
