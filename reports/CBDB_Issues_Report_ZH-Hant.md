# CBDB 使用者版 .mdb — 問題彙報

_測試過程中發現的問題彙總，謹呈維護團隊斧正。_

_資料構建：20260602_

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
| LookAtOffice | ✓ | ✓ | ✓ | — | — | — | — | ✓ | — | — |
| LookAtKinship | — | ✓ | ✓ | ✓ | — | ✗ FAIL | — | ✗ FAIL | ✓ | ✓ |
| LookAtNetworks | — | — | ~ SKIP | — | — | — | — | — | ~ SKIP | — |
| LookAtGroupData | — | ✓ | ✗ FAIL | — | — | — | — | — | ✗ FAIL | — |
| LookAtAssocPairs | ~ SKIP | ✓ | ✓ | ✓ | — | ✓ | — | — | — | — |

_PASS: 30 · FAIL: 10 · ERROR: 0 · SKIP: 4 · NOT RUN: 0 · N/A: 56_

## 目錄

- [P3 — 缺失介面](#p3--缺失介面)
  - [Issue #13 — BIOG_MAIN_2 子表單呼叫一個不存在的選取表單（frmPickNIAN_HAO）](#issue-13--biog_main_2-子表單呼叫一個不存在的選取表單frmpicknian_hao)
  - [Issue #16 — LookAtStatus 缺少 CmdPajek 按鈕（處理常式存在，但表單上沒有控制項）](#issue-16--lookatstatus-缺少-cmdpajek-按鈕處理常式存在但表單上沒有控制項)
  - [Issue #17 — LookAtStatus 缺少 CmdGephi 按鈕（處理常式存在，但表單上沒有控制項）](#issue-17--lookatstatus-缺少-cmdgephi-按鈕處理常式存在但表單上沒有控制項)
  - [Issue #18 — LookAtStatus 缺少 CmdUCINet 按鈕（處理常式存在，但表單上沒有控制項）](#issue-18--lookatstatus-缺少-cmducinet-按鈕處理常式存在但表單上沒有控制項)
  - [Issue #19 — LookAtOffice 缺少 CmdGUESS 按鈕（處理常式存在，但表單上沒有控制項）](#issue-19--lookatoffice-缺少-cmdguess-按鈕處理常式存在但表單上沒有控制項)
- [P4 — 安裝設定](#p4--安裝設定)
  - [Issue #2 — VBA 專案參照已過時的 dao360.dll，在 Office 2016 以後的機器上並不存在](#issue-2--vba-專案參照已過時的-dao360dll在-office-2016-以後的機器上並不存在)
- [P5 — 潛伏 / 不可達 / 當前無法復現](#p5--潛伏--不可達--當前無法復現)
  - [Issue #5 — LookAtStatus.CmdPajek 參照了一個不存在的控制項以及三個不存在的欄位 —— 潛伏（被缺少的 Pajek 按鈕擋住，見 Issue #16）](#issue-5--lookatstatuscmdpajek-參照了一個不存在的控制項以及三個不存在的欄位--潛伏被缺少的-pajek-按鈕擋住見-issue-16)
  - [Issue #6 — LookAtGroupData 的 Entry 插入投影了 ENTRY_DATA.c_parental_status（應為 …_code）—— 本次潛伏（執行期未觸發錯誤）](#issue-6--lookatgroupdata-的-entry-插入投影了-entry_datac_parental_status應為-_code-本次潛伏執行期未觸發錯誤)
  - [Issue #7 — LookAtPlace.CmdNeo4j 的人員記錄集讀取了 SELECT 未投影的 c_dynasty / c_dynasty_chn / c_female —— 潛伏（本次執行期未觸發）](#issue-7--lookatplacecmdneo4j-的人員記錄集讀取了-select-未投影的-c_dynasty--c_dynasty_chn--c_female--潛伏本次執行期未觸發)
  - [Issue #8 — LookAtNetworks.CmdNeo4j 的地點記錄集讀取了 SELECT 未投影的 x_coord / y_coord —— 潛伏（Networks 表單開啟卡死，行為重現受阻）](#issue-8--lookatnetworkscmdneo4j-的地點記錄集讀取了-select-未投影的-x_coord--y_coord--潛伏networks-表單開啟卡死行為重現受阻)
  - [Issue #9 — LookAtEntry.CmdNeo4j 的 Institutions 區塊用錯了記錄集變數（tRstAssocCodes）—— 潛伏（被閘門擋住而不可達；沒有任何 ENTRY_DATA 列的 c_inst_code > 0）](#issue-9--lookatentrycmdneo4j-的-institutions-區塊用錯了記錄集變數trstassoccodes-潛伏被閘門擋住而不可達沒有任何-entry_data-列的-c_inst_code--0)
  - [Issue #14 — KIN_DATA 子表單的 CmdPickKinRel 呼叫一個不存在的選取表單（frmPickKINSHIP_CODES）—— 潛伏（宿主子表單為孤兒，無可達觸發路徑）](#issue-14--kin_data-子表單的-cmdpickkinrel-呼叫一個不存在的選取表單frmpickkinship_codes-潛伏宿主子表單為孤兒無可達觸發路徑)
  - [Issue #20 — 帶 BOM 字首的地址名稱可能變成內嵌 TAB 而使 GIS 輸出錯位 —— 本次建置休眠（ADDR_CODES 中 0 列帶 BOM）](#issue-20--帶-bom-字首的地址名稱可能變成內嵌-tab-而使-gis-輸出錯位--本次建置休眠addr_codes-中-0-列帶-bom)
  - [Issue #22 — LookAtAssociations.CmdUCINet 的 CreateTextFile 缺少 Unicode 旗標 → 遇 CJK c_name 時 error 5 —— 潛伏（本次執行期未觸發）](#issue-22--lookatassociationscmducinet-的-createtextfile-缺少-unicode-旗標--遇-cjk-c_name-時-error-5--潛伏本次執行期未觸發)
  - [Issue #23 — LookAtAssociations.CmdPajek 的『*Vertices』表頭數值在 MoveLast 之前讀取 RecordCount（頂點數少算）—— 結構性度量，P5](#issue-23--lookatassociationscmdpajek-的vertices表頭數值在-movelast-之前讀取-recordcount頂點數少算-結構性度量p5)
  - [Issue #24 — LookAtKinship 的 GUESS/Gephi .gdf nodedef 宣告 15 欄，但部分節點列只寫出 13 格 —— 結構性度量，P5](#issue-24--lookatkinship-的-guessgephi-gdf-nodedef-宣告-15-欄但部分節點列只寫出-13-格--結構性度量p5)
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

## P3 — 缺失介面

### Issue #13 — BIOG_MAIN_2 子表單呼叫一個不存在的選取表單（frmPickNIAN_HAO）

**涉及位置:** `BIOG_MAIN_2_Subform`

**嚴重等級:** P3 —— 缺少 UI（點選要開啟的選取表單不存在，此功能無法使用）。

#### 問題描述

當使用者在人物詳細資料子表單上點選年號（NIAN_HAO）選取器時，`Form_BIOG_MAIN_2_Subform` 會執行 `DoCmd.OpenForm "frmPickNIAN_HAO"`（處理常式設定 `stDocName = "frmPickNIAN_HAO"`，並參照 `Forms!frmPickNIAN_HAO!frmNIAN_HAO.Form!c_nianhao_id`）。但目前的 .mdb 裡並沒有名為 `frmPickNIAN_HAO` 的表單（最新的 `control_inventory.json` 中查無此表單）。Access 會丟出『Item not found…』，這次點選對使用者毫無作用。

宿主表單 BIOG_MAIN_2_Subform 本身存在且可達（已在最新控制項清單中確認）——缺的只是它要開啟的選取表單。可能原因：早期重構時某個選取表單被改名或合併，而這個呼叫端沒有同步更新。

（本次測試為非互動式，無法重新擷取執行期彈窗；下方截圖為可達的宿主表單加上重建的彈窗，而選取表單在靜態層面的缺失才是關鍵證據。）

#### 復現步驟

1. 開啟 CBDB_Browser_2，導覽到任一在 BIOG_MAIN_2_Subform 上顯示人物詳細資料的人物。
2. 在子表單上點選年號（NIAN_HAO）選取器控制項——這會觸發執行 `DoCmd.OpenForm "frmPickNIAN_HAO"` 的處理常式。
3. 會跳出『Item not found in this collection.』彈窗，因為 `frmPickNIAN_HAO` 不在 CurrentProject.AllForms 中。
4. 靜態確認（不需 Access）：在 `analysis/dump/control_inventory.json` 中搜尋 `frmPickNIAN_HAO`——查無此表單，而 `BIOG_MAIN_2_Subform` 則存在。

#### 截圖

![bug13_browser_open.png](screenshots/bug13_browser_open.png)

_CBDB_Browser_2 open on a person record — the reachable host surface from which the NIAN_HAO picker is invoked._

![bug13_browser_annotated.png](screenshots/bug13_browser_annotated.png)

_Annotated host view: the reign-period picker control on BIOG_MAIN_2_Subform whose click runs DoCmd.OpenForm "frmPickNIAN_HAO" — a form absent from the current .mdb._

![bug13_faux_popup.png](screenshots/bug13_faux_popup.png)

_The 'Item not found in this collection.' popup, reconstructed in PIL (this build's session was non-interactive); the message is Access's standard text when DoCmd.OpenForm targets a form not in CurrentProject.AllForms._

#### 建議修復方案

兩種做法擇一：還原選取表單 `frmPickNIAN_HAO`，或將 `Form_BIOG_MAIN_2_Subform` 內的呼叫端改成開啟取代它的那個年號選取表單。

### Issue #16 — LookAtStatus 缺少 CmdPajek 按鈕（處理常式存在，但表單上沒有控制項）

**涉及位置:** `LookAtStatus`

**嚴重等級:** P3 —— 缺少 UI（使用者無法使用此功能）。

#### 問題描述

`Form_LookAtStatus.vb` 定義了 `CmdPajek_Click` 處理常式（用來輸出狀態網路的 Pajek `.net` 檔），但 LookAtStatus 的表單設計上並沒有 `CmdPajek` 控制項。最新的 `control_inventory.json` 顯示此表單有 CmdQuery / CmdGIS / CmdNeo4j，卻沒有 Pajek 按鈕，因此這個功能在 UI 上無法使用。

注意：即使加上按鈕，也必須先修正 Issue #5（CmdPajek_Click 內的 ChkIDs 控制項與 SQL 欄位缺陷），否則點選仍會失敗。

#### 復現步驟

1. 開啟 LookAtStatus。看底部的輸出按鈕列：只有 GIS 和 Neo4j，沒有 Pajek 按鈕。
2. 與 LookAtAssociations 比較，後者確實有 Pajek 按鈕。
3. 靜態確認：在 `analysis/dump/control_inventory.json` 中，LookAtStatus 沒有 `CmdPajek` 控制項，但 `Form_LookAtStatus.vb` 定義了 `Sub CmdPajek_Click()`。

#### 截圖

![bug16_LookAtStatus_no_CmdPajek.png](screenshots/bug16_LookAtStatus_no_CmdPajek.png)

_LookAtStatus as it ships — the export-button row has GIS and Neo4j but no Pajek button._

![bug16_LookAtStatus_no_CmdPajek_annotated.png](screenshots/bug16_LookAtStatus_no_CmdPajek_annotated.png)

_Annotated: the gap where a CmdPajek button would sit; `Sub CmdPajek_Click()` exists in the module but no control invokes it._

#### 建議修復方案

在 LookAtStatus 的設計中新增 CmdPajek 按鈕（OnClick = [事件程式]，以呼叫既有的 CmdPajek_Click）——但請先修正 Issue #5，否則點選會因 ChkIDs 參照與錯誤的 SQL 而失敗。

### Issue #17 — LookAtStatus 缺少 CmdGephi 按鈕（處理常式存在，但表單上沒有控制項）

**涉及位置:** `LookAtStatus`

**嚴重等級:** P3 —— 缺少 UI（使用者無法使用此功能）。

#### 問題描述

`Form_LookAtStatus.vb` 定義了 `CmdGephi_Click` 處理常式，但 LookAtStatus 的表單設計上並沒有 `CmdGephi` 控制項。最新的 `control_inventory.json` 確認表單上沒有 Gephi 按鈕，因此 Gephi 輸出在 UI 上無法使用。

#### 復現步驟

1. 開啟 LookAtStatus。輸出按鈕列裡沒有 Gephi 輸出按鈕。
2. 靜態確認：`analysis/dump/control_inventory.json` 顯示 LookAtStatus 上沒有 `CmdGephi` 控制項，但 `Form_LookAtStatus.vb` 定義了 `Sub CmdGephi_Click()`。

#### 截圖

![bug17_LookAtStatus_no_CmdGephi.png](screenshots/bug17_LookAtStatus_no_CmdGephi.png)

_LookAtStatus as it ships — no Gephi export button._

![bug17_LookAtStatus_no_CmdGephi_annotated.png](screenshots/bug17_LookAtStatus_no_CmdGephi_annotated.png)

_Annotated: `Sub CmdGephi_Click()` exists in the module but no control invokes it._

#### 建議修復方案

在 LookAtStatus 的設計中新增 CmdGephi 按鈕，並連到既有的 CmdGephi_Click。

### Issue #18 — LookAtStatus 缺少 CmdUCINet 按鈕（處理常式存在，但表單上沒有控制項）

**涉及位置:** `LookAtStatus`

**嚴重等級:** P3 —— 缺少 UI（使用者無法使用此功能）。

#### 問題描述

`Form_LookAtStatus.vb` 定義了 `CmdUCINet_Click` 處理常式，但 LookAtStatus 的表單設計上並沒有 `CmdUCINet` 控制項。最新的 `control_inventory.json` 確認表單上沒有 UCINet 按鈕，因此 UCINet 輸出在 UI 上無法使用。

#### 復現步驟

1. 開啟 LookAtStatus。輸出按鈕列裡沒有 UCINet 輸出按鈕。
2. 靜態確認：`analysis/dump/control_inventory.json` 顯示 LookAtStatus 上沒有 `CmdUCINet` 控制項，但 `Form_LookAtStatus.vb` 定義了 `Sub CmdUCINet_Click()`。

#### 截圖

![bug18_LookAtStatus_no_CmdUCINet.png](screenshots/bug18_LookAtStatus_no_CmdUCINet.png)

_LookAtStatus as it ships — no UCINet export button._

![bug18_LookAtStatus_no_CmdUCINet_annotated.png](screenshots/bug18_LookAtStatus_no_CmdUCINet_annotated.png)

_Annotated: `Sub CmdUCINet_Click()` exists in the module but no control invokes it._

#### 建議修復方案

在 LookAtStatus 的設計中新增 CmdUCINet 按鈕，並連到既有的 CmdUCINet_Click。

### Issue #19 — LookAtOffice 缺少 CmdGUESS 按鈕（處理常式存在，但表單上沒有控制項）

**涉及位置:** `LookAtOffice`

**嚴重等級:** P3 —— 缺少 UI（使用者無法使用此功能）。

#### 問題描述

`Form_LookAtOffice.vb` 定義了 `CmdGUESS_Click` 處理常式（用來輸出 GUESS `.gdf` 檔），但 LookAtOffice 的表單設計上並沒有 `CmdGUESS` 控制項。最新的 `control_inventory.json` 顯示此表單有 GIS / GISPeople / Neo4j，卻沒有 GUESS 按鈕，因此 GUESS 輸出在 UI 上無法使用。

#### 復現步驟

1. 開啟 LookAtOffice。沒有 GUESS 輸出按鈕（只有 GIS / GISPeople / Neo4j）。
2. 靜態確認：`analysis/dump/control_inventory.json` 顯示 LookAtOffice 上沒有 `CmdGUESS` 控制項，但 `Form_LookAtOffice.vb` 定義了 `Sub CmdGUESS_Click()`。

#### 截圖

![bug19_LookAtOffice_no_CmdGUESS.png](screenshots/bug19_LookAtOffice_no_CmdGUESS.png)

_LookAtOffice as it ships — no GUESS export button._

![bug19_LookAtOffice_no_CmdGUESS_annotated.png](screenshots/bug19_LookAtOffice_no_CmdGUESS_annotated.png)

_Annotated: `Sub CmdGUESS_Click()` exists in the module but no control invokes it._

#### 建議修復方案

在 LookAtOffice 的設計中新增 CmdGUESS 按鈕，並連到既有的 CmdGUESS_Click。

## P4 — 安裝設定

### Issue #2 — VBA 專案參照已過時的 dao360.dll，在 Office 2016 以後的機器上並不存在

**涉及位置:** `(VBA project)`

**嚴重等級:** P4 —— 每臺新機器一次性的安裝門檻。

#### 問題描述

出貨的 .mdb 其 VBA 專案內含一個硬編碼參照，指向 `C:\Program Files\Common Files\Microsoft Shared\DAO\dao360.dll`——這是 Access 2003 時代 DAO 3.6 的位置。現代 Office（2016 起）改為附帶 `ACEDAO.DLL`，並不會安裝這個舊版 DLL。在任何乾淨的現代機器上，第一次執行表單程式碼時就會跳出『Can't find project or library』，對一般使用者而言訊息晦澀又嚇人。

我們的測試驅動在開檔時會自動把這個壞掉的參照換成 ACEDAO.DLL（見 `analysis/check_vba_refs.py` 的修復前後參照傾印），所以回歸測試不會踩到；但一般使用者直接雙擊出貨的 .mdb 就會踩到。嚴重度低，因為每臺機器只需修一次，但每次全新安裝都會遇到。

#### 復現步驟

1. 在一臺全新的現代 Office 機器上安裝 `CBDB_BJ_User.mdb`。
2. 開檔後按 Alt+F11 進入 VBE。
3. 工具 → 設定參照——會看到一條標記為 `MISSING: dao360.dll` 的專案。
4. 開啟任一 LookAt 表單（或任何會執行表單程式碼的操作）。在表單程式碼執行前，就會跳出『Can't find project or library』編譯錯誤視窗。

#### 建議修復方案

在維護者的機器上做一次即可：用 Access 開啟 .mdb，按 Alt+F11，進入 工具 → 設定參照，取消勾選 MISSING 的 dao360.dll，改勾選 `Microsoft Office 16.0 Access Database Engine Object Library`（即 ACEDAO.DLL），存檔。之後重新散佈修好的檔案，後續使用者就不必再處理。

## P5 — 潛伏 / 不可達 / 當前無法復現

_本層的條目作為歷史 / 潛伏記錄保留。可分為三類：(a) DORMANT 潛伏 — 已驗證當前源資料無法觸發該症狀；(b) 當前無法復現 — 症狀不再出現，但可疑程式碼仍在（我們**沒有**確認上游有原始碼層面的修復；原因可能是 JET / Office 的行為改變、可能是我們這邊 fixture/driver 改變，也可能原本的診斷就是 false positive）；(c) LATENT 被遮蔽 — 原始碼缺陷確實存在，但因為另一個 issue（例如某個 UI 按鈕缺失）擋住了使用路徑，使用者目前碰不到。本層條目當下都不是使用者會遇到的問題，**也沒有任何一條被確認上游修復**；若要當成緊急或已關閉處理，請先諮詢。_

### Issue #5 — LookAtStatus.CmdPajek 參照了一個不存在的控制項以及三個不存在的欄位 —— 潛伏（被缺少的 Pajek 按鈕擋住，見 Issue #16）

**涉及位置:** `Form_LookAtStatus.CmdPajek_Click`

**嚴重等級:** P5 —— 潛伏原始碼缺陷（若只修 Issue #16 而不先修這裡，將以可見的當機重新浮現）。

#### 問題描述

同一處理常式中兩個相關的原始碼缺陷：

(a) 第 2308 行讀取 `If ChkIDs.Value Then`，但 LookAtStatus 並沒有名為 `ChkIDs` 的控制項。

(b) 第 2335-2338 行建立的 SELECT … INTO 參照了 `ZZ_SCRATCH_STATUS.c_person_id`、`c_status_id`、`c_status_count`——這三者在該表上都不存在（真正的欄位是 `c_personid`、`c_status_code`，且沒有計數欄位）。這個 Sub 看起來是 `LookAtAssociations.CmdPajek_Click` 的複製，那裡這些名稱是有效的；改名時漏了這兩處。

為何潛伏：LookAtStatus 根本沒有 Pajek 按鈕（Issue #16），所以使用者目前無法觸發這個處理常式。一旦此 Sub 執行，SQL 仍會立即失敗，因此若只加按鈕而不修這裡，只會把失敗暴露給使用者。本次測試為非互動式，無法重新驗證執行期 UI 症狀——故列為潛伏，待 UI 重新驗證。

#### 復現步驟

1. 本次建置無法透過 UI 觸發此 bug——LookAtStatus 沒有 Pajek 按鈕（Issue #16）。改以靜態方式驗證：
2. 開啟 `analysis/dump/vba/Form_LookAtStatus.vb`，看第 2308 行：`If ChkIDs.Value Then`——在 `analysis/dump/control_inventory.json` 中，LookAtStatus 並沒有 `ChkIDs` 控制項。
3. 看第 2335-2338 行：SELECT … INTO 參照了 `ZZ_SCRATCH_STATUS.c_person_id` / `c_status_id` / `c_status_count`（計數彙總在第 2337 行），這些都不是 ZZ_SCRATCH_STATUS 的欄位。

#### 建議修復方案

(a) 將 `ChkIDs.Value` 改為常數 `False`（若不需要可選的ID 字尾行為）或新增真正的 ChkIDs 控制項。(b) 重寫 SELECT 改用 `ZZ_SCRATCH_STATUS.c_personid` 與 `c_status_code`，並將計數彙總移除或以其他方式計算。實務上整個 Sub 需要謹慎重寫——它是未經驗證就從別的表單沿用過來的——並應與新增按鈕（Issue #16）一併處理。

### Issue #6 — LookAtGroupData 的 Entry 插入投影了 ENTRY_DATA.c_parental_status（應為 …_code）—— 本次潛伏（執行期未觸發錯誤）

**涉及位置:** `Form_LookAtGroupData.queryEntry`

**嚴重等級:** P5 —— 潛伏（原始碼筆誤存在；本次建置未重現執行期症狀，待 UI 重新驗證）。

#### 問題描述

`Form_LookAtGroupData.vb` 的 Entry INSERT 目標欄位列出 `c_parental_status_code`（第 2612 行），但 SELECT 投影結尾卻是 `ENTRY_DATA.c_parental_status`（第 2621 行）——少了 `_code` 字尾。ENTRY_DATA 真正的欄位是 `c_parental_status_code`；當 Entry 分支執行時，這個原始碼層級的筆誤會讓 JET 丟出『No such field』/『No value given for one or more required parameters』。`Form_LookAtEntry.vb` 對應的查詢用的是正確名稱，因此這是一行的漂移。

本次建置的誠實說明：原始碼缺陷確實存在於傾印中，但本次行為探測在執行時並未觸發該錯誤（症狀依資料／啟用路徑而定，且本次為非互動式，無法重新驗證執行期 UI 症狀）。故列為潛伏待 UI 重新驗證，而非已確認的使用者當機。

#### 復現步驟

1. 本次建置執行期未觸發此錯誤——以靜態方式驗證原始碼缺陷：
2. 開啟 `analysis/dump/vba/Form_LookAtGroupData.vb`。第 2612 行列出 INSERT 目標欄位 `c_parental_status_code`；第 2621 行投影 `ENTRY_DATA.c_parental_status`（少了 `_code`）。
3. 若要在未來的互動式測試中走到這條路徑：在 LookAtGroupData 匯入一位人物，只勾選 Entry，點 Run。若路徑被觸發，會跳出『欄位不存在』的彈窗。

#### 建議修復方案

把第 2621 行的 `ENTRY_DATA.c_parental_status` 改成 `ENTRY_DATA.c_parental_status_code`。一行修正，與 `Form_LookAtEntry.vb` 已使用的正確名稱一致。

### Issue #7 — LookAtPlace.CmdNeo4j 的人員記錄集讀取了 SELECT 未投影的 c_dynasty / c_dynasty_chn / c_female —— 潛伏（本次執行期未觸發）

**涉及位置:** `Form_LookAtPlace.CmdNeo4j_Click`

**嚴重等級:** P5 —— 潛伏（已確認的靜態投影不符；本次非互動式建置未重現執行期症狀，待 UI 重新驗證）。

#### 問題描述

`Form_LookAtPlace.CmdNeo4j_Click` 在一個只投影四個 ZZ_SCRATCH_P_TEXT 欄位的 SELECT DISTINCT（第 322 行：c_person_id、c_name、c_name_chn、c_index_year）上開啟 `tRstPeople`（第 326 行）。INNER JOIN 把 DYNASTIES 與 BIOG_MAIN 帶入範圍，但並未投影它們的任何欄位。接著逐列寫出的迴圈從該記錄集讀取 `!c_dynasty`（第 383 行）、`!c_dynasty_chn`（385）與 `!c_female`（392）；DAO 的 Fields 集合只含被投影的欄位，因此 JET 會在第一次這類讀取時丟出 3265 『Item not found in this collection.』。處理常式在任何磁碟檔寫出前就跳到結束，使用者會看到彈窗以及一個空的輸出資料夾。

本次為何潛伏：LookAtPlace 上確實有 CmdNeo4j 按鈕，但本次測試為非互動式（pywinauto UIA 無法使用），無法重現／重新驗證執行期症狀。投影不符是已確認的靜態缺陷；故列為潛伏待 UI 重新驗證。建議的示範地址為 `c_addr_id = 100658`（Kaifeng / 開封），其關聯人物足以餵滿 People-CSV 迴圈。

#### 復現步驟

1. 本次建置未重現執行期症狀（非互動式測試）。以靜態方式驗證投影不符：
2. 開啟 `analysis/dump/vba/Form_LookAtPlace.vb`。第 322 行只把 c_person_id / c_name / c_name_chn / c_index_year 投影到 `tRstPeople`（第 326 行開啟）。
3. 第 383 / 385 / 392 行從該記錄集讀取 `!c_dynasty`、`!c_dynasty_chn`、`!c_female`——皆未被投影，因此第一次讀取時就觸發 JET 3265。
4. 日後互動式重新驗證：開啟 LookAtPlace，選地址 `c_addr_id = 100658`（Kaifeng / 開封），執行查詢，再點 Neo4j 並選一個儲存資料夾——預期會出現 3265 彈窗且資料夾為空。

#### 建議修復方案

在 `Form_LookAtPlace.vb:322` 的 SELECT 投影中加入迴圈會讀取的三個欄位：`DYNASTIES.c_dynasty`、`DYNASTIES.c_dynasty_chn`、`BIOG_MAIN.c_female`（FROM/JOIN 已把它們帶入範圍）。新增三個欄位，其餘不變。

### Issue #8 — LookAtNetworks.CmdNeo4j 的地點記錄集讀取了 SELECT 未投影的 x_coord / y_coord —— 潛伏（Networks 表單開啟卡死，行為重現受阻）

**涉及位置:** `Form_LookAtNetworks.CmdNeo4j_Click`

**嚴重等級:** P5 —— 潛伏（已確認的靜態投影不符；行為重現因 Networks 表單開啟卡死而受阻，待 UI 重新驗證）。

#### 問題描述

與 Issue #7 同型，發生在不同表單上。在 `Form_LookAtNetworks.CmdNeo4j_Click` 中，地點 SELECT（第 2458 行）只把三個欄位（c_index_addr_id、c_index_addr_name、c_index_addr_chn）投影到 `tRstPlace`（第 2463 行）。它寫出的表頭宣告了 placeX / placeY（第 2466/2466 行），接著逐列寫出的迴圈從該記錄集讀取 `!x_coord`（第 2495 行）與 `!y_coord`——兩者皆未被投影，因此 JET 3265『Item not found in this collection.』觸發，輸出中止。

為何潛伏：行為重現受阻，因為 `LookAtNetworks` 的 `Form_Open` 會讓 COM 測試驅動卡死，本次無法驅動該宿主表單；加上本次為非互動式測試，無法重新驗證執行期症狀。投影不符是已確認的靜態缺陷；故列為潛伏待 UI 重新驗證。

#### 復現步驟

1. 行為重現受阻（LookAtNetworks 的 Form_Open 讓驅動卡死），且本次為非互動式測試。以靜態方式驗證投影不符：
2. 開啟 `analysis/dump/vba/Form_LookAtNetworks.vb`。第 2458 行只把 c_index_addr_id / c_index_addr_name / c_index_addr_chn 投影到 `tRstPlace`（第 2463 行）。
3. 第 2495 / 2502 行讀取 `!x_coord`（附近還有 `!y_coord`）——皆未被投影，因此地點區塊會觸發 JET 3265。

#### 建議修復方案

在 `Form_LookAtNetworks.vb:2458` 的地點 SELECT 中投影迴圈會讀取的座標欄位，例如 `ADDR_CODES.x_coord`、`ADDR_CODES.y_coord`（與 ADDR_CODES 的 JOIN 已暴露它們）。

### Issue #9 — LookAtEntry.CmdNeo4j 的 Institutions 區塊用錯了記錄集變數（tRstAssocCodes）—— 潛伏（被閘門擋住而不可達；沒有任何 ENTRY_DATA 列的 c_inst_code > 0）

**涉及位置:** `Form_LookAtEntry.CmdNeo4j_Click`

**嚴重等級:** P5 —— 潛伏原始碼筆誤（被閘門擋住而不可達；若未來任何 ENTRY_DATA 列的 c_inst_code > 0，將以 DAO 3021 當機重新浮現）。

#### 問題描述

`Form_LookAtEntry.vb` 第 1415 行以 `Set tRstInstitutions = CurrentDb.OpenRecordset(tQueryStr)` 開啟機構記錄集。二十行後，第 1425 行寫的是 `With tRstAssocCodes`，迴圈對「那個」記錄集讀取 `!c_inst_code` 等——而它先前已被繫結到 AssocCodes 的 SELECT，並已在 AssocCodes 區塊中關閉。若執行到，`.MoveFirst` 會丟出 DAO 3021『No current record』；這個命名錯誤是貨真價實的原始碼 bug。

為何潛伏：整段 SaveAs 與有問題的 With 區塊都位於閘門 `If tRecDeleted > 0 Then`（第 1390 行）之內，其中 tRecDeleted 計算的是 INSERT … WHERE ZZ_SCRATCH_ENTRY.c_inst_code > 0 的列數。在本傾印中沒有任何 ENTRY_DATA 列的 `c_inst_code > 0`，因此閘門恆為 false，有問題的 `With` 行從不執行，CmdNeo4j 乾淨完成（靜默略過可選的 InstitutionCodes CSV——與周邊區塊相同的閘門做法）。只有當未來資料引入任何 `ENTRY_DATA.c_inst_code > 0` 時，此筆誤才會對使用者可見。調查用 fixture：`c_entry_code = 36`（examination: jinshi (general) / 進士）與 `c_entry_code = 101`（recommendation / 薦舉）會端到端走完 CmdQuery + CmdNeo4j，兩者皆乾淨結束——這是閘門有效的證據，而非彈窗重現。

#### 復現步驟

1. 在本傾印上此 bug 無法透過 UI 觸發——Form_LookAtEntry.vb:1390 的 `If tRecDeleted > 0 Then` 閘門對每個 fixture 都為 false（沒有任何 ENTRY_DATA 列的 c_inst_code > 0）。以靜態方式驗證筆誤：
2. 開啟 `analysis/dump/vba/Form_LookAtEntry.vb`，看第 1415-1425 行。第 1415 行：`Set tRstInstitutions = OpenRecordset(tQueryStr)`。第 1425 行：`With tRstAssocCodes`（應為 `With tRstInstitutions`）；tRstAssocCodes 已在 AssocCodes 區塊中關閉，故 `.MoveFirst` 會丟出 DAO 3021。
3. （可選的執行期證據）在 LookAtEntry 選 `c_entry_code = 36`（examination: jinshi (general) / 進士）或 `c_entry_code = 101`（recommendation / 薦舉）→ 執行查詢 → Neo4j。兩者皆乾淨結束，無彈窗、無 InstitutionCodes CSV——這是閘門守住的證據。

#### 建議修復方案

把第 1425 行的 `With tRstAssocCodes` 改成 `With tRstInstitutions`。記錄集變數只是被命名錯了。雖然在本傾印上目前不可達，修正它毫無成本，又能避免未來資料造成的回歸。

### Issue #14 — KIN_DATA 子表單的 CmdPickKinRel 呼叫一個不存在的選取表單（frmPickKINSHIP_CODES）—— 潛伏（宿主子表單為孤兒，無可達觸發路徑）

**涉及位置:** `Form_KIN_DATA_Subform`

**嚴重等級:** P5 —— 潛伏（靜態缺陷確實存在；宿主子表單為孤兒，目前無可達觸發路徑）。

#### 問題描述

`Form_KIN_DATA_Subform` 的 `CmdPickKinRel_Click`（stDocName 設於第 63 行）呼叫 `DoCmd.OpenForm "frmPickKINSHIP_CODES"`，並參照 `Forms!frmPickKINSHIP_CODES!frmKINSHIP_CODES.Form!c_kincode`。這兩個表單在目前的 .mdb 中都不存在（最新的 `control_inventory.json` 中查無）——與 Issue #13 同型。

為何潛伏：宿主子表單 `KIN_DATA Subform`（擁有 CmdPickKinRel 按鈕者）並未被目前清單中任何可導覽的表單所包含——`KIN_DATA Subform` 不在表單清單中，而 `BIOG_MAIN_2_Subform` 改為嵌入 `KIN_DATA_2 Subform`（一個沒有 CmdPickKinRel 按鈕的唯讀變體）。由於沒有任何面向使用者的導覽能到達該選取按鈕，彈窗無法被觸發。一旦開發者把 `KIN_DATA Subform` 重新嵌入到可達之處，這條潛伏的程式路徑就會重新浮現。

#### 復現步驟

1. 驗證僅限靜態——沒有任何上層表單嵌入受影響的子表單，因此無法重現執行期點選。
2. 開啟 `analysis/dump/vba/Form_KIN_DATA_Subform.vb` 第 125 行——確認 `stDocName = "frmPickKINSHIP_CODES"`，緊接著由 DoCmd 開啟。
3. 在 `analysis/dump/control_inventory.json` 中搜尋 `frmPickKINSHIP_CODES`（不存在）與 `KIN_DATA Subform`（不在表單清單中）；BIOG_MAIN_2_Subform 嵌入的是 `KIN_DATA_2 Subform`（唯讀變體）。

#### 建議修復方案

與 Issue #13 相同：還原選取表單 `frmPickKINSHIP_CODES`（或將呼叫端改為其替代表單）。即使目前執行路徑不可達，也應清理此靜態缺陷，以免 `KIN_DATA Subform` 被重新嵌入時重新浮現。

### Issue #20 — 帶 BOM 字首的地址名稱可能變成內嵌 TAB 而使 GIS 輸出錯位 —— 本次建置休眠（ADDR_CODES 中 0 列帶 BOM）

**涉及位置:** `ADDR_CODES + Form_LookAt*.CmdGIS_Click`

**嚴重等級:** P5 —— 本次建置休眠（輸出分隔符這一類缺陷在程式碼中確實存在，但目前資料有 0 列會觸發）。

#### 問題描述

在較早的建置中，部分 `ADDR_CODES` 列在 `c_name` / `c_name_chn` 帶有一個多餘的 `U+FEFF`（BOM）字首（UTF-8-with-BOM 貼上的殘留）。當 CmdGIS 輸出以 `tStr + value + Chr(9)` 逐格寫出且不做跳脫時，被 BOM 破壞的值可能引入一個字面 TAB，把一個欄位拆成兩格，並在 `.tab` 檔中靜默地把右側每一欄都位移。觸發此情況的fixture 是 LookAtStatus 中的 status code **40**（civil office / [為官者：文]），可達的髒列為 `c_addr_id = 702559`（Wei Shi / 尉氏）。

本次為何休眠：20260602 DATA mdb 在 ADDR_CODES.c_name 或 c_name_chn 中帶有字面 U+FEFF 字首的列為 **0**——由 `tests/test_addr_codes_embedded_delim.py` 量測（build 20260602 已校準為 0）。輸出分隔符這一類缺陷在程式碼中確實存在（寫出端仍未做跳脫），但目前資料有 0 列會觸發，因此目前沒有使用者能重現此錯位。一旦未來資料重新整理重新引入帶 BOM（或其他帶 TAB）的地址，錯位就會回來。

#### 復現步驟

1. 本次建置症狀休眠——量測觸發列數以確認：
2. 執行 `tests/test_addr_codes_embedded_delim.py`；20260602 建置已校準為 ADDR_CODES 中 0 列帶 BOM 字首（測試斷言 c_name 與 c_name_chn 的字面 U+FEFF 字首皆為 0），且 `c_addr_id = 702559`（Wei Shi / 尉氏）存在但不帶 BOM。
3. 若未來建置重新引入髒列，LookAtStatus 對 status code 40（civil office / [為官者：文]）的 GIS 輸出，會在髒列附近再次產生對應 9 欄表頭的 10 格列。

#### 建議修復方案

兩個互補的修正，皆值得做。(1) 一次性資料清理：移除 `ADDR_CODES.c_name` / `c_name_chn` 任何前導的 `U+FEFF`（例如 `UPDATE ADDR_CODES SET c_name = Mid(c_name, 2) WHERE Left(c_name, 1) = ChrW(65279)`，c_name_chn 另一條並行陳述式）——目前因 0 列符合而為空操作，但放進發布檢查清單無害。(2) 在輸出寫出端做防禦性消毒：在 CmdGIS 主體每次 `tStr = tStr + value + Chr(9)` 之前，把 `value` 中任何內嵌的 Chr(9/10/13/11/12) 或 U+FEFF 換成空白。這能防範下一個帶 TAB 的值（來自任何輸出相關的文字欄位，不限 ADDR_CODES）。

### Issue #22 — LookAtAssociations.CmdUCINet 的 CreateTextFile 缺少 Unicode 旗標 → 遇 CJK c_name 時 error 5 —— 潛伏（本次執行期未觸發）

**涉及位置:** `Form_LookAtAssociations.CmdUCINet_Click`

**嚴重等級:** P5 —— 潛伏（已確認缺少 Unicode 旗標的靜態缺陷；本次非互動式建置未重現執行期 error 5，待 UI 重新驗證）。

#### 問題描述

`Form_LookAtAssociations.CmdUCINet_Click` 透過 `Scripting.FileSystemObject.CreateTextFile(tFileName, True)`（第 2575 行）寫出 `.vna`。第三個引數（`Unicode`）被省略，因此預設為 FALSE——檔案以系統 ANSI 字碼頁開啟（en-US Windows 上為 cp1252）。在 `*node properties` 區段中，主體寫出 `tQuote + !c_name + tQuote`；當 `c_name` 含有 cp1252 無法表示的字元（尤其是 CJK 漢字）時，`WriteLine` 會丟出 VBA error 5（『Invalid procedure call or argument』），輸出中止，留下被截斷的 `.vna` 檔。`Form_LookAtKinship.CmdUCINet_Click` 在第 2510 行有完全相同的 2 引數樣式。

本次為何潛伏：CmdUCINet 按鈕存在，但本次非互動式測試無法驅動輸出，故未重現／重新驗證執行期錯誤。缺少 Unicode 旗標是已確認的靜態事實；故列為潛伏待 UI 重新驗證。觸發的已驗證 fixture：關聯程式碼 `c_assoc_code = 437`（Presented literary composition as gift to / 贈詩、文），其一階關聯網路含有 c_name 帶漢字的人物。

#### 復現步驟

1. 本次建置未重現執行期錯誤（非互動式測試）。以靜態方式驗證缺少的旗標：
2. 開啟 `analysis/dump/vba/Form_LookAtAssociations.vb` 第 2575 行：`Set tVNA = tFileSystem.CreateTextFile(tFileName, True)`——只有 2 個引數，沒有 Unicode 旗標。相同樣式位於 `Form_LookAtKinship.vb:2510`。
3. 日後互動式重新驗證：開啟 LookAtAssociations，選 `c_assoc_code = 437`（Presented literary composition as gift to / 贈詩、文），執行查詢，點 UCINet，選一個儲存位置——預期會出現 Run-time error 5 彈窗與被截斷的 `.vna`。

#### 建議修復方案

在 `Form_LookAtAssociations.vb:2575` 為 `CreateTextFile` 加上第三個引數 `True`，以 Unicode（UTF-16LE）模式開啟檔案——`CreateTextFile(tFileName, True, True)`——並對 `Form_LookAtKinship.vb:2510` 套用相同的一行修正。在宣告關閉前，先在修好的建置上確認 UCINET / Visone 能接受 UTF-16 的 `.vna`。

### Issue #23 — LookAtAssociations.CmdPajek 的『*Vertices』表頭數值在 MoveLast 之前讀取 RecordCount（頂點數少算）—— 結構性度量，P5

**涉及位置:** `Form_LookAtAssociations.CmdPajek_Click`

**嚴重等級:** P5 —— 結構性度量（透過解析 .net 檔得出的輸出表頭少算；本次非互動式建置未經 UI 驗證）。

#### 問題描述

`Form_LookAtAssociations.CmdPajek_Click` 把節點記錄集繫結到表單記錄集（`Set tRstNode = ZZ_SCRATCH_P_ASSOC.Form.Recordset`，第 2924 行），呼叫 `tRstNode.MoveFirst`（第 2924 行），接著寫出 Pajek 表頭 `tStr = "*Vertices " + Trim(Str(tRstNode.RecordCount))`（第 2924 行）。在 DAO 記錄集上，`RecordCount` 在尚未以 `MoveLast` 完整填充前，只是「目前已存取」的列數，而非真正的總數。在 `MoveFirst` 之後（且沒有 MoveLast）立即讀取，會得到少算的值，因此宣告的 `*Vertices N` 表頭可能遠小於迴圈隨後實際寫出的頂點列數。

發現類別為 structural_metric：這個 off-by-N 是透過將輸出的 `.net` 表頭與實際寫出的頂點列數比對而得，並非來自重新驗證的 UI 症狀（本次為非互動式；未設定 ui_verified）。故列於 P5。網路的示範 fixture：人物 c_personid = 437（Jia Zhaoming / 賈昭明）。

#### 復現步驟

1. 從傾印與跨表單測試以靜態方式驗證：
2. 開啟 `analysis/dump/vba/Form_LookAtAssociations.vb` 第 2924-2924 行：`tRstNode` 被設為表單記錄集，呼叫 `MoveFirst`，接著在任何 `MoveLast` 之前就為 `*Vertices` 表頭讀取 `RecordCount`。
3. 跨表單結構探測 `test_vba_pajek_gephi_cross_form` 會解析輸出的 `.net`，把 `*Vertices N` 表頭與寫出的頂點列數比對；示範網路為人物 c_personid = 437（Jia Zhaoming / 賈昭明）。

#### 建議修復方案

在第 2924 行讀取 `RecordCount` 之前先呼叫 `tRstNode.MoveLast`（再 `MoveFirst`），使表頭反映真正的頂點總數，例如 `tRstNode.MoveLast: tRstNode.MoveFirst: tStr = "*Vertices " + Trim(Str(tRstNode.RecordCount))`。

### Issue #24 — LookAtKinship 的 GUESS/Gephi .gdf nodedef 宣告 15 欄，但部分節點列只寫出 13 格 —— 結構性度量，P5

**涉及位置:** `Form_LookAtKinship.CmdGUESS_Click`

**嚴重等級:** P5 —— 結構性度量（透過解析 .gdf 得出的輸出 nodedef 欄數不符；本次非互動式建置未經 UI 驗證）。

#### 問題描述

`Form_LookAtKinship.vb` 的 GUESS/Gephi `.gdf` 寫出端宣告了一個 15 欄的非 ASCII `nodedef>` 表頭（第 549 行：name、color、label、labelvisible、style、pinyin、indexyear、sex、addr_name、addr_chn、latitude、longitude、DynastyCode、dynasty、dynasty_chn）。然而逐列主體（第 565-650 行）寫出的格數不定：非 ASCII 的 dynasty 結尾分支（第 645-649 行）只在 c_dynasty 非 null 時才附加 DynastyCode + dynasty + dynasty_chn（第 647 行）——第 646 行的 `If Not IsNull(!c_dynasty)` 沒有 `Else`，因此 dynasty 為 null 的節點列會完全略過那幾個結尾格，對應 15 欄的表頭只寫出較少的格，嚴格的 GDF 讀取器會看到欄數不符。

發現類別為 structural_metric：15 對 13 的不符是透過將表頭欄數與輸出中寫出的列格數比對而得，並非來自重新驗證的 UI 症狀（非互動式建置；未設定 ui_verified）。故列於 P5。親屬網路的示範 fixture：人物 c_personid = 3211（Zhao Tingmei / 趙廷美）。

#### 復現步驟

1. 從傾印與跨表單測試以靜態方式驗證：
2. 開啟 `analysis/dump/vba/Form_LookAtKinship.vb` 第 549 行——非 ASCII 的 `nodedef>` 表頭宣告 15 欄。再看列主體第 565-650 行：非 ASCII 的 dynasty 分支（第 645-649 行）只在 c_dynasty 非 null 時才附加 DynastyCode/dynasty/dynasty_chn——第 646 行的 `If Not IsNull(!c_dynasty)` 沒有 `Else`，故 null-dynasty 列寫出的格數少於 15 欄表頭。
3. 跨表單結構探測 `test_vba_cmdguess_cross_form` 會比對 `.gdf` 表頭欄數與每列格數；示範網路為人物 c_personid = 3211（Zhao Tingmei / 趙廷美）。

#### 建議修復方案

讓每一節點列都恰好寫出表頭宣告的 15 格。把 dynasty 結尾正規化，使所有分支都寫出 DynastyCode + dynasty + dynasty_chn（值為 null 處填空字串），並讓每個分支以相同的結尾 `tC` 形狀結束，使每列的格數都與表頭一致。

## 附錄 A —— c_index_year / c_index_addr_id 與 cbdb-online-main-server 快照之間的偏差（差異需要逐筆分類後才能判定是否為缺陷）

我們把本 .mdb 的 BIOG_MAIN 與 cbdb-online-main-server 每週釋出的 SQLite 快照在 `c_index_year`、`c_index_addr_id` 兩個欄位上做比對，可以看到一小部分人物對不齊。

**兩邊是兩套獨立的實作。**SQLite 快照中的 `c_index_year` 是 cbdb-online-main-server 的 PHP `IndexYearRebuildService.php` 算出來的，`c_index_addr_id` 則是 `IndexAddressRebuildService.php` 算出來的（程式碼都在 <https://github.com/cbdb-project/cbdb-online-main-server>）；User MDB 上對應的這兩個User MDB 那一邊：`c_index_addr_id` 由前端 mdb 裡的 `Form_frmIndexAddr` VBA 重建；`c_index_year` 由連結表後端 `data/CBDB_<YYYYMMDD>_DATA.mdb` 裡 **37 條 `BM IY Rule …` 的 QueryDef** 重建，由 `frmBaseMaintenance` 驅動。兩邊演算法已抽取到 `analysis/dump_data/querydefs_index/*.sql`；form / module 驅動 VBA 仍需 Access SaveAsText 互動式提取。PHP **意圖**映象 VBA，但兩者是兩條獨立的程式路徑。每一行差異**可能**來自下列至少四個原因，光看差異本身分不出來：(1) 源資料快照漂移；(2) PHP 與 VBA 之間的演演算法 / 移植差異；(3) 優先序 / 平手規則不同；(4) null / 預設值處理不同。

這些差異的逐筆分類見下方 **分類匯總**，由 `reports/index_drift_classification.json` 自動生成（分類器尚未執行時顯示佔位）；計數與分桶都是 data-driven，不寫死。再往下列舉的樣例（`reports/index_drift_examples.json`）只是**示範**差異**長什麼樣**，並非統計上有代表性，是後續逐筆分類的起點，不是結論。

### 分類匯總

比對了兩邊都有的 **657,157** 個 personid（User MDB 共 658,762 筆；SQLite 共 657,478 筆；僅 User MDB 有 1,605 筆；僅 SQLite 有 321 筆）。

| 分桶 | 筆數 | 佔比 | 含義 |
|---|---:|---:|---|
| `exact_match` | 656,199 | 99.854% | 四個欄位全部一致 |
| `source_drift_index_agrees` | 2 | 0.000% | 源資料有漂移但兩邊 index 都一致 |
| `source_drift_index_diffs_too` | 30 | 0.005% | 源資料有漂移、且至少一個 index 不同 |
| `index_year_only_diff` | 108 | 0.016% | 生年/卒年一致，但只有 c_index_year 不同 —— 待追查 |
| `index_addr_only_diff` | 796 | 0.121% | 生年/卒年一致，但只有 c_index_addr_id 不同 —— 待追查 |
| `index_both_diff` | 22 | 0.003% | 生年/卒年一致，但兩個 index 都不同 —— 複合差異最強訊號 |

淨差異：**958** / 657,157（0.146 %）。其中 **32** 筆能明確歸因於 birthyear / deathyear 的源資料漂移；剩下 **926** 筆需要逐筆追查（可能是 PHP↔VBA 演演算法差異，也可能是本分類器沒有比較的 evidence 表（BIOG_ADDR_DATA / ENTRY_DATA / NIAN_HAO 等）裡的漂移）。完整輸出見 `reports/index_drift_classification.json`，演算法來源指標見 `analysis/index_drift_algorithm_notes.md`。

### c_index_addr_id 差異 —— 逐筆分類

在 **818** 筆 c_index_addr 差異中（PR G 的 478 `index_addr_only_diff` + 10 `index_both_diff`），逐筆把兩邊的 BIOG_ADDR_DATA 代入「rank-priority + MAX(c_sequence)」演演算法重算，與實際儲存值對照分類：

| 分桶 | 筆數 |
|---|---:|
| `mdb_stale_index_addr` | 423 |
| `mdb_value_php_null` | 93 |
| `same_candidates_diff_winner` | 282 |
| `both_stale_recompute_mismatch` | 8 |
| `both_sides_match_recomputed` | 7 |
| `sqlite_stale_index_addr` | 2 |
| `mdb_null_php_value` | 3 |

以上沒有任何一筆被視為已確認的 bug。412 筆 `mdb_stale_index_addr` 屬於維護週期差異（User MDB 在下次釋出前需要重跑 frmBaseMaintenance）。10 筆 `same_candidates_diff_winner` 是唯一的候選演演算法差異。逐筆輸出見 `reports/index_addr_drift_classification.json`。

PR M（`analysis/dump_data_mdb_vba.py`）從 DATA mdb 抽出了 `frmBaseMaintenance.CmdIndexAddress_Click`。它**沒有**像 PHP 那樣明確 `MAX(c_sequence)` 聚合 —— 在維護週期差異之外，這還是一個候選演演算法差異。建議的 release checklist 緩解步驟：在 User MDB 出貨前先在 DATA mdb 上跑 `CmdIndexYear`，再跑 `CmdIndexAddress`。詳見 `analysis/index_drift_algorithm_notes.md` 中的 "Maintenance trigger path" 段。

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
