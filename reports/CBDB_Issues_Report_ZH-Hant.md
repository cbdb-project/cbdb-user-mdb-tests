# CBDB 使用者版 .mdb — 問題彙報

_測試過程中發現的問題彙總，謹呈維護團隊斧正。_

尊敬的維護者：

下面是我們在為 CBDB 使用者版 .mdb 編寫自動化迴歸測試套件過程中，陸續整理出來的一些問題清單。我們希望這份報告能在您繼續主持這份寶貴資料集時有所幫助；同時，對您多年來在這套資料上的辛勤付出，我們由衷地表示感謝和敬意。

問題按嚴重程度排序（P0 最高）。每一條都包括：簡明描述、使用者端一步一步的復現步驟、（在介面上能看到時）相關截圖，以及一份建議的修復方案。這些問題並不緊急，整理在此只是為了方便您在合適的時候逐一處理。

## 目錄

- [P0 — 靜默資料錯誤](#p0--靜默資料錯誤)
  - [Issue #7 — LookAtPlace.CmdNeo4j 在寫入第一條 people-CSV 時靜默失敗](#issue-7--lookatplacecmdneo4j-在寫入第一條-people-csv-時靜默失敗)
  - [Issue #8 — LookAtNetworks.CmdNeo4j 的 people/place CSV 在第一條上靜默失敗](#issue-8--lookatnetworkscmdneo4j-的-peopleplace-csv-在第一條上靜默失敗)
  - [Issue #20 — 地址名中的 BOM 會在 GIS 匯出中變成 tab，造成欄位錯位](#issue-20--地址名中的-bom-會在-gis-匯出中變成-tab造成欄位錯位)
- [P1 — 可見的執行時報錯](#p1--可見的執行時報錯)
  - [Issue #6 — LookAtGroupData 的 ChkEntry 路徑引用了不存在的列 ENTRY_DATA.c_parental_status](#issue-6--lookatgroupdata-的-chkentry-路徑引用了不存在的列-entry_datac_parental_status)
  - [Issue #13 — BIOG_MAIN_2 子表單試圖開啟一個不存在的 picker 表單 (frmPickNIAN_HAO)](#issue-13--biog_main_2-子表單試圖開啟一個不存在的-picker-表單-frmpicknian_hao)
  - [Issue #21 — LookAtGroupData.CmdNeo4j 在匯出空分部時崩潰報「No current record」](#issue-21--lookatgroupdatacmdneo4j-在匯出空分部時崩潰報no-current-record)
  - [Issue #22 — LookAtAssociations.CmdUCINet 在被匯出的人物網路含有 c_name 中帶 CJK 漢字時崩潰報「Invalid procedure call or argument」](#issue-22--lookatassociationscmducinet-在被匯出的人物網路含有-c_name-中帶-cjk-漢字時崩潰報invalid-procedure-call-or-argument)
  - [Issue #23 — LookAtAssociations.CmdNeo4j 的 INSERT 引用了 ZZ_SCRATCH_PEOPLE 上不存在的目標列 c_index_addr_type_code（疑似本意為 c_addr_type）](#issue-23--lookatassociationscmdneo4j-的-insert-引用了-zz_scratch_people-上不存在的目標列-c_index_addr_type_code疑似本意為-c_addr_type)
  - [Issue #24 — LookAtPlace.CmdNeo4j 中 tRstPeople 的 SELECT 投影缺少 c_dynasty / c_dynasty_chn / c_female，下游迴圈仍然讀取這些欄位並崩潰報「Item not found in this collection」（JET 3265）](#issue-24--lookatplacecmdneo4j-中-trstpeople-的-select-投影缺少-c_dynasty--c_dynasty_chn--c_female下游迴圈仍然讀取這些欄位並崩潰報item-not-found-in-this-collectionjet-3265)
- [P2 — 靜默顯示問題](#p2--靜默顯示問題)
  - [Issue #10 — EVENT_ADDR_2 子表單的地址列默默地顯示為空（ControlSource 寫錯了）](#issue-10--event_addr_2-子表單的地址列默默地顯示為空controlsource-寫錯了)
- [P3 — 缺失介面](#p3--缺失介面)
  - [Issue #15 — LookAtPlace 缺少 CmdGIS 按鈕（程式碼裡有 handler 但介面上沒控制元件）](#issue-15--lookatplace-缺少-cmdgis-按鈕程式碼裡有-handler-但介面上沒控制元件)
  - [Issue #16 — LookAtStatus 缺少 CmdPajek 按鈕](#issue-16--lookatstatus-缺少-cmdpajek-按鈕)
  - [Issue #17 — LookAtStatus 缺少 CmdGephi 按鈕](#issue-17--lookatstatus-缺少-cmdgephi-按鈕)
  - [Issue #18 — LookAtStatus 缺少 CmdUCINet 按鈕](#issue-18--lookatstatus-缺少-cmducinet-按鈕)
  - [Issue #19 — LookAtOffice 缺少 CmdGUESS 按鈕](#issue-19--lookatoffice-缺少-cmdguess-按鈕)
- [P4 — 安裝設定](#p4--安裝設定)
  - [Issue #2 — VBA 工程引用了過時的 dao360.dll，Office 2016+ 機器上沒這個檔案](#issue-2--vba-工程引用了過時的-dao360dlloffice-2016-機器上沒這個檔案)
- [P5 — 潛伏 / 不可達 / 當前無法復現](#p5--潛伏--不可達--當前無法復現)
  - [Issue #1 — View_StatusData 會把首年份範圍顯示成末年份範圍 — DORMANT（當前 dump 沒有源資料能觸發）](#issue-1--view_statusdata-會把首年份範圍顯示成末年份範圍--dormant當前-dump-沒有源資料能觸發)
  - [Issue #9 — LookAtEntry.CmdNeo4j 的機構 (Institutions) 部分用錯了記錄集變數 — LATENT（被資料 gate 跳過、不可達；當前 dump 中沒有 c_inst_code > 0 的 ENTRY_DATA 行）](#issue-9--lookatentrycmdneo4j-的機構-institutions-部分用錯了記錄集變數--latent被資料-gate-跳過不可達當前-dump-中沒有-c_inst_code--0-的-entry_data-行)
  - [Issue #4 — LookAtPlace.CmdGIS 會報「Object required」 — LATENT，被 Issue #15（表單上沒有 CmdGIS 按鈕）所遮蔽](#issue-4--lookatplacecmdgis-會報object-required--latent被-issue-15表單上沒有-cmdgis-按鈕所遮蔽)
  - [Issue #5 — LookAtStatus.CmdPajek 引用了不存在的控制元件，且 SQL 用了三個不存在的列](#issue-5--lookatstatuscmdpajek-引用了不存在的控制元件且-sql-用了三個不存在的列)
  - [Issue #14 — KIN_DATA 子表單的 CmdPickKinRel 呼叫不存在的 picker（frmPickKINSHIP_CODES）——但目前該子表單在主表中無入口（LATENT）](#issue-14--kin_data-子表單的-cmdpickkinrel-呼叫不存在的-pickerfrmpickkinship_codes但目前該子表單在主表中無入口latent)
  - [Issue #11 — EVENTS_DATA_2 上 c_event_record_id 控制元件綁到不存在的欄位——但該控制元件本身是隱藏的（LATENT）](#issue-11--events_data_2-上-c_event_record_id-控制元件綁到不存在的欄位但該控制元件本身是隱藏的latent)
  - [Issue #12 — POSTED_TO_OFFICE_DATA_2 上 c_appt_type_code 控制元件綁到沒投影的欄位——但該控制元件是隱藏的，且使用者實際看的任職型別欄位是正常的（LATENT）](#issue-12--posted_to_office_data_2-上-c_appt_type_code-控制元件綁到沒投影的欄位但該控制元件是隱藏的且使用者實際看的任職型別欄位是正常的latent)
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

### Issue #7 — LookAtPlace.CmdNeo4j 在寫入第一條 people-CSV 時靜默失敗

**涉及位置:** `Form_LookAtPlace.CmdNeo4j_Click`

**嚴重等級:** P0 — 靜默資料缺失（匯出無聲地什麼都沒生成）

#### 問題描述

`LookAtPlace.CmdNeo4j_Click` 中負責生成 People-CSV 的部分（約第 322 行起）用 SELECT 開啟記錄集，但 SELECT 裡只投影了四個 `ZZ_SCRATCH_P_TEXT` 欄位；接下來的寫入迴圈卻試著讀 `!c_dynasty`、`!c_dynasty_chn`、`!c_female`。迴圈一碰到第一行，JET 立即報「集合中找不到專案」（Item not found in this collection）。錯誤處理把它彈了一個 MsgBox 就結束了，所以使用者只看到一個對話方塊，之後整個 Neo4j 匯出鏈下游的任何檔案都不會產生。

#### 復現步驟

1. 開啟 **LookAtPlace**。透過地址 picker 選一個資料量足夠的地址——例如 **c_addr_id = 100658（開封）**（這也是 `tests/test_vba_inline.py` 的 kaifeng-yin fixture 用的 addr_id）——這樣查詢結果有足夠人物餵給 People-CSV 迴圈。點 **Run Query**。
2. 等查詢跑完，點 **Neo4j** 匯出按鈕。
3. 在第一個另存對話方塊（「People 檔」對話方塊）裡選好儲存路徑。
4. 幾乎立刻彈出 `執行時錯誤 3265 ——集合中找不到專案` 對話方塊。
5. 點確定後，剛才選的資料夾裡一個檔案也沒有——整個 Neo4j 匯出什麼都沒寫出。

#### 截圖

![bug7_step1_annotated.png](screenshots/bug7_step1_annotated.png)

_Step 1 — open LookAtPlace, run any query, click **Neo4j**.  The CmdNeo4j button is in the bottom export-button row; reachability verified against `analysis/dump/control_inventory.json` (LookAtPlace has a `CmdNeo4j` control with the `CmdNeo4j_Click` event bound — re-checked 2026-05-03)._

![bug7_step2_faux_popup.png](screenshots/bug7_step2_faux_popup.png)

_Step 2 — the popup users see.  Reconstructed in PIL because the real popup would block the COM test driver; the error code (DAO 3265 'Item not found in this collection') and message text come from JET's documented behaviour for a recordset field that isn't in the underlying SELECT._

#### 建議修復方案

把 People-CSV 部分的 SELECT 擴充套件，加入迴圈裡讀到的欄位，例如 `DYNASTIES.c_dynasty`、`DYNASTIES.c_dynasty_chn`、`BIOG_MAIN.c_female`（FROM 子句裡 JOIN 已經把它們暴露出來了）。

### Issue #8 — LookAtNetworks.CmdNeo4j 的 people/place CSV 在第一條上靜默失敗

**涉及位置:** `Form_LookAtNetworks.CmdNeo4j_Click`

**嚴重等級:** P0 — 靜默資料缺失

#### 問題描述

症狀與 Issue #7 相同，只是在另一個表單上。`LookAtNetworks.CmdNeo4j_Click` 中兩條 SELECT 都漏寫了迴圈裡要讀的欄位：

  • `tRstPlace` 的 SELECT（第 2458 行）只投影 3 個欄位，迴圈卻讀 `!x_coord` / `!y_coord`（沒在 SELECT 裡）。
  • `tRstPeoplePlace` 的 SELECT 也漏了 `c_person_id` / `c_index_addr_id`，迴圈要讀它們。

症狀與 Issue #7 完全相同：靜默失敗。

#### 復現步驟

1. 開啟 **LookAtNetworks**（注意：這個表單已知開啟會延遲，請給它幾秒鐘）。
2. 跑一次查詢，然後點 **Neo4j**。
3. 匯出走到 People-with-Place 檔案那一步時，同樣的「Item not found」對話方塊彈出來，之後的檔案都不會再寫了。

#### 截圖

![bug8_faux_popup.png](screenshots/bug8_faux_popup.png)

_Reconstructed-in-PIL popup showing the JET 'Item not found in this collection' error users would see.  **Important caveat:** the backdrop in this image is LookAtPlace, NOT LookAtNetworks — LookAtNetworks's `Form_Open` currently hangs the COM test driver, so a real runtime view of the host form couldn't be captured.  The popup text is reconstructed from VBA static inspection of `Form_LookAtNetworks.vb:2458` / `:2475`._

#### 建議修復方案

把兩條 SELECT 都擴充套件，加入迴圈裡讀到的欄位。對 tRstPlace 加上 `ADDR_CODES.x_coord`、`ADDR_CODES.y_coord`。對 tRstPeoplePlace 加上 `c_person_id` 和 `c_index_addr_id`。

### Issue #20 — 地址名中的 BOM 會在 GIS 匯出中變成 tab，造成欄位錯位

**涉及位置:** `ADDR_CODES + Form_LookAt*.CmdGIS_Click`

**嚴重等級:** P0 — 靜默匯出欄位錯位（數字欄位落到文本欄，所有欄位向右挪一格，結尾多出一欄）

#### 問題描述

`ADDR_CODES` 中有 315 行在 `c_name` **和** `c_name_chn` 裡都帶著 `U+FEFF`（BOM）字首，幾乎可以確定是資料匯入時從 UTF-8-with-BOM 文件複製貼上留下的痕跡。當 `LookAtStatus.CmdQuery`（以及其他 LookAt 表單的對應 CmdQuery / CmdRun）把這些行通過 SQL UPDATE/INSERT 複製到自己的 scratch 暫存表時，JET 會先把 BOM 去掉，再把剩下的 UTF-16 LE 位元組重新當成單位元組字元——升回 Unicode 之後值就被破壞了。以 `c_addr_id = 702559`（尉氏）為例，源字串 `﻿尉氏`（UTF-16 位元組 `ff fe 09 5c 0f 6c`）變成了暫存字串 `\t\\\x0fl`（UTF-16 位元組 `09 00 5c 00 0f 00 6c 00`），第 0 位上多了一個**真正的 TAB 字元**。

隨後 `Form_LookAtStatus.CmdGIS_Click`（第 1554–1636 行）把每個欄位寫成 `tStr + value + tC`，其中 `tC = Chr(9)` （第 1552 行）——完全沒有做任何轉義。這個嵌入的 TAB 就被當作分隔符，把 AddrChn 拆成兩欄，往後所有的欄位都悄無聲息地往右挪一格。使用者在 Excel 裡開啟這份 `.tab` 檔，會看到座標落在錯誤的欄位、還多出一個尾欄。LookAtTexts / LookAtPlace / LookAtAssociations / LookAtOffice / LookAtKinship 的 CmdGIS 都用同樣的 `tStr + value + tC` 模式，所以任何 LookAt 表單只要查詢結果裡碰到這 315 個髒地址裡的任何一個，都會重現同樣的欄位錯位。

證據——完整的位元組級追蹤在 `analysis/gis_status_embedded_delim_root_cause.md`；源端掃描在 `reports/gis_embedded_delimiter_findings.json`；實際匯出檔的位元組級 dump 在 `reports/gis_status_export_bytes_dump.json`。迴歸測試 `tests/test_addr_codes_embedded_delim.py` 會在上游資料被清理後**主動失敗**，提醒重新評估。

**已知影響面（PR W）。** 在這 315 行髒 `ADDR_CODES` 裡，**只有 1 行**（`c_addr_id = 702559` / 尉氏）真的被任何人物記錄引用——透過 `BIOG_MAIN.c_index_addr_id` 或 `BIOG_ADDR_DATA`；其餘 314 行在 ADDR_CODES 表裡是孤立的，沒有任何人物掛上去。所以今天的使用者實際影響面其實很小：在 **LookAtStatus**（`c_status_code=40` fixture，正是本 issue 立案的那一行）已有位元組級實證；在 **LookAtKinship**（如果選到那 3 位以阮孚為親屬的人）和 **LookAtPlace**（如果使用者選 `c_addr_id = 702559`）屬於「按源資料看應該會觸達」；在 **LookAtTexts / LookAtAssociations / LookAtOffice** 在當前源資料下根本觸達不到。完整的逐表分析在 `analysis/gis_embedded_delimiter_reach.md` 與 `reports/gis_embedded_delimiter_reach.json`。其餘 314 行是一個**潛伏的資料品質問題**——它們一旦有第一個人物掛上去，就會重現同樣的欄位錯位。前面建議的兩條修法依然都值得做。

#### 復現步驟

1. 開啟 **LookAtStatus**。在 status picker 裡挑 status code **40**（[為官者：文] / civil office），不要設年份過濾——測試 fixture 裡 `FrameFilterYears = 1`。
2. 點 **Run Query**。結果網格里大約填進 17 000 行。
3. 點 **GIS**，把編碼選成 UTF-8（`GISFrame = 1`）。把匯出的 `.tab` 檔存下來。
4. 在任意支援 tab 的工具（Excel / 帶欄位標尺的文本編輯器）裡開啟這個檔。第 **11476** 行附近（對應人物阮孚，`c_addr_id = 702559` / 尉氏）有一行包含 10 個 tab 欄位、卻對著 9 欄的表頭。AddrChn 是空的、X 欄裡塞了文字，真正的 X / Y 值都往右挪了一欄。

#### 建議修復方案

兩條互補的修法，建議都做：

  1. **一次性資料清理。** 把這 315 行 `ADDR_CODES.c_name` / `c_name_chn` 開頭的 `U+FEFF` 去掉（例如 `UPDATE ADDR_CODES SET c_name = Mid(c_name, 2) WHERE Left(c_name, 1) = ChrW(65279)`，再對 `c_name_chn` 重複一遍）。這一步可以立即解決使用者能看到的欄位錯位。

  2. **匯出端做防禦性 sanitisation。** 在 LookAtStatus / Texts / Place / Associations / Office / Kinship 各自的 CmdGIS 裡，每一個 `tStr = tStr + value + tC` 之前，先把 `value` 裡的 Chr(9)、Chr(10)、Chr(13)、Chr(11)、Chr(12)、`U+FEFF` 全部替換成空格。這樣以後任何 text 欄位如果再混進類似的髒字符，匯出依然能保持欄位對齊——少了這一層，下一次只要 `ADDR_CODES.c_name`（或 `BIOG_MAIN.c_name`、或其他這些匯出會碰到的 text 欄位）裡悄悄塞進一個 tab 字元，又會重現完全一樣的靜默錯位。

  3. **建議增加一個釋出前的檢查指令碼。** 寫一個簡短的指令碼，釋出前掃描所有會被匯出的 text 欄位，看裡面有沒有分隔符或控制字元，可以在每次釋出前提前抓到這一類髒資料問題。`analysis/probe_status_gis_embedded_delim.py` 是一個現成的起點。

## P1 — 可見的執行時報錯

### Issue #6 — LookAtGroupData 的 ChkEntry 路徑引用了不存在的列 ENTRY_DATA.c_parental_status

**涉及位置:** `Form_LookAtGroupData.queryEntry`

**嚴重等級:** P1 — 常用路徑上的可見報錯（Entry 子查詢）

#### 問題描述

`Form_LookAtGroupData.vb` 第 2621 行的 INSERT INTO 目標列裡寫的是 `c_parental_status_code`，但 SELECT 投影寫的是 `ENTRY_DATA.c_parental_status`（少了 `_code` 字尾）。`ENTRY_DATA` 上真實列名是 `c_parental_status_code`；這個筆誤讓使用者一旦勾上 **Entry** 子型別再點 **Run**，SQL 就會報「無此欄位」。

`Form_LookAtEntry.vb:1650` 寫的是同一段邏輯查詢而且名字是對的，可以參考。

#### 復現步驟

**建議使用的範例人物：** `c_personid=1`（安惇，An Dun）

用人物 1（安惇，An Dun）當匯入清單（資料少、只有 2 條 entry 記錄，方便復現）。在 LookAtGroupData 上只勾 **Entry**，點 **Run**。 _由 `reports/probe_demo_persons.py` 透過 SQL probe 挑選；之所以選這位，是因為其底層記錄數確實滿足這個 bug 的觸發條件。_

1. 在 **LookAtGroupData** 上把匯入清單設為一個人——例如 c_personid = 1（安惇 An Dun），他只有 2 條 ENTRY_DATA 記錄，足以讓有缺陷的 queryEntry SQL 在一個小而熟知的樣本上跑起來。
2. **只**勾 **Entry** 核取方塊（Status / Office / Text / Addr 都不勾，避免無關的查詢分支幹擾）。
3. 點 **Run**。
4. 彈出「欄位不存在」之類的對話方塊（JET 在不同 Office 版本下給出的措辭是「沒有為一個或多個必要引數提供值」或「沒有此欄位」——都是因為 SQL 引用了根本不存在的 `ENTRY_DATA.c_parental_status`）。

#### 截圖

![bug6_form_annotated.png](screenshots/bug6_form_annotated.png)

_Step 1 — open LookAtGroupData, leave only the Entry checkbox ticked, click Run.  (Demo input from `reports/demo_persons.json`: import list = c_personid 1, 安惇.)_

![bug6_faux_popup.png](screenshots/bug6_faux_popup.png)

_Step 2 — the JET error popup users see.  The popup graphic is reconstructed in PIL because the real popup would block the COM test driver; the error code (3061) and message text come from JET's documented behaviour for unknown-identifier-as-parameter on the line cited in the summary._

#### 建議修復方案

把第 2621 行的 `ENTRY_DATA.c_parental_status` 改成 `ENTRY_DATA.c_parental_status_code`。一行修復。

### Issue #13 — BIOG_MAIN_2 子表單試圖開啟一個不存在的 picker 表單 (frmPickNIAN_HAO)

**涉及位置:** `Form_BIOG_MAIN_2_Subform.c_fl_ey_notes_Click`

**嚴重等級:** P1 — 使用者點選時可見的報錯

#### 問題描述

使用者在某位人物生平詳情子表單上點選 `c_fl_ey_notes` 欄位時，`Sub c_fl_ey_notes_Click` 會呼叫 `DoCmd.OpenForm "frmPickNIAN_HAO"`。.mdb 的 CurrentProject.AllForms 集合中並沒有名為 `frmPickNIAN_HAO` 的表單。Access 報「集合中找不到專案」，使用者的這一次點選就此無效。

可能原因：picker 表單在某次重構中被重新命名或合併了，而這個呼叫處沒有跟著更新。

#### 復現步驟

**建議使用的範例人物：** `c_personid=5`（查籥，Zha Yue）

開啟人物 5（查籥，Zha Yue）。其 `c_fl_ey_notes` 欄位有真實內容（樣例：「紹興二十一年進士。…」），所以點選它會真的觸發這個有缺陷的 Sub。 _由 `reports/probe_demo_persons.py` 透過 SQL probe 挑選；之所以選這位，是因為其底層記錄數確實滿足這個 bug 的觸發條件。_

1. 開啟人物 **c_personid = 5（查籥 Zha Yue）** 的生平詳情——之所以選他，是因為其 BIOG_MAIN 上 `c_fl_ey_notes` 欄位有實際內容，欄位可點（點一個空欄位不會觸發這個 Sub）。
2. 在 BIOG_MAIN_2 子表單上點 `c_fl_ey_notes` 欄位——這會觸發 `c_fl_ey_notes_Click` Sub。
3. 彈出「集合中找不到專案」對話方塊（因為 Sub 試圖 `DoCmd.OpenForm "frmPickNIAN_HAO"`，而該表單根本不存在）。

#### 截圖

![bug13_browser_annotated.png](screenshots/bug13_browser_annotated.png)

_Step 1 — open CBDB_Browser_2, navigate to c_personid=5 (查籥 Zha Yue), click the `c_fl_ey_notes` field on the Birth/Death sub-tab (this fires `c_fl_ey_notes_Click`)._

![bug13_faux_popup.png](screenshots/bug13_faux_popup.png)

_Step 2 — the popup users see.  Reconstructed in PIL because the real popup would block the COM test driver; error 2102 + 'misspelled or refers to a form that doesn't exist' is Access's standard message when DoCmd.OpenForm targets a form not in CurrentProject.AllForms._

#### 建議修復方案

要麼把 `frmPickNIAN_HAO` 表單恢復回來，要麼在 `Form_BIOG_MAIN_2_Subform.c_fl_ey_notes_Click` 裡把呼叫改成替代的那個 picker 表單。

### Issue #21 — LookAtGroupData.CmdNeo4j 在匯出空分部時崩潰報「No current record」

**涉及位置:** `Form_LookAtGroupData.CmdNeo4j_Click`

**嚴重等級:** P1 — 正常使用者點選下的可見報錯（只要查詢的人沒有 Entry 資料，GroupData 的 Neo4j 匯出就會觸發 —— 這對入仕記錄稀薄的人物屬於常見情況）

#### 問題描述

`Form_LookAtGroupData.CmdNeo4j_Click` 的多個 CSV 匯出區塊在開啟暫存記錄集後直接呼叫 `.MoveFirst`，沒有先檢查記錄集是否為空（缺少 `.EOF` 或 `.RecordCount > 0` 防護）。如果使用者查詢的群體在某個類別沒有資料（例如沒有 Entry 資料，導致 `ZZ_SCRATCH_ENTRY` 為空），`.MoveFirst` 呼叫會立刻丟擲 DAO 3021「No current record」，彈出報錯框並中斷整個 Neo4j 匯出。

在當前 dump 和典型小 fixture 下，**使用者能首先觸發的失敗是 block #9 PeopleEntry**（line 1243-1245）—— `Set tRstPeopleEntry = CurrentDb.OpenRecordset("ZZ_SCRATCH_ENTRY", dbOpenDynaset)` 後緊接無防護的 `.MoveFirst`。同樣的缺少防護模式也存在於緊隨其後的 **block #10 EntryCode**（line 1383-1385）以及其他沒有上游 gate 的 tail block；block #10 之所以目前沒單獨表現為故障，只是因為鏈條已經在 block #9 中斷了。

關於程式碼範圍的說明：block #1-#8 共用相同的 `Set ... = OpenRecordset(...)` + `.MoveFirst` 寫法，但其上游暫存表（ZZ_SCRATCH_STATUS, ZZ_SCRATCH_OFFICE, ZZ_PLACE, ZZ_SCRATCH_P_TEXT, ZZ_ADDRESSES）在任何正常使用者啟用範圍下都不會為空，所以 empty-feeder 失敗模式 在這些 block 上 使用者無法觸達。Block #11 InstitutionCodes（line 1485-1487）在上游已經被 `If tRecDeleted > 0 Then` 正確 gate 住，不屬於本 issue 範圍。要修復使用者可見的症狀，只需在 block #9 和 #10 加 guard。

這與 Issue #6 **不同**：Issue #6 是 `queryEntry` 裡的列名筆誤（`ENTRY_DATA.c_parental_status` 應為 `c_parental_status_code`），導致 ChkEntry 勾選時 `ZZ_SCRATCH_ENTRY` 寫不進資料。Issue #21 是 `CmdNeo4j_Click` 裡獨立的下游缺少防護的 bug，只要 `ZZ_SCRATCH_ENTRY` 為空就觸發 —— 既可能是上游 Issue #6 的連帶影響，也可能是使用者單純沒勾 ChkEntry。兩個不同層級的程式碼缺陷，應分別歸檔與修復。

#### 復現步驟

1. 在 **LookAtGroupData** 上把匯入清單設為 c_personid = 1（安惇 An Dun）——他有 2 條 STATUS_DATA / 2 條 ENTRY_DATA / 約 12 條 POSTED_TO_OFFICE，鏈條針對 Status / Office 有實際資料，但只要不勾 ChkEntry，`ZZ_SCRATCH_ENTRY` 就會保持為空（勾上 ChkEntry 會另外觸發 Issue #6）。
2. 勾 **Status**、**Office**、**Addr** 及對應的 **GIS** 三個子項——**Entry 不要勾**（讓 ZZ_SCRATCH_ENTRY 維持為空）。點 **Run**。
3. CmdRun 完成後（ZZ_SCRATCH_STATUS 寫入 2 行，ZZ_SCRATCH_OFFICE 寫入 12 行），點選 **Neo4j** 匯出按鈕。
4. 鏈條會先順利產出 8 份 CSV（People / Places / PeoplePlaces / PersonPlaceCodes / PeopleStatus / StatusCode / PeopleOffice / OfficeCodes），然後在進入第 9 個 block（PeopleEntry）時彈出 `執行時錯誤 3021 —— No current record` 對話方塊。餘下的 2-3 份預期檔案（PeopleEntry / EntryCode / 可選的 InstitutionCodes）不會寫出。
5. 已在 `analysis/probe_groupdata_cmdneo4j.py` 與 `analysis/probe_groupdata_cmdneo4j_tail.py` 端到端驗證 —— tail probe 的 iter 3 split-then-seed 手動向 ZZ_SCRATCH_ENTRY 插入一行後，鏈條產出 10 份檔案且沒有 ERR（證明觸發條件就是空記錄集，沒有其他變數）。

#### 建議修復方案

在 `Form_LookAtGroupData.CmdNeo4j_Click` 的 **block #9 PeopleEntry（約 line 1245）和 block #10 EntryCode（約 line 1385）** 上，於 `.MoveFirst` 呼叫前加上 `.EOF`（或 `.RecordCount > 0`）防護。這兩段是使用者能觸發的兩塊；其他 tail block 要麼上游有 gate（block #11 InstitutionCodes 上游有 `If tRecDeleted > 0 Then`），要麼其上游暫存表在任何正常使用者啟用範圍下都非空（block #1-#8）。建議寫法：

```vb
Set tRstPeopleEntry = CurrentDb.OpenRecordset("ZZ_SCRATCH_ENTRY", dbOpenDynaset)
With tRstPeopleEntry
    If Not .EOF Then
        .MoveFirst
        Do While Not .EOF
            ' ... 原有的逐行寫出 ...
            .MoveNext
        Loop
    End If
End With
```

防禦性範圍選項（關閉本 issue 不需要）：把同樣的 guard 加到 block #1-#8 也無害，可對未來出現新的啟用路徑讓其上游表為空的情況做防禦。當前沒有這種路徑，所以不必做。

### Issue #22 — LookAtAssociations.CmdUCINet 在被匯出的人物網路含有 c_name 中帶 CJK 漢字時崩潰報「Invalid procedure call or argument」

**涉及位置:** `Form_LookAtAssociations.CmdUCINet_Click`

**嚴重等級:** P1 — 正常使用者點選下的可見報錯。任何使用者只要在 `LookAtAssociations × CmdUCINet` 上發起的查詢結果集中含有 c_name 帶 CJK 漢字的人，就會遇到 Run-time error 5 對話方塊，匯出全部失敗。當前 dump 上至少以 c_assoc_code = 437 為 fixture 驗證的 8087 行暫存表會觸發；更大範圍的影響取決於 BIOG_MAIN 中 c_name（理論上是 Latin / Pinyin）含漢字的行數 —— 至少 2 行落在該查詢結果集裡，只要結果集含其中任一行就會受影響。

#### 問題描述

`Form_LookAtAssociations.CmdUCINet_Click` 透過 `Scripting.FileSystemObject.CreateTextFile(tFileName, True)`（約 line 2575）寫出 `.vna` 檔案。第 3 個引數（`Unicode`）省略掉了，預設值為 FALSE，所以檔案以系統預設的 ANSI 碼頁（en-US Windows 上是 cp1252）開啟。

在 `*node properties` 區段，每一行寫出 `tQuote + !c_name + tQuote`。當 `c_name` 含有 cp1252 編碼無法對應、且 FSO 也沒有替代對映的字元（特別是 CJK 漢字，例如 U+7A1C 稜），`tVNA.WriteLine` 就會丟出 VBA 5 錯誤（「Invalid procedure call or argument」），整個 CmdUCINet 匯出中斷。殘破的 `.vna` 檔案會留在硬碟上 —— `*node data` 完整，`*node properties` 截斷，`*tie data` 完全沒寫。

使用者可見症狀：彈出 Run-time error 5 對話方塊；匯出的 `.vna` 檔案不完整，UCINET / Visone 無法讀取。

**受影響表單（按當前 dump 的證據）：**

- **LookAtAssociations** — 直接 canonicalize。由 `tests/test_known_bugs.py::test_bug22_associations_cmducinet_createtextfile_no_unicode_arg`（靜態）和 `tests/test_vba_bug_behaviors.py::test_bug22_associations_cmducinet_fires_invalid_procedure_call`（執行時）雙重釘死。原始調查證據見 `analysis/probe_associations_cmducinet_error5.md`。
- **LookAtKinship** — **同 root cause 的 runtime-confirmed sibling form。**`Form_LookAtKinship.CmdUCINet_Click` 在約 line 2510 用同樣的 `CreateTextFile(tFileName, True)` 2-arg 模式。透過 probe `investigate/kinship-cmducinet-sibling-risk`（commit 154bb4b）以 picker = pid 152930（He Jing 何淨，唯一 1-hop kin 是 pid 140733 He Mou 取，U+53D6 = 與 Associations 的 稜 = U+7A1C 同屬 CJK Han ideograph 觸發類）復現。Probe 結果：同樣的 `:ERR Invalid procedure call or argument`，同樣的殘破檔案形狀（`*node data` 完整 + `*node properties` 截斷 + `*tie data` 完全沒寫）。本 PR 已擴充套件靜態 marker `tests/test_known_bugs.py::test_bug22_associations_cmducinet_createtextfile_no_unicode_arg`，讓它同時檢查 `Form_LookAtKinship.vb` 的同樣 2-arg 模式；Kinship 的執行時 pin 暫緩（見下方 Coverage caveat）。
- **LookAtPlace** — 可能存在的獨立風險；**本 issue 的確認範圍不包含 Place。**`Form_LookAtPlace.CmdUCINet_Click` 用的是 ADO Stream（`tStream.WriteText`），不是 FSO （`tVNA.WriteLine`），編碼行為可能不同，需要單獨的 per-form probe 才能下同 bug-family 的結論。Place CmdUCINet 在 inventory 仍維持 `gap`。

**Coverage caveat：** 現有的 Kinship × CmdUCINet 覆蓋測試（`tests/test_vba_cmducinet_kinship.py`）在 inventory 上仍是 `covered`，但已 **明確標註為 fixture-fragile** —— 它能通過只是因為 matrix 提供的 person 3211 網路剛好沒有 Han 字元 c_name。換成一個網路能觸達Han 名字的 fixture（sibling probe 直接示範了這一點）就會在同一段 .vna 寫出路徑上崩潰。已在測試的 docstring 與 inventory manifest 的 notes 欄位同步備註。

#### 復現步驟

1. 在 Microsoft Access 裡開啟 CBDB_BJ_User.mdb。
2. 開啟 **LookAtAssociations** 表單（F11 → 導航窗格 → 表單 → 雙擊 `LookAtAssociations`）。
3. 用關聯程式碼 picker（**CmdPickAssoc**）選取 **c_assoc_code = 437（贈詩、文）** —— 在當前 dump 上，該程式碼的查詢結果包含 person 445395（c_name = `Hu Fa稜`），其 c_name 含 CJK 漢字（U+7A1C 稜），在 cp1252 碼頁下無法編碼也無 FSO 替代對映。
4. 點 **Run**（CmdQuery），等它把 ZZ_SCRATCH_ASSOC 和 ZZ_SCRATCH_P_ASSOC 填好。
5. 點 **UCINet** 匯出按鈕（CmdUCINet），隨便選一個 `.vna` 檔案的存檔位置。
6. 彈出對話方塊：`Run-time error '5': Invalid procedure call or argument`。匯出中斷，硬碟上只剩殘破的 `.vna` 檔：`*node data` 區段完整，`*node properties` 區段被截斷，`*tie data` 區段完全沒寫 —— UCINET / Visone 都沒法當成 import 檔案使用。
7. 已透過 `analysis/probe_associations_cmducinet_error5.py` 端到端驗證 —— 約 15 秒可穩定復現（完整證據鏈與崩潰定位見 `analysis/probe_associations_cmducinet_error5.md`）。

#### 建議修復方案

在 `CreateTextFile` 第 3 個引數加上 `True`，讓檔案以 Unicode（UTF-16LE）模式開啟：

```vb
' 修改前（Form_LookAtAssociations.vb:2575）
Set tVNA = tFileSystem.CreateTextFile(tFileName, True)

' 修改後 —— 第 3 個引數 = Unicode = True
Set tVNA = tFileSystem.CreateTextFile(tFileName, True, True)
```

這樣 `tVNA.WriteLine` 應能用 UTF-16LE 寫入所有字元，避免 cp1252 寫出時在非 cp1252 c_name 上的崩潰。下游 UCINET / Visone 對 UTF-16 `.vna` 的相容性 **不在本 PR 證據範圍內**，仍應在修補版上再驗證一次，才能宣告本 issue 關閉。

替代方案（不太推薦）：在 `WriteLine` 之前把 `c_name` 裡的非 cp1252 字元剝掉或轉寫。會丟失匯出的真實資料，而且程式碼量更大；Unicode flag 才是正解。

`Form_LookAtKinship.CmdUCINet_Click`（約 line 2510）也需要套用同樣的一行修改 —— 詳見上方「受影響表單」段。Kinship 是本 issue 的 runtime-confirmed sibling form（同 root cause、同 failure class，只是宿主表單與觸發 fixture 不同），所以一次上游修補應同時給兩個 CreateTextFile 加上 Unicode flag。Place （LookAtPlace.CmdUCINet）**不在** 本 issue 的範圍內 —— 它用的是 ADO Stream 而非 FSO，需要單獨的 per-form probe 才能下同 bug-family 的結論或加入修補範圍。

### Issue #23 — LookAtAssociations.CmdNeo4j 的 INSERT 引用了 ZZ_SCRATCH_PEOPLE 上不存在的目標列 c_index_addr_type_code（疑似本意為 c_addr_type）

**涉及位置:** `Form_LookAtAssociations.CmdNeo4j_Click`

**嚴重等級:** P1 — 正常使用者點選下的可見報錯（只要 LookAtAssociations 查詢有非空結果，CmdNeo4j 匯出就會觸發；JET 3061 在 matrix Associations fixture 上穩定復現，runtime 證據在 PR #112 的 probe，靜態 schema 證據在 PR #114 的調查）

#### 問題描述

`Form_LookAtAssociations.CmdNeo4j_Click` 通過 `INSERT INTO ZZ_SCRATCH_PEOPLE ( c_person_id, c_name, c_name_chn, c_index_year, c_index_year_type_code, c_dy, c_addr_id, c_index_addr_type_code, c_female ) SELECT DISTINCT … BIOG_MAIN.c_index_addr_type_code, BIOG_MAIN.c_female …` 來填充 `ZZ_SCRATCH_PEOPLE` 工作表。INSERT 目標列裡把 **c_index_addr_type_code** 當作 `ZZ_SCRATCH_PEOPLE` 上的列引用，但當前 dump 的 `ZZ_SCRATCH_PEOPLE` 共 22 列，並沒有這一列（canonical 列名見 `analysis/dump/tables.json`）。JET 報 **3061「INSERT INTO 語句包含未知的欄位名 c_index_addr_type_code」**，整個 Neo4j 匯出在 body 中途中斷 —— 在任何磁碟檔案被寫出之前 —— 即便 `dlgSaveAs.Show` 已經彈出且鏈條已進入 People-block True 分支。

靜態 schema 互相印證：source 端 `BIOG_MAIN` **有** `c_index_addr_type_code`（共 55 列；該列在 `tests/test_schema.py` 的 REQUIRED_COLUMNS 中，schema 測試通過即獨立證實）。所以缺的是 **target** 表，不是 source。

關於作者意圖的強靜態推斷：失敗 INSERT 之後緊鄰的 `UPDATE` 是 LEFT JOIN `ZZ_SCRATCH_PEOPLE.c_addr_type = BIOG_ADDR_CODES.c_addr_type` 並 SET 一組地址描述列。target 表 **有** `c_addr_type`，但失敗的 INSERT 從未寫入 `c_addr_type` —— 而 `c_addr_type` 正是 source `BIOG_MAIN.c_index_addr_type_code` 的自然 rename 目標（同一 INSERT 在前一列已經做過一模一樣的 rename：`BIOG_MAIN.c_index_addr_id ↦ ZZ_SCRATCH_PEOPLE.c_addr_id`）。看起來作者把 source 列名直接複製到 INSERT 目標列表裡，本意是 rename 成 `c_addr_type`。與 Bug #4 / #5 / #6 同形（per-form column-name typo class）。

**與 Issue #22（LookAtAssociations × CmdUCINet）無關。** Issue #22 是 FSO `CreateTextFile` 缺 Unicode 引數導致 CJK 漢字走 ANSI cp1252 時崩潰；本 issue 是 `CmdNeo4j_Click` 裡 SQL JET 3061 列不存在，在任何 export-encoder 執行之前就觸發。不同 sub、不同 VBA 錯誤族、不同修法。

**與 AssociationPairs × CmdNeo4j 的 blocking-MsgBox 層不同**（canonical 證據：main 上 PR #109 的 driver patch + PR #110 的 coverage）。AssocPairs 的 CmdNeo4j 會先寫出 ≥1 個檔案再被 UI debug-MsgBox 層擋住（現已由 driver suppress）。Associations 的 CmdNeo4j 寫 0 個檔案就被上面的 SQL schema 不匹配擋住。兩者鏈條深度不同、VBA 錯誤族不同、修法路徑無重疊 —— 不應合併為同一 issue。

#### 復現步驟

1. 開啟 **LookAtAssociations**。
2. 選一個有實質資料的 **association code**（任一關聯人數上百的 code 都可以；matrix `_make_assoc_fixtures` 第一條 fixture `assoc_<top_code>_unfiltered` 在當前 dump 上能讓 ZZ_SCRATCH_ASSOC 約 11,867 行、ZZ_SCRATCH_P_ASSOC 約 8,087 行 —— 同 Associations × CmdGIS / CmdPajek / CmdGephi 測試用的是同一條 fixture）。
3. **FrameFilterYears** 保持預設（不篩年）。不要勾任何與本次查詢無關的子分支。
4. 點 **Run Query** —— CmdQuery 順利完成，底層 scratch 表被填充。
5. 點 **Neo4j**（匯出按鈕）。
6. 彈出 **執行時錯誤 3061 —— INSERT INTO 語句包含未知的欄位名 c_index_addr_type_code** 對話方塊（或 headless ／driver-instrumented 跑法下，對應的 `LookAtAssociations:ERR ...` ZZ_TEST_DEBUG marker 被寫入）。Neo4j 匯出產出 **0 份 CSV** —— 沒有 `People_*.csv`、沒有 `Places_*.csv`，什麼都沒有。

已通過 probe `analysis/probe_associations_cmdneo4j.py` 端到端驗證（PR #112，已 merge `1145219`）；根因經靜態調查 `analysis/investigate_associations_cmdneo4j_c_index_addr_type_code.{py,md}` 確認（PR #114，已 merge `68cfa9b`）。

#### 建議修復方案

**推薦的上游 CBDB 修復：** 把 INSERT 目標列名從 `c_index_addr_type_code` 改為 `c_addr_type`。緊鄰的 UPDATE 本來就在 `ZZ_SCRATCH_PEOPLE.c_addr_type = BIOG_ADDR_CODES.c_addr_type` 上 LEFT JOIN，這個 rename 順帶還填上 UPDATE 早就需要的 join key（今天該 join key 從未被填充，所以即使只是「把那一列刪掉」的修法也會讓 UPDATE 靜默地在 NULL 上 join）。

建議寫法（只改一個 identifier；SELECT 端不變，因為 source 列名是對的）：

```vb
tQueryStr = "INSERT INTO ZZ_SCRATCH_PEOPLE ( " & _
    "c_person_id, c_name, c_name_chn, c_index_year, " & _
    "c_index_year_type_code, c_dy, c_addr_id, " & _
    "c_addr_type, c_female ) " & _
    "SELECT DISTINCT ZZ_SCRATCH_P_TEXT.c_person_id, BIOG_MAIN.c_name, ... " & _
    "BIOG_MAIN.c_dy, BIOG_MAIN.c_index_addr_id, " & _
    "BIOG_MAIN.c_index_addr_type_code, BIOG_MAIN.c_female " & _
    "FROM ZZ_SCRATCH_P_TEXT INNER JOIN BIOG_MAIN ON ..."
```

**Driver-side workaround 選項**（另一份 brief；本 issue 不包含）：仿照 `_PER_FORM_CMDGIS_PATCHES` 裡 Issue #4（`GISFrame → CodeFrame`）和 Issue #5（`ChkIDs → False`）的寫法，在僅限於 `Form_LookAtAssociations` 的 CmdNeo4j_Click 內做一個把字面量 `c_index_addr_type_code` 改寫為 `c_addr_type` 的 per-form rewrite。這能讓測試套不依賴上游修復就跑通。

### Issue #24 — LookAtPlace.CmdNeo4j 中 tRstPeople 的 SELECT 投影缺少 c_dynasty / c_dynasty_chn / c_female，下游迴圈仍然讀取這些欄位並崩潰報「Item not found in this collection」（JET 3265）

**涉及位置:** `Form_LookAtPlace.CmdNeo4j_Click`

**嚴重等級:** P1 — 正常使用者點選下的可見報錯（只要 LookAtPlace 查詢有非空的 place-people 結果，CmdNeo4j 匯出就會觸發；JET 3265 在第 757 行 `!c_dynasty` 讀取上穩定復現，runtime 證據在 PR #120 的 probe，靜態調查在 PR #121）

#### 問題描述

`Form_LookAtPlace.CmdNeo4j_Click` 在第 651 行用 `Set tRstPeople = CurrentDb.OpenRecordset(tQueryStr, dbOpenDynaset)` 開啟一個 recordset。繫結的 SQL（第 643-647 行）只從 `ZZ_SCRATCH_P_TEXT` 投影 **4** 個欄位：

```
SELECT DISTINCT
    ZZ_SCRATCH_P_TEXT.c_person_id,
    ZZ_SCRATCH_P_TEXT.c_name,
    ZZ_SCRATCH_P_TEXT.c_name_chn,
    ZZ_SCRATCH_P_TEXT.c_index_year
FROM ZZ_SCRATCH_P_TEXT INNER JOIN
     ( DYNASTIES RIGHT JOIN BIOG_MAIN ON
       DYNASTIES.c_dy = BIOG_MAIN.c_dy )
ON ZZ_SCRATCH_P_TEXT.c_person_id = BIOG_MAIN.c_personid
```

`INNER JOIN` 把 `DYNASTIES` 和 `BIOG_MAIN` 帶入範圍，但 SELECT 子句沒投影這兩個表的任何欄位。DAO 的 `Recordset.Fields` 集合只包含 SELECT 投影的欄位；JOIN 用於 filtering / row-shaping，**不**用於欄位存取。

儘管如此，迴圈本體（第 689-783 行）仍然讀取 **3 個來自 JOINed 表但未被 SELECT 投影的欄位**：

  - `tRstPeople!c_dynasty`     （第 757 行；預期來自 `DYNASTIES`）
  - `tRstPeople!c_dynasty_chn` （第 769 行；預期來自 `DYNASTIES`）
  - `tRstPeople!c_female`      （第 777、783 行；預期來自 `BIOG_MAIN`）

JET 在第一次此類讀取時報 **3265「Item not found in this collection」**（即第 757 行 `!c_dynasty`）。錯誤處理跳到 `Exit_CmdNeo4j_Click`，**在 `gStream.WriteText` 把任何資料寫入磁碟之前**就退出，所以使用者看到彈窗且匯出產生 **0 個 CSV** —— 即使第 545 行的 `dlgSaveAs.Show` 已經觸發並捕獲了檔名。

靜態 schema 互相印證（基於 `analysis/dump/tables.json`）：`DYNASTIES` **有** `c_dynasty` 與 `c_dynasty_chn`；`BIOG_MAIN` **有** `c_female`（也在 `tests/test_schema.py::REQUIRED_COLUMNS` 之中）。所以這 **不是** source-side 欄位 rename / removal —— 欄位在源表都在，SELECT 只是沒投影。純粹是 recordset projection mismatch。

**與 Issue #23（LookAtAssociations × CmdNeo4j target-column mismatch）不同。** Issue #23 是 JET 3061（SQL parser）：INSERT 目標欄位列表引用了不存在的目標表欄位。本 issue 是 JET 3265（DAO 欄位集合查詢）：VBA 層的 `Recordset!field` 引用在執行期欄位集合中找不到對應欄位。相同的表面根因類別（缺欄位引用）但**不同的觸發面**（DAO 欄位查詢 vs SQL parser）和**不同的修法面**（SELECT 投影 vs INSERT 目標列表）。不應併入 Issue #23。

**與 Issue #21（LookAtGroupData × CmdNeo4j 未保護的 `.MoveFirst`）也不同。** Issue #21 是 DAO 3021「No current record」—— 對空但有效的 recordset 的狀態機錯誤。本 issue 是對非空但欄位集合中不存在某欄位的 recordset 做欄位存取的錯誤。不同 DAO 錯誤碼、不同修法路徑。

#### 復現步驟

1. 開啟 **LookAtPlace**。
2. 在地址挑選器選一個有實質資料的 **address**（任何關聯了大量傳記資料的 c_addr_id；matrix `_make_place_fixtures` 第一條 fixture `place_addr_<top_addr_by_indexed_persons>` 在當前 dump 上能讓 ZZ_SCRATCH_PLACE_PEOPLE ≈ 5,764 行 —— 與現有 Place × CmdGIS / CmdPajek 測試共用同一條 fixture）。
3. 在表單上將 tab 設為 **Places**（TabPlaces=0；當前 cross-form CmdNeo4j 測試在 `_seed_query_inputs` 中已處理）。勾 **ChkIndividual**，不勾 **ChkOffice / ChkAssoc / ChkPosting / ChkEntry**。
4. 點 **Run Query** —— CmdQuery 順利完成；scratch 表都被填充（ZZ_SCRATCH_PEOPLE、ZZ_SCRATCH_PLACE_PEOPLE、ZZ_SCRATCH_PLACE_AGG 都非空）。
5. 點 **Neo4j**（匯出按鈕）。
6. 彈出 **執行時錯誤 3265 —— Item not found in this collection.** 對話方塊（或 headless ／ driver-instrumented 跑法下，對應的 `LookAtPlace:ERR Item not found in this collection.` ZZ_TEST_DEBUG marker 被寫入）。Neo4j 匯出產出 **0 份 CSV** —— 沒有 `People_*.csv`、沒有 `Places_*.csv`，什麼都沒有。

已通過 probe `analysis/probe_place_cmdneo4j.py` 端到端驗證（PR #120，已 merge `8f94276`）；具體的失敗引用點由靜態調查 `analysis/investigate_place_cmdneo4j_item_not_found.{py,md}` 確認（PR #121，已 merge `97e1162`）。

#### 建議修復方案

**推薦的上游 CBDB 修復：** 擴充套件 `Form_LookAtPlace.vb:643-647` 的 SELECT 投影，包含迴圈讀取的 3 個欄位：

```vb
tQueryStr = "SELECT DISTINCT " & _
    "ZZ_SCRATCH_P_TEXT.c_person_id, " & _
    "ZZ_SCRATCH_P_TEXT.c_name, " & _
    "ZZ_SCRATCH_P_TEXT.c_name_chn, " & _
    "ZZ_SCRATCH_P_TEXT.c_index_year, " & _
    "DYNASTIES.c_dynasty, " & _
    "DYNASTIES.c_dynasty_chn, " & _
    "BIOG_MAIN.c_female " & _
    "FROM ZZ_SCRATCH_P_TEXT INNER JOIN " & _
    "( DYNASTIES RIGHT JOIN BIOG_MAIN ON " & _
    "DYNASTIES.c_dy = BIOG_MAIN.c_dy ) " & _
    "ON ZZ_SCRATCH_P_TEXT.c_person_id = BIOG_MAIN.c_personid"
```

FROM / JOIN 結構已經把源表帶進範圍；這純粹是 SELECT 缺欄位的問題。增加 3 個欄位，其他不變。

**Driver-side workaround 選項**（另一份 brief；本 issue 不包含）：仿照 `_PER_FORM_CMDGIS_PATCHES` 裡 Issue #4（`GISFrame -> CodeFrame`）、Issue #5（`ChkIDs -> False`）和 Issue #23（`c_index_addr_type_code -> c_addr_type` INSERT 目標重寫）的寫法，在僅限於 `Form_LookAtPlace` 的 CmdNeo4j_Click 內做一個擴充套件 SELECT 投影字面量的 per-form rewrite。這能讓測試套不依賴上游修復就跑通。

## P2 — 靜默顯示問題

### Issue #10 — EVENT_ADDR_2 子表單的地址列默默地顯示為空（ControlSource 寫錯了）

**涉及位置:** `EVENT_ADDR_2 Subform`

**嚴重等級:** P2 — 靜默顯示問題（EVENT_ADDR_2 的 TxtAddrCHN / TxtAddrPY 每一列都空白）

#### 問題描述

在 EVENT_ADDR_2 子表單（帶地址的事件）上，兩個地址控制元件的繫結如下：

  • `TxtAddrCHN`.ControlSource = `c_name_chn`
  • `TxtAddrPY`.ControlSource = `c_name`

但該表單的 RecordSource 是存檔查詢 `View_EventAddrData`，裡面把 ADDR_CODES.c_name_chn 起別名成 `c_event_addr_chn`、把 ADDR_CODES.c_name 起別名成 `c_event_addr_name`。投影裡既沒有 `c_name` 也沒有 `c_name_chn`，所以這兩個控制元件每一行都默默地顯示為空。

#### 復現步驟

**建議使用的範例人物：** `c_personid=44872`（孫才，Sun Cai）

開啟人物 44872（孫才，Sun Cai）。EVENTS 子資料表會顯示 1 條事件，其中 1 條有對應地址。相關繫結控制元件在每一列都會顯示空白。 _由 `reports/probe_demo_persons.py` 透過 SQL probe 挑選；之所以選這位，是因為其底層記錄數確實滿足這個 bug 的觸發條件。_

1. 開啟 CBDB_Browser_2，導航到 **c_personid = 44872（孫才 Sun Cai）**——選他是因為他有 1 條 EVENTS_DATA 記錄，對應 1 條 EVENT_ADDR 指向 `c_addr_id = 12603`（ADDR_CODES 裡是 Anfeng / 安豐）。切到 **Events** 子分頁。
2. 看事件那行內嵌的 EVENT_ADDR_2 子表單（在主事件那行下方的一小條）。那裡的兩個位址列位 `TxtAddrCHN` 與 `TxtAddrPY` 都是空白。
3. **注意：**外層的 EVENTS_DATA_2 子表單也有自己的位址列位（也叫 TxtAddrCHN / TxtAddrPY，但是綁到 `c_addr_chn` / `c_addr_name`，這兩個 `View_EventsData` 確實有 project）——這兩個欄位是正常的，會顯示「安豐 / Anfeng」。Bug #10 講的是內層 EVENT_ADDR_2 那兩個空欄位，不是父層那條看得見的地址值。
4. SQL 驗證（不需開 Access）：`SELECT c_name_chn FROM View_EventAddrData` 會拋 `Too few parameters. Expected 2.`——JET 把未知識別字當作引數對待，這就確認了該欄位不在投影裡。

#### 截圖

![bug10_subform_annotated.png](screenshots/bug10_subform_annotated.png)

_Runtime view of CBDB_Browser_2 → BIOG_MAIN_2 → Events tab with c_personid=44872 (孫才) loaded.  **The visible '安豐' / address values come from the parent EVENTS_DATA_2 sub-form's TxtAddrCHN (correctly bound to `c_addr_chn`).**  Bug #10's blank controls live in the smaller EVENT_ADDR_2 sub-form nested inside the event row — those two controls (TxtAddrCHN / TxtAddrPY bound to `c_name_chn` / `c_name`, neither in `View_EventAddrData`'s projection) render empty.  COM probe confirms both are Visible=True with widths 2340 / 2100 twips (≈4cm / 3.5cm) — i.e. real user-visible blank columns, just smaller than the parent row's address display.  Verification scripts: `analysis/probe_bug_10_11_12_visibility.py` (control visibility) + the SQL probe in the steps above._

#### 建議修復方案

在 `EVENT_ADDR_2 Subform` 的表單設計檢視裡，把 `TxtAddrCHN`.ControlSource 由 `c_name_chn` 改成 `c_event_addr_chn`；把 `TxtAddrPY`.ControlSource 由 `c_name` 改成 `c_event_addr_name`（這才是 View_EventAddrData 裡真實的別名）。

## P3 — 缺失介面

### Issue #15 — LookAtPlace 缺少 CmdGIS 按鈕（程式碼裡有 handler 但介面上沒控制元件）

**涉及位置:** `LookAtPlace`

**嚴重等級:** P3 — 缺失介面（使用者用不到該功能）

#### 問題描述

`Form_LookAtPlace.vb` 裡定義了一個完整可用的 `CmdGIS_Click` handler——構造並輸出 GIS .tab 檔案，邏輯和 Status / Texts / Associations / Office / Kinship 上的 GIS 按鈕一模一樣。但 LookAtPlace 的設計檢視上根本沒有 `CmdGIS` 按鈕。使用者在 Place 上可以使用 Pajek / Gephi / Neo4j 匯出，但用不了 GIS 匯出——程式碼在那裡，只是介面進不去。

#### 復現步驟

1. 開啟 **LookAtPlace**。
2. 看右下方那一排匯出按鈕。
3. 沒有 GIS 按鈕。可以對比 LookAtStatus / LookAtAssociations / LookAtOffice 等，它們都有這個按鈕。

#### 截圖

![bug15_LookAtPlace_no_CmdGIS_annotated.png](screenshots/bug15_LookAtPlace_no_CmdGIS_annotated.png)

_LookAtPlace as it ships — no GIS button is rendered, even though `Sub CmdGIS_Click()` exists in the module._

#### 建議修復方案

在 LookAtPlace 的設計視圖裡，在已有的 CmdPajek / CmdGephi 旁邊加一個 CmdGIS 按鈕，把 `OnClick` 設為 `[Event Procedure]`，這樣它就會呼叫已經寫好的 CmdGIS_Click。（同時務必先修 Issue #4，否則按鈕一點就會報 Object required。）

### Issue #16 — LookAtStatus 缺少 CmdPajek 按鈕

**涉及位置:** `LookAtStatus`

**嚴重等級:** P3 — 缺失介面

#### 問題描述

形態與 Issue #15 相同。`Sub CmdPajek_Click()` 在 `Form_LookAtStatus.vb` 裡有定義（本應輸出 Pajek .net 檔案），但 Status 的設計檢視上沒有 CmdPajek 按鈕。

注意：即便把按鈕加回去，也得先解決 Issue #5（CmdPajek_Click 本身的 SQL/控制元件缺陷）。

#### 復現步驟

1. 開啟 **LookAtStatus**。匯出按鈕一欄只有 GIS 和 Neo4j，沒有 Pajek。

#### 截圖

![bug16_LookAtStatus_no_CmdPajek_annotated.png](screenshots/bug16_LookAtStatus_no_CmdPajek_annotated.png)

#### 建議修復方案

在 LookAtStatus 的設計視圖裡加一個 CmdPajek 按鈕（先把 Issue #5 修好）。

### Issue #17 — LookAtStatus 缺少 CmdGephi 按鈕

**涉及位置:** `LookAtStatus`

**嚴重等級:** P3 — 缺失介面

#### 問題描述

`Sub CmdGephi_Click()` 在 `Form_LookAtStatus.vb` 裡有定義，但表單設計視圖裡沒有相應按鈕。

#### 復現步驟

1. 開啟 **LookAtStatus**。沒有 Gephi 匯出按鈕。

#### 截圖

![bug17_LookAtStatus_no_CmdGephi_annotated.png](screenshots/bug17_LookAtStatus_no_CmdGephi_annotated.png)

#### 建議修復方案

在 LookAtStatus 的設計視圖裡加一個 CmdGephi 按鈕。

### Issue #18 — LookAtStatus 缺少 CmdUCINet 按鈕

**涉及位置:** `LookAtStatus`

**嚴重等級:** P3 — 缺失介面

#### 問題描述

`Sub CmdUCINet_Click()` 在 `Form_LookAtStatus.vb` 裡有定義，但表單設計視圖裡沒有相應按鈕。

#### 復現步驟

1. 開啟 **LookAtStatus**。沒有 UCINet 匯出按鈕。

#### 截圖

![bug18_LookAtStatus_no_CmdUCINet_annotated.png](screenshots/bug18_LookAtStatus_no_CmdUCINet_annotated.png)

#### 建議修復方案

在 LookAtStatus 的設計視圖裡加一個 CmdUCINet 按鈕。

### Issue #19 — LookAtOffice 缺少 CmdGUESS 按鈕

**涉及位置:** `LookAtOffice`

**嚴重等級:** P3 — 缺失介面

#### 問題描述

`Sub CmdGUESS_Click()` 在 `Form_LookAtOffice.vb` 裡有定義，但 Office 的設計檢視上沒有 CmdGUESS 按鈕。Office 上的使用者可以使用 GIS / GISPeople / Neo4j 匯出，但用不了 GUESS。

#### 復現步驟

1. 開啟 **LookAtOffice**。沒有 GUESS 匯出按鈕。

#### 截圖

![bug19_LookAtOffice_no_CmdGUESS_annotated.png](screenshots/bug19_LookAtOffice_no_CmdGUESS_annotated.png)

#### 建議修復方案

在 LookAtOffice 的設計視圖裡加一個 CmdGUESS 按鈕。

## P4 — 安裝設定

### Issue #2 — VBA 工程引用了過時的 dao360.dll，Office 2016+ 機器上沒這個檔案

**涉及位置:** `VBE Project References`

**嚴重等級:** P4 — 每臺新機器一次性的安裝障礙

#### 問題描述

.mdb 中的 VBA 工程硬性引用了 `C:\Program Files\Common Files\Microsoft Shared\DAO\dao360.dll`，這是 Access 2003 時代 DAO 3.6 的位置。現代 Office（2016 起）改用 `ACEDAO.DLL`，並不會安裝舊版 DLL。在任何全新的現代機器上，第一次嘗試開啟任意 LookAt 表單都會報「找不到工程或庫」（Can't find project or library），對終端使用者來說既看不懂又嚇人。

嚴重等級較低，因為每臺機器只需修一次，但每臺新裝都會撞上。

#### 復現步驟

1. 在全新的現代 Office 機器上安裝 `CBDB_BJ_User.mdb`。
2. 開啟檔案，按 Alt+F11 進入 VBE 編輯器。
3. 工具 → 引用。可以看到一行寫著 `MISSING: dao360.dll`。
4. 開啟任意 LookAt 表單。會彈出「Can't find project or library」錯誤。

#### 建議修復方案

在維護者的機器上做一次：
  1. 用 Access 開啟 .mdb，按 Alt+F11。
  2. 工具 → 引用。取消勾選標著 MISSING 的 dao360.dll。
  3. 勾選 `Microsoft Office 16.0 Access Database Engine Object Library`（即 ACEDAO.DLL）。
  4. 儲存 .mdb。

然後重新分發修好的檔案。以後的終端使用者什麼都不用做。

## P5 — 潛伏 / 不可達 / 當前無法復現

_本層的條目作為歷史 / 潛伏記錄保留。可分為三類：(a) DORMANT 潛伏 — 已驗證當前源資料無法觸發該症狀；(b) 當前無法復現 — 症狀不再出現，但可疑程式碼仍在（我們**沒有**確認上游有原始碼層面的修復；原因可能是 JET / Office 的行為改變、可能是我們這邊 fixture/driver 改變，也可能原本的診斷就是 false positive）；(c) LATENT 被遮蔽 — 原始碼缺陷確實存在，但因為另一個 issue（例如某個 UI 按鈕缺失）擋住了使用路徑，使用者目前碰不到。本層條目當下都不是使用者會遇到的問題，**也沒有任何一條被確認上游修復**；若要當成緊急或已關閉處理，請先諮詢。_

### Issue #1 — View_StatusData 會把首年份範圍顯示成末年份範圍 — DORMANT（當前 dump 沒有源資料能觸發）

**涉及位置:** `View_StatusData`

**嚴重等級:** P5 — 在當前 dump 上潛伏（若任何 STATUS_DATA 列同時填了 fy/ly range 且不同，會升為 P0）

#### 問題描述

存檔查詢 `View_StatusData` 把 `YEAR_RANGE_CODES` 表 JOIN 了兩次（其中一次別名是 `YEAR_RANGE_CODES_1`，用於末年份範圍），但 SELECT 列表裡所有範圍欄位都從 _1 別名取值。結果是 Status 子資料表裡每一行顯示的「首年份範圍」其實是末年份範圍。

#### 復現步驟

⚠ **在當前資料快照下無法在 UI 上復現——請看下方說明**

在當前資料快照下，這個 bug 處於 **潛伏 (dormant) 狀態**——STATUS_DATA 共 70,761 行，但只有 13 行 c_fy_range > 0、0 行 c_ly_range > 0；兩個都有值且不同的只有 0 行。所以目前沒有任何人物能在 UI 上重現這個別名錯位。SQL 缺陷仍然存在；只要未來某次資料更新插入一條 fy/ly range 都填了且不同的 STATUS_DATA 記錄，對應的子資料表那一行就會顯示錯誤文字。今天若要驗證這個 bug，可以直接跑 SQL：
  SELECT c_personid, c_fy_range, c_fy_range_desc, c_ly_range, c_ly_range_desc FROM View_StatusData WHERE c_fy_range > 0 OR c_ly_range > 0;

1. 由於本 .mdb 當前快照下，沒有任何 STATUS_DATA 列同時填了 c_fy_range 和 c_ly_range，這個 bug 暫時無法在 UI 上復現。請直接用 SQL 驗證：
2. 在 Access 裡開啟 .mdb，按 F11 顯示導航窗格，雙擊查詢 **View_StatusData**。
3. 檢視 SELECT 子句：所有 `c_fy_range_*` 別名都從 `YEAR_RANGE_CODES_1` 取值，但 FROM 子句把這個別名 JOIN 在末年份範圍上——這就是錯位。
4. （可選）在 Access 查詢視窗執行 `SELECT TOP 100 c_personid, c_fy_range, c_fy_range_desc, c_ly_range, c_ly_range_desc FROM View_StatusData WHERE c_fy_range > 0 OR c_ly_range > 0` ——未來某次資料更新如果同時填了這兩個欄位且取值不同，每一條結果都會顯示錯誤的首年份文字。

#### 建議修復方案

在 `View_StatusData` 中，把 `YEAR_RANGE_CODES_1.c_range AS c_fy_range_desc` 和 `YEAR_RANGE_CODES_1.c_range_chn AS c_fy_range_chn` 改成不帶別名的 `YEAR_RANGE_CODES.*`（FROM 子句已經按 `c_fy_range` JOIN 了它）。

### Issue #9 — LookAtEntry.CmdNeo4j 的機構 (Institutions) 部分用錯了記錄集變數 — LATENT（被資料 gate 跳過、不可達；當前 dump 中沒有 c_inst_code > 0 的 ENTRY_DATA 行）

**涉及位置:** `Form_LookAtEntry.CmdNeo4j_Click`

**嚴重等級:** P5 — Source-level latent typo（在當前 dump 被 gate 跳過、不可達；若未來任一 ENTRY_DATA 列出現 c_inst_code > 0，會回歸為 P1）

#### 問題描述

**Source-level typo，目前在此 dump 上不可達。**

`Form_LookAtEntry.vb` 第 1415 行用 `Set tRstInstitutions = CurrentDb.OpenRecordset(tQueryStr)` 開啟 institutions 記錄集。十行之後，第 1425 行卻寫成 `With tRstAssocCodes`，接下來迴圈讀 `!c_inst_code`、`!c_inst_name_code` 等欄位 ── 這個 recordset 早在 AssocCodes 區塊的第 1373 行就已經 `Close` 掉了。一旦真的執行到，`.MoveFirst` 那行就會丟出 DAO 3021（「No current record」）。這個錯名 reference 確實是 source-level bug。

但在當前 dump，整個 SaveAs 對話方塊與有 bug 的 `With` 區塊都包在第 1389 行的 gate `If tRecDeleted > 0 Then` 裡；`tRecDeleted` 是緊鄰的 `INSERT INTO ZZ_SCRATCH_P_TEXT … WHERE ZZ_SCRATCH_ENTRY.c_inst_code > 0` 寫入列數。CmdQuery 把 `ENTRY_DATA.c_inst_code` 原樣複製到 `ZZ_SCRATCH_ENTRY.c_inst_code`（第 1645-1652 行），而 **當前 MDB 263,454 筆 ENTRY_DATA 中，`c_inst_code > 0` 的列數為 0**（`c_inst_name_code > 0` 也為 0）。因此任何 LookAtEntry 條件下 `tRecDeleted` 都會是 0、gate 永遠為假、SaveAs 對話方塊不會出現、`With tRstAssocCodes` 永遠不會被執行，CmdNeo4j 會順利走完並顯示「Finished saving to Neo4j」── 只是因為沒有 institution rows，所以靜默地省略 `InstitutionCodes_*.csv`。

**缺 `InstitutionCodes_*.csv` 本身在當前 dump 不是使用者可見錯誤**。當該分割槽資料表計數為 0 時跳過該 optional 檔，與鄰近區塊的 gating 行為完全一致（fixture 若 `c_assoc_code = 0`，AssocCodes 區塊同樣會被靜默跳過 ── 詳見 re-verification artifacts 的對照行為）。只有當未來某次 MDB 更新引入任一 `c_inst_code > 0` 的 ENTRY_DATA 列，這個 latent typo 才會變成使用者可見錯誤。

**Re-verification 證據：** SQL pre-image 與對 `c_entry_code = 36`（科舉：進士）、`c_entry_code = 101`（薦舉/保任）的真實 Access COM 探查均確認：無 popup、chain 順利完成、產出檔案中沒有 `InstitutionCode` 形狀的檔。細節請見 `analysis/issue9_neo4j_institutioncodes_reverification.md` 與 `reports/issue9_neo4j_institutioncodes_reverification.json`。

#### 復現步驟

1. **當前 dump 此 bug 無法從 UI 觸發** ── Form_LookAtEntry.vb:1389 的 `If tRecDeleted > 0 Then` gate 對任何 LookAtEntry 條件都為假（263,454 筆 ENTRY_DATA 中 `c_inst_code > 0` 為 0）。請改用 source-level 靜態驗證：
2. 開啟 `analysis/dump/vba/Form_LookAtEntry.vb`，讀第 1415-1425 行。第 1415 行寫 `Set tRstInstitutions = CurrentDb.OpenRecordset(tQueryStr)`，第 1425 行寫 `With tRstAssocCodes`（原意應為 `With tRstInstitutions`）。`tRstAssocCodes` 早在 AssocCodes 區塊的第 1373 行就已經 `Close` 掉，因此 `.MoveFirst` 一旦執行即丟 DAO 3021。
3. （可選，SQL）在當前 dump 確認 gate 條件：`SELECT COUNT(*) FROM ENTRY_DATA WHERE c_inst_code > 0` 結果為 0（`c_inst_name_code > 0` 同樣為 0）。正是這點讓有 bug 的區塊變成不可達。
4. （可選，runtime 證據）在 LookAtEntry 選 `c_entry_code = 36`（科舉：進士）或 `c_entry_code = 101`（薦舉/保任）→ Run Query → Neo4j。chain 會順利完成並顯示「Finished saving to Neo4j」；無 popup，輸出資料夾中沒有 `InstitutionCodes_*.csv`。這在今天不是使用者可見錯誤 ── 它是 gate 確實起作用的證據。
5. Source-level typo 會在未來某次 MDB 更新引入任一 `c_inst_code > 0` 的 ENTRY_DATA 列時變成使用者可見錯誤。屆時 gate 開啟，`With tRstAssocCodes` 那行對已 `Close` 的 recordset 執行，`.MoveFirst` 立即丟出 DAO 3021（「No current record」）。

#### 具體復現

下列兩個 `c_entry_code` 值是 re-verification 用的investigation 證據 ── 它們會走完 CmdQuery + CmdNeo4j，用來示範 InstitutionCodes branch 在當前 dump 被 gate 跳過。**這兩個 fixture 不是 popup 復現**，兩者都會顯示「Finished saving to Neo4j」、無錯誤、無 `InstitutionCodes_*.csv`：

  - `c_entry_code = 36`（科舉：進士）── ENTRY_DATA 92,514 筆，`c_inst_code > 0` 為 0
  - `c_entry_code = 101`（薦舉/保任）── ENTRY_DATA 878 筆，`c_inst_code > 0` 為 0

用 `python analysis/investigate_issue9_neo4j_institutioncodes.py` 重跑 SQL evidence；加 `--com` 會額外跑真實 Access COM。

#### 建議修復方案

把第 1425 行的 `With tRstAssocCodes` 改成 `With tRstInstitutions`。屬於一字之差的筆誤，底層記錄集變數只是寫錯了。雖然目前不可達，順手修掉成本極低，也能避免未來資料一旦變動就回歸成 user-visible bug。

### Issue #4 — LookAtPlace.CmdGIS 會報「Object required」 — LATENT，被 Issue #15（表單上沒有 CmdGIS 按鈕）所遮蔽

**涉及位置:** `Form_LookAtPlace.CmdGIS_Click`

**嚴重等級:** P5 — 潛伏（若先修了 Issue #15 而沒同時修本條，會變成 P1）

#### 問題描述

說明：在當前 .mdb 上這個問題暫時不會被使用者觸發，因為 LookAtPlace的設計視圖裡根本沒有 CmdGIS 按鈕（即 Issue #15）——使用者無法點選。但底層 VBA 問題依然存在：`Form_LookAtPlace.vb` 第 1539 行寫的是 `If GISFrame.Value = 1 Then`，而該表單上根本沒有 `GISFrame` 控制元件（真正的編碼選擇控制元件叫 `CodeFrame`）。一旦 Issue #15 裡把缺失的按鈕加回去而沒先修這一行，每一次點選都會拋錯。

#### 復現步驟

1. （在 Issue #15 修好之後才能復現）開啟 **LookAtPlace**。
2. 跑任意一次查詢。
3. 點 GIS 按鈕。
4. 彈出 `執行時錯誤 424 ——必要的物件（Object required）` 對話方塊，匯出什麼都沒做。

#### 截圖

![bug4_step3_faux_popup.png](screenshots/bug4_step3_faux_popup.png)

_**Hypothetical** popup, reconstructed in PIL.  Users currently CAN'T trigger this — Bug #15 means the CmdGIS button does not exist on LookAtPlace, so the click that would fire `CmdGIS_Click` (and produce this 'Object required' error) has nowhere to come from.  This image shows what the user would see if a future change restored the CmdGIS button without first fixing the GISFrame → CodeFrame typo on line 1539.  The earlier bug4 step1 / step2 runtime screenshots were misleading (their annotations implied a clickable GIS button) and were removed in PR C — only this faux popup is kept as latent-state evidence._

#### 建議修復方案

把 `Form_LookAtPlace.vb` 第 1539 行的 `GISFrame.Value` 改成 `CodeFrame.Value`。同表單的 `CmdNeo4j_Click`、`CmdGephi_Click`、`CmdPajek_Click` 已經寫對了，可以參考。

### Issue #5 — LookAtStatus.CmdPajek 引用了不存在的控制元件，且 SQL 用了三個不存在的列

**涉及位置:** `Form_LookAtStatus.CmdPajek_Click`

**嚴重等級:** P5 — 潛伏（若先修了 Issue #16 而沒同時修本條，會變成 P1）

#### 問題描述

同一個 handler 裡有兩個相關缺陷：

  (a) 第 2308 行寫 `If ChkIDs.Value Then`，但 Status 上沒有名為 `ChkIDs` 的控制元件——只有 `ChkXYRef`、`ChkKML`、`ChkSubUnits`。

  (b) 第 2335–2338 行構造的 SELECT 引用 `ZZ_SCRATCH_STATUS.c_person_id`、`c_status_id`、`c_status_count`——這三列都不在 schema 裡（真實列名是 `c_personid`、`c_status_code`，count 列根本沒有）。

整段 sub 看起來是從 `LookAtAssociations.CmdPajek_Click` 整段拷過來的，那邊列名都對得上；改名時這兩處都漏了。和 Issue #4 一樣，因為 LookAtStatus 當前也沒有 Pajek 按鈕（Issue #16），使用者暫時碰不到；但只要按鈕加回去而沒先修這兩處，使用者就會立刻看到錯誤。

#### 復現步驟

1. （在 Issue #16 修好之後才能復現）開啟 **LookAtStatus**。
2. 跑一次查詢，然後點 Pajek 按鈕。
3. 第一次會彈 `Object required`（ChkIDs 引用所致）。
4. 如果繞過它，下一次點就會觸發 SQL：因為 SELECT 引用了三個不存在的列，會報 `No such field` 之類的錯誤。

#### 建議修復方案

兩處都要改：
  (a) 把 `ChkIDs.Value` 替換成常量 `False`（如果這個可選行為可以去掉），或者在 LookAtStatus 的設計視圖裡真的加一個 ChkIDs 控制元件。
  (b) 把 SELECT 改成 `ZZ_SCRATCH_STATUS.c_personid` 和 `ZZ_SCRATCH_STATUS.c_status_code`，並去掉對 `c_status_count` 的聚合，或用別的方式計算（源表裡就沒有 c_status_count 列）。

建議整段 sub 通盤重寫而不是單點修補——它顯然是從另一個表單整段複製過來的，列名沒校對過。

### Issue #14 — KIN_DATA 子表單的 CmdPickKinRel 呼叫不存在的 picker（frmPickKINSHIP_CODES）——但目前該子表單在主表中無入口（LATENT）

**涉及位置:** `Form_KIN_DATA_Subform`

**嚴重等級:** P5 — Latent（若日後把 `KIN_DATA Subform` 重新嵌入使用者可達的位置，會回到 P1）

#### 問題描述

**靜態缺陷確實存在，但執行時的觸發路徑當前不可達。**`Form_KIN_DATA_Subform` 第 52 行的 Sub `CmdPickKinRel_Click` 呼叫 `DoCmd.OpenForm "frmPickKINSHIP_CODES"`，並引用 `Forms!frmPickKINSHIP_CODES!frmKINSHIP_CODES.Form!c_kincode`。這兩個表單在目前的 .mdb 都不存在——形狀與 Issue #13 相同。

**為何 LATENT。**承載該按鈕的子表單 `KIN_DATA Subform` （即擁有 `CmdPickKinRel` 按鈕者）在當前的 `control_inventory.json` 裡沒有被任何 active 表單包含。使用者實際從 CBDB_Browser_2 進入的親屬介面是 `BIOG_MAIN_2_Subform`，而它包含的是 `KIN_DATA_2 Subform`（另一個版本，15 個控制元件全是唯讀欄位，沒有 `CmdPickKinRel`按鈕）。唯一仍然嵌入該子表單的位置是 `Form__TMPCLP487951`，那是設計時的備份快照，不是可導航的表單。

因為沒有任何使用者介面能走到那個 picker 按鈕，正常使用下不會彈出錯誤。但只要日後有人把 `KIN_DATA Subform` 重新嵌進可達的位置，這條潛在錯誤路徑就會立刻浮現，所以底層的修復仍值得做。

#### 復現步驟

**建議使用的範例人物：** `c_personid=1`（安惇，An Dun）

開啟人物 1（安惇，An Dun）。KIN_DATA 子資料表會顯示 5 條親屬記錄——點任一條的「kinship code」picker 欄位即可觸發這個有缺陷的 Sub。 _由 `reports/probe_demo_persons.py` 透過 SQL probe 挑選；之所以選這位，是因為其底層記錄數確實滿足這個 bug 的觸發條件。_

1. 驗證路徑**只能靜態驗證**——當前 .mdb 中並無主表單嵌入受影響的子表單，因此無法在執行時重現點選。
2. 靜態證據 (1)：開啟 `analysis/dump/vba/Form_KIN_DATA_Subform.vb` 第 52 行——可見該 Sub 確實呼叫 `DoCmd.OpenForm "frmPickKINSHIP_CODES"`。
3. 靜態證據 (2)：開啟 `analysis/dump/control_inventory.json`，以 `"frmPickKINSHIP_CODES"` 為鍵搜尋——不存在。該 picker 表單已不存在於 .mdb 中。
4. 可達性證據：在同一份 JSON 中搜索 `"KIN_DATA Subform"` 作為 `source_object` 或子表單控制元件名——只有 `Form__TMPCLP487951`（設計時備份快照）引用它。使用者實際走到的 `BIOG_MAIN_2_Subform` 嵌入的是 `KIN_DATA_2 Subform`，那一版沒有 `CmdPickKinRel` 按鈕。

#### 建議修復方案

與 Issue #13 相同：把 picker 表單恢復，或把呼叫方改成指向替代的 picker。雖然目前執行路徑不可達，靜態缺陷仍應清理，以免日後 `KIN_DATA Subform` 被重新嵌入時又冒出來。

### Issue #11 — EVENTS_DATA_2 上 c_event_record_id 控制元件綁到不存在的欄位——但該控制元件本身是隱藏的（LATENT）

**涉及位置:** `EVENTS_DATA_2 Subform`

**嚴重等級:** P5 — Latent（若日後把該控制元件改成 Visible=True 或加寬到 240 twips 以上，就會回到 P2）

#### 問題描述

**靜態缺陷確實存在，但執行時使用者看不到。**EVENTS_DATA_2 子表單上有一個叫 `c_event_record_id` 的控制元件，ControlSource 也是 `c_event_record_id`。EVENTS_DATA 與 `View_EventsData` 都沒有這個欄位（SQL 驗證：`SELECT c_event_record_id FROM View_EventsData` 會拋 `Too few parameters. Expected 1.`）。所以如果該控制元件顯示出來，確實會空白。

**為何 LATENT。**對 runtime 表單做 COM 探測（`analysis/probe_bug_10_11_12_visibility.py`）顯示該控制元件 `Visible = False`，寬 240 twips（~4mm）、高 270 twips——這就是一個隱藏的內部控制元件，幾乎可以肯定是早期殘留的 join-key 欄位，本來就不打算給使用者看。使用者不會看到空白欄位，因為根本看不到這個控制元件。2026-05-03 從 P2 降到 P5。

#### 復現步驟

**建議使用的範例人物：** `c_personid=44872`（孫才，Sun Cai）

開啟人物 44872（孫才，Sun Cai）。EVENTS 子資料表會顯示 1 條事件，其中 1 條有對應地址。相關繫結控制元件在每一列都會顯示空白。 _由 `reports/probe_demo_persons.py` 透過 SQL probe 挑選；之所以選這位，是因為其底層記錄數確實滿足這個 bug 的觸發條件。_

1. 驗證路徑**只能靜態 + COM 探測**——沒有 UI 上的可見徵狀。
2. 靜態證據：對 user mdb 跑 `SELECT c_event_record_id FROM View_EventsData` 會拋 `Too few parameters. Expected 1.`，確認該欄位不在投影裡。
3. 可見性證據：跑 `python analysis/probe_bug_10_11_12_visibility.py`，看 `analysis/dump/bug_10_11_12_visibility.json` 中 bug #11 那筆——`control_summary.visible` 是 `False`、`width` 是 240 twips。

#### 建議修復方案

若這個隱藏控制元件用不到了，直接刪除即可；若原意是隱藏的 join-key 容器，把 ControlSource 改成真實的欄位（例如 `c_event_code`），免得帶著一個失效的繫結。無論怎麼改，使用者都看不到差別——這純粹是程式碼整潔。

### Issue #12 — POSTED_TO_OFFICE_DATA_2 上 c_appt_type_code 控制元件綁到沒投影的欄位——但該控制元件是隱藏的，且使用者實際看的任職型別欄位是正常的（LATENT）

**涉及位置:** `POSTED_TO_OFFICE_DATA_2 Subform`

**嚴重等級:** P5 — Latent（若日後把該控制元件改成 Visible=True 或加寬到 180 twips 以上，就會回到 P2）

#### 問題描述

**靜態缺陷確實存在，但執行時使用者看不到。**POSTED_TO_OFFICE_DATA_2 上隱藏的內部控制元件 `c_appt_type_code` ControlSource 是 `c_appt_type_code`，而 `View_PostingOfficeData` 沒有投影這個欄位（SQL 驗證：拋 `Too few parameters. Expected 1.`）。

**為何 LATENT。** 兩個理由：

1. COM 探測（`analysis/probe_bug_10_11_12_visibility.py`）顯示該控制元件 `Visible = False`，寬 180 twips（~3mm）、高 330 twips——典型的隱藏 join-key 控制元件。
2. 同一個子表單上**真正給使用者看的**任職型別欄位是 `TxtApptType`（綁 `c_appt_desc`）與 `TxtApptTypeChn`（綁 `c_appt_desc_chn`）。這兩個欄位都在 `View_PostingOfficeData` 的投影裡——SQL 探測能拿到真實值（例如 `'Regular Appointment'` / `'正授'`）。所以 Postings 分頁上的任職型別**是正常顯示的**；只有隱藏的 `c_appt_type_code` 控制元件壞掉。

2026-05-03 從 P2 降到 P5——原本「任職型別列每一列都是空白」的 P2 說法是錯的；使用者實際看的任職型別欄位是正常的。

#### 復現步驟

**建議使用的範例人物：** `c_personid=2`（安邡，An Fang）

開啟人物 2（安邡，An Fang）。POSTED-TO-OFFICE 子資料表會顯示 1 條官職任命記錄，c_appt_code 都不為 NULL；但任職型別那一列每一列都是空白。 _由 `reports/probe_demo_persons.py` 透過 SQL probe 挑選；之所以選這位，是因為其底層記錄數確實滿足這個 bug 的觸發條件。_

1. 驗證路徑**只能靜態 + COM 探測**——沒有 UI 上的可見徵狀。
2. 靜態證據：`SELECT c_appt_type_code FROM View_PostingOfficeData` 會拋 `Too few parameters. Expected 1.`，確認該欄位不在投影裡。
3. 可見性證據：跑 `python analysis/probe_bug_10_11_12_visibility.py`，看 `analysis/dump/bug_10_11_12_visibility.json` 中 bug #12 那筆——`control_summary.visible` 是 `False`、width 是 180 twips。
4. 反證（使用者實際看的欄位是正常的）：`SELECT TOP 1 c_appt_desc, c_appt_desc_chn FROM View_PostingOfficeData` 能返回真實值（例如 `'Regular Appointment'` / `'正授'`），這就是 `TxtApptType` / `TxtApptTypeChn`（可見的兩個控制元件）渲染出來的內容。

#### 建議修復方案

若這個隱藏控制元件用不到了，刪除即可；若是有意為之的隱藏 join-key 容器，把 ControlSource 改成真實的欄位（例如 `c_appt_code`）。無論怎麼改，使用者都看不到差別——這純粹是程式碼整潔。

## 附錄 A —— c_index_year / c_index_addr_id 與 cbdb-online-main-server 快照之間的偏差（差異需要逐筆分類後才能判定是否為缺陷）

我們把本 .mdb 的 BIOG_MAIN 與 cbdb-online-main-server 每週釋出的 SQLite 快照在 `c_index_year`、`c_index_addr_id` 兩個欄位上做比對，可以看到一小部分人物對不齊。

**兩邊是兩套獨立的實作。**SQLite 快照中的 `c_index_year` 是 cbdb-online-main-server 的 PHP `IndexYearRebuildService.php` 算出來的，`c_index_addr_id` 則是 `IndexAddressRebuildService.php` 算出來的（程式碼都在 <https://github.com/cbdb-project/cbdb-online-main-server>）；User MDB 上對應的這兩個User MDB 那一邊：`c_index_addr_id` 由前端 mdb 裡的 `Form_frmIndexAddr` VBA 重建；`c_index_year` 由連結表後端 `data/CBDB_<YYYYMMDD>_DATA.mdb` 裡 **37 條 `BM IY Rule …` 的 QueryDef** 重建，由 `frmBaseMaintenance` 驅動。兩邊演算法已抽取到 `analysis/dump_data/querydefs_index/*.sql`；form / module 驅動 VBA 仍需 Access SaveAsText 互動式提取。PHP **意圖**映象 VBA，但兩者是兩條獨立的程式路徑。每一行差異**可能**來自下列至少四個原因，光看差異本身分不出來：(1) 源資料快照漂移；(2) PHP 與 VBA 之間的演算法 / 移植差異；(3) 優先序 / 平手規則不同；(4) null / 預設值處理不同。

**我們並沒有對目前看到的 ~575 / 657 246 筆差異做完整分類。**下方列舉的樣本（目前共 13 筆、3 種分桶，來自 `reports/index_drift_examples.json`）只是**示範**這些差異**長什麼樣**，並非統計上有代表性，是後續逐筆分類的起點，不是結論。

### 分類匯總

比對了兩邊都有的 **657,245** 個 personid（User MDB 共 657,784 筆；SQLite 共 657,478 筆；僅 User MDB 有 539 筆；僅 SQLite 有 233 筆）。

| 分桶 | 筆數 | 佔比 | 含義 |
|---|---:|---:|---|
| `exact_match` | 656,682 | 99.914% | 四個欄位全部一致 |
| `source_drift_index_agrees` | 2 | 0.000% | 源資料有漂移但兩邊 index 都一致 |
| `source_drift_index_diffs_too` | 14 | 0.002% | 源資料有漂移、且至少一個 index 不同 |
| `index_year_only_diff` | 59 | 0.009% | 生年/卒年一致，但只有 c_index_year 不同 —— 待追查 |
| `index_addr_only_diff` | 478 | 0.073% | 生年/卒年一致，但只有 c_index_addr_id 不同 —— 待追查 |
| `index_both_diff` | 10 | 0.002% | 生年/卒年一致，但兩個 index 都不同 —— 複合差異最強訊號 |

淨差異：**563** / 657,245（0.086 %）。其中 **16** 筆能明確歸因於 birthyear / deathyear 的源資料漂移；剩下 **547** 筆需要逐筆追查（可能是 PHP↔VBA 演算法差異，也可能是本分類器沒有比較的 evidence 表（BIOG_ADDR_DATA / ENTRY_DATA / NIAN_HAO 等）裡的漂移）。完整輸出見 `reports/index_drift_classification.json`，演算法來源指標見 `analysis/index_drift_algorithm_notes.md`。

### 年份差異 —— 逐筆 rule 分類

在 **69** 筆「只有 c_index_year 不一致」的行中，逐筆比對 PR N (`analysis/index_year_rule_comparison.md`) 的runtime-vs-PHP 規則級差異。保守分類如下：

| 分桶 | 筆數 |
|---|---:|
| `php_returned_sentinel` (PHP 寫了 sentinel／溢位值) | 1 |
| `php_did_not_compute` (PHP 沒算出值（覆蓋率缺口）) | 19 |
| `access_did_not_compute` (Access 沒算出值（覆蓋率缺口）) | 7 |
| `iteration_order_diff` (Phase-C 迭代次數不同) | 5 |
| `consistent_within_rule` (多列共享同一 (php_tcode, access_tcode, diff)) | 14 |
| `candidate_algorithm_divergence` (形狀符合 K1 的歷史 hypothesis probe 但無法以單筆證據重建) | 5 |
| `unclassified` (尚未對上任何模式) | 18 |

以上沒有任何一筆被視為已確認的 bug。逐筆輸出見 `reports/index_year_drift_rule_classification.json`。

PR K2 進一步的 triage (`analysis/triage_index_year_drift_groups.py` → `reports/index_year_drift_rule_groups.json`) 把剩下的桶命名清楚：

- `consistent_within_rule` × 14 → 5 個 signature 分組。PR AI + AJ 的逐筆探測推翻了原本的 tie-break 假說：14 筆全是 `source_data_drift_biog_main_or_kin_data_between_sides`（8 筆 BIOG_MAIN birthyear 漂移 + 6 筆 KIN_DATA evidence-pid 漂移）。屬於 PHP-side / SQLite snapshot 的上游資料漂移，並非 CBDB 演算法差異。
- `unclassified` × 18 → 18 筆已命名，17 筆標為 `blocked_by_runtime_priority_triage_pending`（PR M 已 dump frmBaseMaintenance，原始碼已在 repo；要逐筆判斷哪邊正確仍需走一遍 runtime 的 priority／iteration 順序）。
- `php_did_not_compute` × 19 → 按 Access tcode 分 6 組；最大的是 `access_tcode='05'` × 7（jinshi 進士類的 `candidate_php_entry_code_mapping_gap`）。

### c_index_addr_id 差異 —— 逐筆分類

在 **488** 筆 c_index_addr 差異中（PR G 的 478 `index_addr_only_diff` + 10 `index_both_diff`），逐筆把兩邊的 BIOG_ADDR_DATA 代入「rank-priority + MAX(c_sequence)」演算法重算，與實際儲存值對照分類：

| 分桶 | 筆數 |
|---|---:|
| `mdb_stale_index_addr` | 412 |
| `mdb_value_php_null` | 47 |
| `same_candidates_diff_winner` | 10 |
| `both_stale_recompute_mismatch` | 10 |
| `both_sides_match_recomputed` | 6 |
| `sqlite_stale_index_addr` | 2 |
| `mdb_null_php_value` | 1 |

以上沒有任何一筆被視為已確認的 bug。412 筆 `mdb_stale_index_addr` 屬於維護週期差異（User MDB 在下次釋出前需要重跑 frmBaseMaintenance）。10 筆 `same_candidates_diff_winner` 是唯一的候選演算法差異。逐筆輸出見 `reports/index_addr_drift_classification.json`。

PR M（`analysis/dump_data_mdb_vba.py`）從 DATA mdb 抽出了 `frmBaseMaintenance.CmdIndexAddress_Click`。它**沒有**像 PHP 那樣明確 `MAX(c_sequence)` 聚合 —— 在維護週期差異之外，這還是一個候選演算法差異。建議的 release checklist 緩解步驟：在 User MDB 出貨前先在 DATA mdb 上跑 `CmdIndexYear`，再跑 `CmdIndexAddress`。詳見 `analysis/index_drift_algorithm_notes.md` 中的 "Maintenance trigger path" 段。

### 目前能解釋的 drift 原因

每個 bucket 的成因／證據／信心度／下一步追查都寫在 `analysis/index_drift_cause_analysis.md`。本節只列每個 bucket 的計數和信心度摘要；目前沒有任何 bucket 被列為已確認的 CBDB bug。

**c_index_year 原因桶**

| Bucket | 筆數 | 信心度 |
|---|---:|---|
| `php_returned_sentinel` | 1 | high |
| `php_did_not_compute` | 19 | tcode='05' × 7: supported_by_focused_probe (PR Z).  tcode='11' × 5: medium.  Phase-C tcodes (14/20/2304): medium.  tcode='07' × 1: medium (vestigial-vs-intentional unresolved). |
| `access_did_not_compute` | 7 | medium |
| `iteration_order_diff` | 5 | medium |
| `consistent_within_rule` | 14 | supported_by_focused_probe (PR AI + AJ) |
| `candidate_algorithm_divergence` | 5 | low-medium |
| `blocked_by_runtime_priority_triage_pending` | 17 | low (per-row causes); high (category label) |

**c_index_addr_id 原因桶**

| Bucket | 筆數 | 信心度 |
|---|---:|---|
| `mdb_stale_index_addr` | 412 | high |
| `mdb_value_php_null` | 47 | medium |
| `same_candidates_diff_winner` | 10 | high |
| `both_stale_recompute_mismatch` | 10 | medium-high |
| `both_sides_match_recomputed` | 6 | low-medium |
| `sqlite_stale_index_addr` | 2 | medium |
| `mdb_null_php_value` | 1 | medium-high |

建議優先處理的調查專案（完整列表見 cause-analysis md）：

1. B1 release-process step (CmdIndexYear → CmdIndexAddress before shipping User MDB) —— 可消化 412 筆；工程成本：zero (process change)。
2. B3 secondary tie-break (MIN(c_addr_id)) added to both implementations —— 可消化 10 筆；工程成本：small algorithm tweak per side。
3. A2 tcode 05 entry-code-mapping check — DONE by PR Z (6 mapping gaps + 1 c_year=0 gap; all PHP-side upstream data) —— 可消化 7 筆；工程成本：single SQL probe (already run)。

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

本節將 `CBDB_20260430_DATA.mdb` 中 `TablesFields` 表的內容與 `reports/collect_schema_diffs.py` 透過 Access DAO（TableDefs）重建的資料庫結構進行比對。若存在差異，表示文件表可能已過時。

TablesFields 共 875 筆。從資料庫重建：996 筆。

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
