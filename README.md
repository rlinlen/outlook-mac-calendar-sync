# Mac Outlook Calendar Reader

這個專案包含了用於讀取Mac Outlook行事曆事件的Python腳本，可以將行事曆資料匯出為CSV格式，並支援同步到Google Calendar。

## ！限制(Limitation)
- 不支援recurring event
- 本版本為beta版，未經過廣泛測試，請謹慎使用

## 🆕 最新更新 (2025-06-30)

### 重大修復和改進
- ✅ **修復中文字符亂碼問題**：正確處理UTF-16編碼，支援中文主題如「【Online】銷售預測第三階段成果分享」
- ✅ **Location欄位智能提取**：支援Google Calendar地點、Amazon Chime會議、Microsoft Teams會議和中文地址
- ✅ **強制更新功能**：新增 `--force` 參數，可強制更新所有Google Calendar事件
- ✅ **快取管理**：新增 `--clear-cache` 參數，支援清除同步快取重新開始
- ✅ **增量同步優化**：智能檢測事件變更，只同步需要更新的事件
- ✅ **刪除檢測功能**：自動檢測已刪除的Outlook事件，在Google Calendar中標記為 `[DELETED]`
- ✅ **預設天數調整**：從7天改為14天，提供更好的覆蓋範圍
- ✅ **自動刷新機制**：智能的Google OAuth憑證自動刷新，Access Token剩餘時間<10分鐘時自動提前刷新
- ✅ **定期執行排程**：新增自動排程器，每週自動執行同步保持憑證活躍
- ✅ **時區問題修復**：修復 "can't subtract offset-naive and offset-aware datetimes" 錯誤

### 協議級解析成就
- 🔧 **二進制協議逆向工程**：完全解析.olk15Event文件結構
- 🔧 **多語言支援**：完整支援中英文混合內容和特殊字符
- 🔧 **智能刪除檢測**：基於快取比較自動檢測已刪除事件


## 系統需求

- macOS系統
- Python 3.x
- Microsoft Outlook for Mac已安裝並有行事曆資料
- Google Calendar API憑證（用於同步功能）

## 使用方法
請先git clone，然後在專案目錄下(./outlook-mac-calendar-sync)執行以下操作

### 1. 安裝Dependency
```bash
uv venv
source .venv/bin/activate
uv pip install -r requirements.txt
```


### 2. Google Calendar 同步

#### 設定 Google Calendar API

請參考 `SETUP_GOOGLE_CALENDAR_API.md` 完成以下步驟：
1. 建立 Google Cloud 專案
2. 啟用 Google Calendar API
3. 建立 OAuth 2.0 憑證
4. 下載憑證檔案為 `client_secret.json`
5. 在專案目錄下，創建data資料夾，並將`client_secret.json`放到data中

未來如果token過期，會跳出Google OAuth視窗，接受即可。

#### 執行同步

**方法一：使用一鍵腳本（推薦）**
在專案目錄下，執行以下腳本：
```bash
./script/sync_outlook_to_google.sh
```

**方法二：分步執行**
```bash
# 步驟 1: 生成 Outlook CSV
uv run script/dump_outlook_calendar.py

# 指定其他時區
uv run script/dump_outlook_calendar.py --timezone UTC+0    # UTC時間
uv run script/dump_outlook_calendar.py --timezone UTC-5   # 美國東岸時間
uv run script/dump_outlook_calendar.py --timezone UTC+9   # 日本時間

# 步驟 2: 同步到 Google Calendar
uv run script/sync_csv_with_google_calendar_improved.py
```

### 3. [Optional] 設定排程
使用cron自動進行同步
3-1:
```bash
crontab -e
```

3-2:
```bash
*/51 * * * * cd <your-folder>/outlook-mac-calendar-sync && ./script/sync_outlook_to_google.sh >> data/cron.log 2>&1
```

3-3:
Grant Full Disk Access to cron：
1. Open System Preferences
   • Click the Apple menu → System Preferences
   • Or use Spotlight: Press Cmd + Space, type "System Preferences"

2. Navigate to Privacy Settings
   • Click on "Security & Privacy"
   • Select the "Privacy" tab at the top

3. Access Full Disk Access Settings
   • In the left sidebar, scroll down and click "Full Disk Access"
   • You'll see a list of applications that have full disk access

4. Unlock Settings
   • Click the lock icon (🔒) in the bottom left corner
   • Enter your administrator password when prompted
   • The lock should now show as unlocked (🔓)

5. Add uv to the List
   • If the uv is already in the list, enable it.
   • Otherwise Click the "+" (plus) button below the application list
      • In the file browser that opens:
         • Press Cmd + Shift + G to open "Go to Folder"
         • Find the uv binary
         • Select the cron file and click "Open"


7. Lock Settings Again
   • Click the lock icon to prevent further changes
   • Close System Preferences


如未設定，會跳出““uv” would like to access data from other apps.”

## 檔案說明

### 主要程式檔案

1. **`dump_outlook_calendar.py`** - 完整版行事曆讀取器（推薦）
   - 包含所有資料庫欄位：Calendar_UID、Record_ModDate
   - 正確處理UTC時間和使用者時區轉換
   - 智能主題提取，支援中文和特殊字符
   - 修正CSV格式問題
   - **新增**：完整的UTF-16中文字符解碼支援
   - **新增**：Google Calendar HTML內容解析
   - **新增**：Amazon Chime/Teams會議自動識別
   - 輸出檔案：`dump_outlook_calendar.csv`

2. **`sync_csv_with_google_calendar.py`** - Google Calendar同步器
   - 智能事件去重複和更新檢測
   - 本地快取機制避免重複同步
   - 支援事件創建和更新
   - **新增**：強制更新模式 (`--force`)
   - **新增**：快取清除功能 (`--clear-cache`)
   - 完整錯誤處理和進度顯示

### 配置檔案

- **`SETUP_GOOGLE_CALENDAR_API.md`** - Google API設定指南
- **`sync_outlook_to_google.sh`** - 一鍵同步腳本

## 功能特色

### 核心功能
- **時間範圍**：讀取今天起算接下來七天的行事曆事件
- **時間格式**：正確處理Outlook使用的「從1601-01-01 UTC開始的分鐘數」時間格式
- **時區處理**：正確處理UTC時間，支援使用者自訂時區顯示
- **完整資料庫欄位**：包含Calendar_UID和Record_ModDate等重要欄位

### 資料提取
- Calendar_UID（行事曆唯一識別碼）
- Record_ModDate（記錄修改日期）
- Subject（主題）
- Location（地點）
- Organizer（組織者）
- Duration（持續時間）
- Starts（開始時間）
- Ends（結束時間）
- Body（內容）

### Google Calendar同步功能 🆕
- **智能去重複**：使用Calendar_UID和Record_ModDate避免重複同步
- **增量同步**：只同步變更的事件，提高效率
- **本地快取**：記錄同步狀態，支援中斷恢復
- **強制更新模式**：忽略快取，強制更新所有事件
- **快取管理**：支援清除快取重新同步

#### 進階同步選項

```bash
# 正常增量同步（只同步變更的事件）
uv run sync_csv_with_google_calendar_improved.py

# 強制更新所有事件（忽略快取檢查）
uv run sync_csv_with_google_calendar_improved.py --force

# 清除快取並重新同步
uv run sync_csv_with_google_calendar_improved.py --clear-cache

# 查看幫助信息
uv run sync_csv_with_google_calendar_improved.py --help
```

## 輸出檔案

程式執行後會在當前目錄生成CSV檔案，包含以下欄位：

### 完整版輸出欄位（推薦）

| 欄位 | 說明 |
|------|------|
| Calendar_UID | 行事曆事件唯一識別碼 |
| Record_ModDate | 記錄修改日期（Unix時間戳） |
| Subject | 會議主題 |
| Location | 會議地點 |
| Organizer | 會議組織者電子郵件 |
| Duration | 會議持續時間（小時） |
| Starts | 開始時間（使用者時區） |
| Ends | 結束時間（使用者時區） |
| Starts_UTC | 開始時間（UTC） |
| Ends_UTC | 結束時間（UTC） |
| Body | 會議內容/描述 |

### Calendar_UID 格式類型

- **Amazon Meetings**: `Meetings-1750998579963-ty-6b8b69b3a5657505fc82c1edd4d77b92`
- **Exchange格式**: `040000008200E00074C5B7101A82E00800000000005C9B6C81E7DB010000000000000000100000005265EC731D93EA4695D069C844E4F519`
- **GUID格式**: `A14A9B4A-1742-4971-AADB-09EDB65F0B52`
- **Google Calendar**: `35denat4othqdc4omp2jbhvmk8@google.com`

## 技術細節

### 資料來源

程式從以下位置讀取Outlook資料：
- **SQLite資料庫**：`~/Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles/Main Profile/Data/Outlook.sqlite`
- **事件檔案**：`~/Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles/Main Profile/Data/Events/`

### 時間格式轉換

Outlook使用特殊的時間格式：從1601-01-01 UTC開始計算的分鐘數。程式會自動轉換為標準的datetime格式，並支援多種時區顯示：

- **UTC時間**：資料庫中儲存的原始UTC時間
- **使用者時區**：根據指定時區轉換後的本地時間
- **時區支援**：UTC+8（台北）、UTC-5（美東）、UTC+0（倫敦）等

### 解析邏輯

1. 從SQLite資料庫查詢指定時間範圍的事件
2. 根據PathToDataFile找到對應的.olk15Event檔案
3. **基於二進制協議解析**：
   - **長度字段讀取**：從固定位置0x180和0x188讀取Subject和Location長度
   - **UTF-16邊界檢測**：基於長度字段實現精確的字段分離
   - **協議級提取**：直接從`</html>`標籤後提取UTF-16編碼的Subject和Location
   - **通用解析方法**：適用於所有事件類型，不依賴規則式匹配

### 二進制協議發現 🎯

#### .olk15Event文件結構
```
文件頭部:
0x180: Subject長度字段 (4字節小端序)
0x184: 標記字段 (4字節)
0x188: Location長度字段 (4字節小端序)
0x18c: 標記字段 (4字節)

數據區域:
</html>\r\n + Subject(UTF-16 LE) + Location(UTF-16 LE)
```

#### 協議解析流程
1. **搜索標記字節**：在文件頭部搜索 `02 00 00 1f` 標記字節
2. **讀取長度字段**：從標記位置+4和+12字節讀取Subject和Location長度
3. **定位HTML結束標記**：搜索`</html>`的UTF-16編碼
4. **精確提取**：基於長度字段精確提取Subject和Location
5. **UTF-16解碼**：處理中英文混合內容和特殊字符

### 智能解析改進 🆕

#### UTF-16字符解碼
- **改進的字節對齊處理**：正確處理UTF-16 Little Endian編碼
- **中文字符保護**：保留Unicode範圍 0x4e00-0x9fff 的中文字符
- **替換字符清理**：移除UTF-16解碼失敗的替換字符（\ufffd）

#### .olk15Event二進制協議解析 🔥
基於對.olk15Event文件二進制結構的深度分析，實現了協議級別的Subject和Location提取：

**Subject存儲協議**：
1. **定位標記**：`</html>` 標籤結束 (UTF-16 LE: `3c 00 2f 00 68 00 74 00 6d 00 6c 00 3e 00`)
2. **分隔符**：`0d 00` (回車符)
3. **Subject內容**：緊接著的UTF-16 LE編碼文本
4. **字節序列示例**：
   ```
   32f0: 3e 00 0d 00 3c 00 2f 00  68 00 74 00 6d 00 6c 00  |>...<./.h.t.m.l.|
   3300: 3e 00 0d 00 41 00 57 00  53 00 20 00 53 00 65 00  |>...A.W.S. .S.e.|
         ^^^^^^^^^ ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
         </html>\r  A  W  S     S  e
   ```

**Location存儲協議**：
1. **位置**：Subject文本結束後直接相鄰（無分隔符）
2. **編碼**：UTF-16 LE編碼
3. **結束標記**：控制字節序列（如 `03 00 00 00`）

**長度字段協議** 🆕：
1. **標記字節發現**：搜索標記字節 `02 00 00 1f` 來定位長度字段
2. **Subject長度字段**：標記字節位置 + 4字節 (4字節小端序)
3. **Location長度字段**：標記字節位置 + 12字節 (4字節小端序)
4. **協議結構**：
   ```
   0x????: xx xx 00 00  (某個值)
   0x????+4: 02 00 00 1f  <-- 標記字節
   0x????+8: xx 00 00 00  <-- Subject長度字段
   0x????+12: 04 00 00 1f  <-- 第二個標記字節
   0x????+16: xx 00 00 00  <-- Location長度字段
   ```
5. **通用性**：此標記字節模式在所有.olk15Event文件中一致

#### 協議解析改進 🎯
- **完全移除規則式代碼**：不再依賴硬編碼的關鍵詞或模式匹配
- **精確邊界檢測**：基於長度字段實現精確的Subject和Location分離
- **通用協議方法**：適用於所有.olk15Event文件，不限於特定內容
- **智能邊界檢測**：基於Unicode字符範圍和控制字節模式
- **中文字符支援**：正確識別和保留中文Unicode範圍（0x4e00-0x9fff）
- **混合數據處理**：能夠從包含控制字節的數據中提取純文本內容

## 注意事項

1. **權限**：程式需要讀取Outlook資料目錄的權限
2. **Outlook版本**：專為Outlook for Mac 15.x版本設計
3. **時區**：所有時間都會轉換為台北時間（UTC+8）顯示
4. **編碼**：支援中文和其他Unicode字符

## 故障排除

### Outlook Calendar 讀取問題

1. **找不到資料庫檔案**
   - 確認Outlook已安裝且有資料
   - 檢查檔案路徑是否正確

2. **沒有找到事件**
   - 確認指定時間範圍內有行事曆事件
   - 檢查事件是否已同步到本地

3. **解析錯誤**
   - 某些特殊格式的事件可能無法完全解析
   - 程式會跳過有問題的檔案並繼續處理其他事件

4. **中文字符亂碼**
   - 已修復：使用最新版本的 `outlook_calendar_complete_fixed.py`
   - 支援完整的UTF-16中文字符解碼

### Google Calendar 同步問題

1. **認證失敗**
   - 確認 `client_secret.json` 檔案存在且正確
   - 重新下載 Google API 憑證檔案
   - 刪除 `token.json` 重新認證

2. **API 配額超限**
   - Google Calendar API 有每日請求限制
   - 等待24小時後重試
   - 考慮申請更高的配額

3. **事件重複**
   - 使用 `--clear-cache` 清除快取
   - 檢查 Google Calendar 中是否有重複事件
   - 手動刪除重複事件後重新同步

4. **同步失敗**
   - 檢查網路連線
   - 確認 Google Calendar API 已啟用
   - 查看錯誤訊息進行診斷

5. **強制更新不生效**
   - 使用 `--force` 參數：`uv run sync_csv_with_google_calendar_improved.py --force`
   - 確認 CSV 檔案是最新的
   - 檢查快取檔案是否正確更新

6. **刪除檢測問題** 🆕
   - 檢查快取檔案 `sync_cache.json` 是否存在
   - 使用 `--clear-cache` 清除快取重新建立
   - 確認已刪除的事件確實不在當前CSV中
   - 檢查Google Calendar中是否有 `[DELETED]` 標記的事件

## Google Calendar 同步功能

### 新增檔案

4. **`sync_csv_with_google_calendar_improved.py`** - Google Calendar 同步器
   - 智能事件去重複和更新檢測
   - 本地快取機制避免重複同步
   - 支援事件創建和更新
   - **強制更新模式**：`--force` 忽略快取，更新所有事件
   - **快取管理**：`--clear-cache` 清除快取重新同步
   - 完整錯誤處理和進度顯示
   - 輸出：同步事件到 Google Calendar

5. **`sync_outlook_to_google.sh`** - 一鍵同步腳本
   - 自動執行 Outlook 讀取和 Google 同步
   - 完整的錯誤檢查和狀態顯示

6. **`SETUP_GOOGLE_CALENDAR_API.md`** - Google API 設定指南
   - 詳細的 Google Cloud Console 設定步驟
   - OAuth 2.0 憑證配置說明

### Google Calendar 同步特色

- **智能去重複**：使用 Calendar_UID 和 Record_ModDate 避免重複同步
- **增量同步**：只同步變更的事件，提高效率
- **本地快取**：記錄同步狀態，支援中斷恢復
- **事件識別**：在 Google Calendar 事件描述中嵌入 Outlook UID 以便識別
- **完整資料同步**：
  - 事件標題、時間、地點
  - 組織者資訊
  - 事件內容/描述
  - UTC 時間正確轉換
- **錯誤處理**：詳細的錯誤報告和重試機制

### Google Calendar 同步使用方法

#### 1. 設定 Google Calendar API

請參考 `SETUP_GOOGLE_CALENDAR_API.md` 完成以下步驟：
1. 建立 Google Cloud 專案
2. 啟用 Google Calendar API
3. 建立 OAuth 2.0 憑證
4. 下載憑證檔案為 `client_secret.json`

#### 2. 安裝依賴套件

```bash
# 使用 uv（推薦）
uv sync

# 或使用 pip
pip3 install -r requirements.txt
```

#### 3. 執行同步

**方法一：使用一鍵腳本（推薦）**
```bash
./sync_outlook_to_google.sh
```

**方法二：分步執行**
```bash
# 步驟 1: 生成 Outlook CSV
uv run outlook_calendar_complete_fixed.py

# 步驟 2: 同步到 Google Calendar
uv run sync_csv_with_google_calendar_improved.py
```

#### 4. 同步選項

```bash
# 正常增量同步（只同步變更的事件）
uv run sync_csv_with_google_calendar_improved.py

# 強制更新所有事件（忽略快取檢查）
uv run sync_csv_with_google_calendar_improved.py --force

# 清除快取並重新同步
uv run sync_csv_with_google_calendar_improved.py --clear-cache

# 停用刪除檢測（預設啟用）
uv run sync_csv_with_google_calendar_improved.py --no-mark-deleted

# 查看幫助信息
uv run sync_csv_with_google_calendar_improved.py --help
```

#### 5. 刪除檢測功能 🆕

同步器會自動檢測已從Outlook中刪除的事件：

- **自動檢測**：比較當前事件與快取，找出已刪除的事件
- **安全標記**：在Google Calendar中將已刪除事件標題改為 `[DELETED] 原標題`
- **保留記錄**：在事件描述中添加刪除時間戳記
- **智能清理**：如果Google Calendar事件已不存在，自動從快取中移除

```bash
# 啟用刪除檢測（預設）
uv run sync_csv_with_google_calendar_improved.py

# 停用刪除檢測
uv run sync_csv_with_google_calendar_improved.py --no-mark-deleted
```

#### 5. 首次執行認證

第一次執行時會：
1. 開啟瀏覽器進行 OAuth 認證
2. 要求授權存取 Google Calendar
3. 生成 `token.json` 檔案供後續使用

### 同步檔案說明

| 檔案 | 說明 |
|------|------|
| `client_secret.json` | Google API OAuth 憑證（需要下載） |
| `token.json` | OAuth 存取令牌（自動生成） |
| `sync_cache.json` | 同步快取檔案（自動生成） |
| `requirements.txt` | Python 依賴套件清單 |
| `pyproject.toml` | 現代 Python 專案配置檔案 |

### 同步結果

同步完成後會顯示：
- ✅ 成功同步的事件數量
- ❌ 失敗的事件數量
- 📊 處理進度和詳細狀態

事件會出現在你的 Google Calendar 主行事曆中，並在描述中包含原始的 Outlook Calendar UID 以便識別和管理。

## 注意事項

#### 1. 設定 Google Calendar API

請參考 `SETUP_GOOGLE_CALENDAR_API.md` 完成以下步驟：
1. 建立 Google Cloud 專案
2. 啟用 Google Calendar API
3. 建立 OAuth 2.0 憑證
4. 下載憑證檔案為 `client_secret.json`

#### 2. 安裝依賴套件

```bash
# 使用 uv（推薦）
uv sync

# 或使用 pip
pip3 install -r requirements.txt
```

#### 3. 執行同步

**方法一：使用一鍵腳本（推薦）**
```bash
./sync_outlook_to_google.sh
```

**方法二：分步執行**
```bash
# 步驟 1: 生成 Outlook CSV
uv run outlook_calendar_complete_fixed.py

# 步驟 2: 同步到 Google Calendar
uv run sync_csv_with_google_calendar_improved.py
```

#### 4. 首次執行認證

第一次執行時會：
1. 開啟瀏覽器進行 OAuth 認證
2. 要求授權存取 Google Calendar
3. 生成 `token.json` 檔案供後續使用

### 同步檔案說明

| 檔案 | 說明 |
|------|------|
| `client_secret.json` | Google API OAuth 憑證（需要下載） |
| `token.json` | OAuth 存取令牌（自動生成） |
| `sync_cache.json` | 同步快取檔案（自動生成） |
| `requirements.txt` | Python 依賴套件清單 |
| `pyproject.toml` | 現代 Python 專案配置檔案 |

### 同步結果

同步完成後會顯示：
- ✅ 成功同步的事件數量
- ❌ 失敗的事件數量
- 📊 處理進度和詳細狀態

事件會出現在你的 Google Calendar 主行事曆中，並在描述中包含原始的 Outlook Calendar UID 以便識別和管理。
