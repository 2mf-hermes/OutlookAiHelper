# Outlook AI 助手 — D3 Design

狀態：待使用者確認後才進入 D4 / M1 實作  
已確認決策：COM 為主 + PST 為輔；本機規則優先、AI 選擇性；使用者零安裝  
技術棧：C# / .NET Framework 4.8 + WPF + Outlook COM 後期繫結 + 本地 JSON

---

## 0. 設計順序

1. Information Architecture  
2. Interaction  
3. Components  
4. Visual Layer  

系統模組與資料模型與 UI 平行定義，但 UI 不先於 IA。

---

## 1. Information Architecture

### 1.1 主導覽（單一主視窗）

| 區域 | 職責 | 說明 |
|---|---|---|
| 側欄（左） | 一級入口 | 四象限、待回待辦、設定 |
| 工具列（上） | 範圍與動作 | 資料夾、天數、掃描、重新整理 |
| 列表（中） | 郵件或待辦清單 | 可虛擬化長清單 |
| 詳情（右） | 單筆細節 | 命中原因、手動改象限、待辦備註 |

小視窗（寬 < 1100）時：詳情改為滑出面板或全頁返回，側欄收成圖示列。

### 1.2 Screen Map

```text
MainWindow
├── QuadrantBoard（預設）
│   ├── Segment: 全部 | 緊急重要 | 緊急不重要 | 重要不緊急 | 不重要不緊急
│   ├── MailList
│   └── MailDetail
│       ├── 元資料（主旨、寄件者、時間、資料夾）
│       ├── ScoreReasonList（可解釋）
│       └── Actions（改象限、加入待辦、開啟於 Outlook）
├── TodoBoard
│   ├── Filter: 待處理 | 已完成 | 全部
│   ├── TodoList
│   └── TodoDetail / Inline edit
│       └── RelatedMails（同主旨串的所有信）
└── Settings
    ├── 來源（本機 Outlook／雲端 Microsoft 365）
    ├── 掃描範圍（資料夾、天數、帳號）
    ├── 分類規則（VIP、關鍵字、權重說明）
    ├── AI（預設關、Provider、白名單欄位）
    ├── 語言（繁中 / 簡中 / English）
    ├── 隱私（資料位置、匯出、清除）
    └── 關於（版本、唯讀聲明）
```

### 1.3 導覽原則

- 三入口固定位置，不隨內容重排。  
- 四象限是預設落地頁，先看到「現在該處理什麼」。  
- 設定不進主工作流，但隱私權限在首次掃描前可到達。  
- 不做多層返回；詳情永遠可關閉回列表。

---

## 2. Interaction

### 2.1 核心流程

**掃描**

```text
使用者選範圍 → 點「開始掃描」
→ 進度（已讀 n / 上限 m，可取消）
→ 完成：列表依象限填入
→ 失敗：錯誤卡 + 重試 / 開啟設定
```

**更正分類**

```text
選一封信 → 詳情看原因
→ 「改到其他象限」
→ 寫入本地 Override（不寫 Outlook）
→ 列表立即更新；原因區顯示「手動指定」
```

**待回待辦**

```text
從郵件「加入待辦」或在 TodoBoard 手動新增
→ 標題預填主旨，可改
→ 勾完成 / 加備註 / 刪除
→ 點一列 → 右側面板顯示該待辦與「相關的郵件」（同一主旨串的所有信）
→ 相關的郵件依共用排序（§12）
→ 關閉程式後保留（JSON）
```

**AI（選擇性，預設關）**

```text
設定開啟 AI → 使用增強功能前
→ 顯示將送出的欄位白名單預覽
→ 使用者確認 → 呼叫 Provider
→ 失敗：僅關閉增強，核心分類不受影響
```

### 2.2 錯誤流

| 情境 | 使用者訊息方向 | 復原 |
|---|---|---|
| Outlook 未安裝/未啟動 | 找不到可用的 Outlook | 啟動 Outlook 後重試 |
| 無郵件帳號 | 尚未設定郵件帳戶 | 指引到 Outlook |
| COM 失敗/逾時 | 無法讀取郵件，已中止本次掃描 | 重試、縮小範圍 |
| 範圍內無信件 | 這個範圍沒有郵件 | 改天數或資料夾 |
| JSON 損毀 | 設定檔無法讀取 | 備份壞檔、重設單一檔 |
| 磁碟無法寫入 | 無法儲存待辦 | 檢查權限/空間 |
| AI 失敗 | 建議功能暫時不可用 | 核心功能仍在 |

禁止空白主畫面。任何失敗都要有：人話訊息、技術 Log、下一步動作。

### 2.3 狀態矩陣（每個資料視圖）

| 狀態 | 要求 |
|---|---|
| Empty | 說明原因 + 主行動（例如「開始掃描」） |
| Loading | 進度、可取消（長任務） |
| Error | 原因 + 重試/設定 |
| Default / Hover / Focus / Active / Selected / Disabled | 互動元件必備 |

---

## 3. Components

不先做漂亮畫面再塞功能。元件清單服務於上面的 IA/互動。

| 元件 | 用途 | 必備狀態 |
|---|---|---|
| `NavRail` | 三入口；設定齒輪在**右下角**只佔一個圖示寬（38×38），不隨 panel 撐滿 | selected / hover / focus |
| `GlassPanel` | 玻璃容器（詳情、設定卡） | default |
| `Toolbar` | 範圍 + 掃描 | disabled（掃描中） |
| `QuadrantSegment` | 象限切換 | selected 等 |
| `MailList` / `MailRow` | 郵件清單；清單上方一行摘要（共 N 封 · M 個主旨串 · K 封已加入待辦）。同一主旨串（去掉 Re:/回覆：等前綴後同主旨）合併成一列，列尾有展開箭頭，展開後每個成員各佔一列、縮排於主旨右側。**列身與箭頭分工**：點列身＝在右側面板顯示這一列代表的那封信（主旨串列就是串頭那封），點箭頭才展開／收合；列身的點擊排除勾選框、刪除鍵與箭頭本身。**排序**：預設最新在上，主旨串列以串內最新一封的時間就位，而不是以它顯示的那封 | empty / loading / error / selected / expanded |
| `MailRowBadge` | 列尾小標籤，接在象限名稱之後（例：已加入待辦；同一主旨串有多封已加入時帶數量，例：已加入待辦 2）；展開後的成員列只標自己那封是否已加入；只標狀態，不取代象限名稱 | 有／無 |
| `ScoreReasonList` | 可解釋原因 | empty |
| `QuadrantPicker` | 手動改象限 | default / confirm |
| `TodoList` / `TodoRow` | 待辦 | empty / loading / error / completed |
| `RelatedMails` | 待回待辦右側面板的「相關的郵件」區塊：列出該待辦同一主旨串的所有信，每列可點開看那一封的資訊；排序與郵件清單共用同一組設定 | empty / default |
| `TodoToolbar` | 排序 → 清除已完成 → 待辦狀態：同一列，讀序就是操作序 | default |
| `ScanProgress` | 進度 + 取消 | determinate |
| `ErrorCard` | 錯誤 + 動作 | — |
| `EmptyState` | 空狀態 | — |
| `SettingsForm` | 設定分節 | valid / invalid / saving |
| `PrivacyPanel` | 資料位置與清除 | — |
| `AiConsentDialog` | 白名單預覽 + 確認 | — |
| `PrimaryButton` / `SecondaryButton` / `GhostButton` | 動作 | 全狀態矩陣 |
| `Toast` | 輕量結果 | success / error |

約束：列表行不做重玻璃，避免可讀性下降。玻璃只用在面板與詳情層。

列表行的家具固定為三格：狀態點 → 主旨（＋副行）→ 列尾。列尾是一條橫向堆疊，依序為展開箭頭（僅主旨串列）、象限名稱、小標籤，整組靠列尾對齊、不換行。標籤只標記狀態，不搶走象限名稱的位置。主旨串列尾的標籤帶數量（該主旨串已加入待辦的封數），展開後的成員列只標自己那封是否已加入。

列身與箭頭是兩個不同的落點：箭頭只管展開／收合，列身則把這一列代表的那封信送進右側面板。收合的主旨串列也不例外——它代表串頭那封，所以收合時仍看得到串頭資訊。箭頭最小點擊高度 ≥ 44px（§9.3）。

---

## 4. Visual Layer（Design Token）

**Style anchor**  
Apple macOS / iOS 系統設定與郵件工具列的玻璃材質語彙，搬到 Windows 工具型桌面應用。不做行銷首頁、不做 AI 對話主導介面。

**Palette（Light）**

| Token | Hex | 用途 |
|---|---|---|
| `--bg` | `#F2F2F7` | 視窗底 |
| `--surface` | `#FFFFFF` | 列表/卡片底 |
| `--glass` | `rgba(255,255,255,0.72)` | 玻璃面 |
| `--glass-stroke` | `rgba(255,255,255,0.55)` | 玻璃高光邊 |
| `--ink` | `#1C1C1E` | 主文字 |
| `--ink-muted` | `#6C6C70` | 次要文字 |
| `--separator` | `rgba(60,60,67,0.12)` | 分隔線 |
| `--accent` | `#0A84FF` | 主操作 |
| `--danger` | `#FF3B30` | 錯誤 |
| `--q1` | `#FF453A` | 緊急重要 |
| `--q2` | `#FF9F0A` | 緊急不重要 |
| `--q3` | `#0A84FF` | 重要不緊急 |
| `--q4` | `#8E8E93` | 不重要不緊急 |

**Palette（Dark）**

| Token | Hex |
|---|---|
| `--bg` | `#000000` |
| `--surface` | `#1C1C1E` |
| `--glass` | `rgba(28,28,30,0.72)` |
| `--glass-stroke` | `rgba(255,255,255,0.12)` |
| `--ink` | `#F2F2F7` |
| `--ink-muted` | `#98989D` |
| `--separator` | `rgba(84,84,88,0.6)` |
| `--accent` | `#0A84FF` |
| 象限色 | 同 Light，必要時加亮 |

**Typography**

| Role | Family | Size / Weight |
|---|---|---|
| Title | `Segoe UI Variable Display`, `Segoe UI Semibold`, `Microsoft JhengHei UI` | 22 / 600 |
| Body | `Segoe UI Variable Text`, `Segoe UI`, `Microsoft JhengHei UI`, `Microsoft YaHei UI` | 14 / 400 |
| Caption | 同上 | 12 / 400 |
| Mono | `Cascadia Mono`, `Consolas` | 12 / 400 |

中文優先微軟正黑/雅黑 UI 字型；不用外部字型檔（零安裝）。

**Spacing / Radius / Elevation**

- 基準 4px：4 / 8 / 12 / 16 / 24 / 32  
- Radius：control 8、panel 12、sheet 16  
- Elevation：玻璃面板用 blur + 細邊，不用重陰影  

**Motion**

- 僅列表插入、面板開合、進度，≤ 200ms  
- 尊重 `prefers-reduced-motion`（WPF 側以系統動畫設定或等價開關）  
- 禁止裝飾性持續動畫  

**Glass 實作原則（WPF）**

- 背景：半透明填充 + 系統 DWM blur（可用時），不可用時退化為純半透明面 + 描邊  
- 對比：玻璃面上正文對比 ≥ 4.5:1  
- 列表列、長文、原因文字落在實心 `--surface` 或足夠遮罩，不落在強模糊區  

**UI 不能有 AI 味**

- 不用機器人圖示、不用對話氣泡當主 UI、不用「智能為您…」文案  
- AI 僅在設定與選擇性對話框出現，動詞清楚  

---

## 5. System Architecture

```text
┌─────────────────────────────────────────────┐
│ UI (WPF)                                    │
│  QuadrantBoard / TodoBoard / Settings       │
└──────────────────┬──────────────────────────┘
                   │ ViewModel / 用例呼叫
┌──────────────────▼──────────────────────────┐
│ Application                                 │
│  ScanMailUseCase                            │
│  ClassifyUseCase                            │
│  TodoUseCase                                │
│  SettingsUseCase                            │
└───────┬──────────────────────────┬──────────┘
        │                          │
┌───────▼────────┐        ┌────────▼─────────┐
│ Core Domain    │        │ AI Capability    │
│ RuleEngine     │        │ Interface        │
│ Quadrant       │        └────────┬─────────┘
│ Todo model     │                 │
│ Override       │        ┌────────▼─────────┐
└───────┬────────┘        │ Provider Adapter │
        │                 │ (optional)       │
┌───────▼────────┐        └────────┬─────────┘
│ Adapters       │                 │
│ Outlook COM    │        ┌────────▼─────────┐
│ JsonStore      │        │ Provider API     │
│ Logger         │        └──────────────────┘
└────────────────┘
```

### 5.1 模組邊界

| 模組 | 依賴 | 禁止 |
|---|---|---|
| `Core` | 無 UI、無 COM、無 AI、無磁碟 | 不引用 WPF / Outlook / HTTP |
| `Application` | Core + 介面（Port） | 不直接 `new` COM 物件 |
| `Adapters.Outlook` | COM | 不寫回郵件（M1–M3） |
| `Adapters.Storage` | JSON 檔 | 不存 Token / 信件內文（Log 同理） |
| `Adapters.AI` | HTTP（可選） | 不得讓例外潰到 UI 崩潰 |
| `UI` | Application + Core 型別 | 不寫業務規則 |

### 5.2 外部依賴

- OS：.NET Framework 4.8、Outlook 桌面、DWM（可選玻璃）  
- 人為：使用者設定的 AI Endpoint（可選）  
- 本專案 MVP 不引 NuGet 套件  

### 5.3 平台抽象與擴充點

| 介面 | 目的 | MVP 實作 |
|---|---|---|
| `IMailReader` | 讀郵件 | `OutlookComMailReader`（本機）、`GraphMailReader`（雲端，v1.2.0） |
| `IMailArchiveImporter` | PST 匯入 | `PstImportReader`（可 M2） |
| `IOverrideStore` / `ITodoStore` / `ISettingsStore` | 持久化 | `JsonStore` |
| `IClassifier` | 分類 | `RuleEngine` |
| `IAiCapability` | 增強 | `NullAiCapability` + 可選 HTTP adapter |

換 AI 廠商只動 `Provider Adapter`。換儲存只動 Storage adapter。Core 不變。

換郵件來源同理：`MailReaderFactory` 依來源種類挑實作，`MailSourceRuntime` 決定當下作用中的來源，`ScanMailUseCase` 只認 `IMailReader`。v1.2.0 起 `ProbeInbox`／`TryOpenInOutlook` 也納入此介面，介面外不再有 COM 專屬呼叫。

---

## 6. Data Flow

### 6.1 掃描與分類

```mermaid
sequenceDiagram
  participant U as 使用者
  participant UI as WPF UI
  participant AS as ScanMailUseCase
  participant MR as IMailReader
  participant OL as Outlook COM
  participant RE as RuleEngine
  participant ST as JsonStore

  U->>UI: 選資料夾/天數、開始掃描
  UI->>AS: RunAsync(scope)
  AS->>MR: EnumerateAsync(scope)
  MR->>OL: 只讀取得欄位
  OL-->>MR: MailSummary*
  MR-->>AS: MailSummary stream
  AS->>ST: LoadOverrides()
  loop 每封信
    AS->>RE: Classify(mail, overrides)
    RE-->>AS: ClassificationResult
    AS->>UI: Progress + Batch results
  end
  AS-->>UI: Completed
  UI-->>U: 四象限列表 + 原因
```

**掃描範圍**

- 遞迴走訪所有郵件儲存區（主要信箱 + 封存 PST）中的郵件資料夾。
- 一律排除「刪除的郵件」資料夾及其子資料夾：各儲存區以 `GetDefaultFolder(olFolderDeletedItems)` 的 EntryID 判定，資料夾被改名仍可辨識；另以資料夾名稱（繁中／簡中／英文）作為備援判定。
- 天數：只取 `ReceivedTime` 在 N 天內的郵件。
- 上限：單次掃描最多 2000 封。

### 6.2 唯讀原則

- COM 只呼叫取得屬性/資料夾列舉。  
- 不呼叫 `MarkAsRead`、`Move`、`Flag`、`Save`、`Send`。  
- 「開啟於 Outlook」用 EntryID 以 Inspector 顯示，仍不改內容。  
- 本地 Override / Todo 與 Outlook 郵件屬性分離存放。

### 6.3 更新機制

版本單一來源：repo 根目錄 `VERSION`（例：`1.0.0`）。`build.ps1` 讀取後蓋入組件的 `AssemblyInformationalVersion`，App 以 `AppInfo` 回報，GitHub Release 的標籤與產物名稱用同一個號碼，三者不會漂移。

- 檢查：`CheckForUpdateUseCase` 經由 `IReleaseFeed`（實作 `GitHubReleaseFeed`，讀 `releases/latest`）取得最新版本；`AppVersion` 比較現行版與標籤（接受 `v1.2.3`、忽略 build metadata、拒絕 prerelease 與格式錯誤的標籤）。
- 決策物件：`UpdateOffer` / `UpdateStatus`（`UpToDate`、`UpdateAvailable`、`NotPublished`、`Failed`）。UI 只呈現結果，不做任何判斷。
- 確認：`UpdateSheet` 以 Apple/macOS 語彙的 modal sheet 呈現（標題、版本對照、版本說明、主要按鈕「下載並更新」、次要「稍後」、文字鈕「開啟下載頁面」）。按下主要按鈕前不會發出任何下載。
- 安裝：`UpdateInstaller` 下載至 `%TEMP%` → 驗證 → 產生批次腳本 → 程式結束後替換 exe 並重新啟動。安裝前先做寫入探測，程式所在資料夾不可寫（如 Program Files）時改為只提供下載頁面。
- 啟動時檢查為選項（設定頁「啟動時自動檢查更新」），預設關閉；失敗只寫入 log，不打擾使用者。

資安邊界（全部集中在 `Core/Security/UpdateSecurityPolicy.cs`，可用單元測試驗證）：

- 僅 HTTPS，且主機須在 GitHub 白名單（`github.com`、`api.github.com`、`objects.githubusercontent.com`、`github-releases.githubusercontent.com`、`release-assets.githubusercontent.com`）；主機為精確比對、不做後綴比對（避免 `github.com.evil.example` 這類混淆網域），禁自訂埠、禁網址內帳密，重導後的最終位址需再通過同一檢查。`release-assets.githubusercontent.com` 是實際觀察到下載重導的資產主機；若 GitHub 之後改用別的主機，行為是「拒絕下載」而非信任未知主機（fail closed）。
- 檢查與下載共用同一個 TLS 進入點（`UpdateHttp.EnsureModernTls`，冪等）：程序預設的 `Ssl3|Tls` 會被 GitHub 的 API 與資產主機拒絕；兩者共用可避免「下載只有在同一程序先做過檢查時才成功」的隱性順序相依。
- 只接受資產名稱 `OutlookAiHelper.exe`（純檔名、非路徑）、大小上限 200 MiB，且實際位元組數需與 Release 宣告相符。
- 下載內容需具 Windows 執行檔（MZ）檔頭，不符即刪除暫存檔並中止。
- 批次腳本中的路徑需通過 `IsSafeForUpdaterScript`（擋 `"`、`%`、`&`、`|` 等 cmd 特殊字元），避免指令注入。
- 更新路徑只對 GitHub 發出 GET；不送出任何本機資料、API key 或信件內容。

對應檔案：`Core/Models/AppVersion.cs`、`AppInfo.cs`、`ReleaseInfo.cs`、`UpdateOffer.cs`；`Core/Security/UpdateSecurityPolicy.cs`；`Core/Abstractions/IReleaseFeed.cs`；`Core/Application/CheckForUpdateUseCase.cs`；`Adapters/Update/GitHubReleaseFeed.cs`、`UpdateInstaller.cs`、`UpdateHttp.cs`；`UI/UpdateSheet.cs`、`UI/MainWindow.Update.cs`；`tests/UpdateTests.cs`；`VERSION`；`release.ps1`（建置 → 標籤 → 發佈產物的發佈腳本）。

---

## 7. Data Model

### 7.1 Entities

```text
MailSummary
  EntryId, FolderPath, Subject, FromName, FromAddress,
  ReceivedOn, IsRead, HasFlag, IsMeetingRequest, Importance

ScoreReason
  Code, LabelKey, Delta, Detail

ClassificationResult
  EntryId, Quadrant (Q1..Q4), Score, Reasons[]

Override
  EntryId, Quadrant, UpdatedOn, Source = Manual

TodoItem
  Id, Title, Note, SourceEntryId?, QuadrantHint?,
  Status (Open|Done), CreatedOn, UpdatedOn, DueOn?

AppSettings
  SchemaVersion, Language, FolderScope[], ScanDays,
  VipAddresses[], UrgentKeywords[], ImportantKeywords[],
  AiEnabled, AiProviderConfig? (無金鑰明文；Key 另存 user DPAPI/檔且權限最緊)
  MailSources[], ActiveMailSourceId, SourceChoiceAsked

MailSourceProfile
  Id, Kind (文字：local | graph), DisplayName, AccountUpn?
  （不含任何憑證；refresh token 另存 DPAPI 檔，永不進 settings.json）
```

### 7.2 象限定義（可解釋起點）

預設權重（可在設定調整，必須顯示目前權重）：

| 訊號 | 預設效果 |
|---|---|
| 主旨/內文前 500 字含緊急詞（今天、盡快、ASAP、urgent…） | 緊急 + |
| 會議邀請、工作階段邀請 | 緊急 + |
| 未讀且「直接寄給我」 | 緊急 + 小 |
| VIP 寄件者 | 重要 + |
| 重要詞（合約、付款、客訴、報價…） | 重要 + |
| 寄件者在忽略清單 | 重要 − |
| 本地 Override | 最終象限以此為準，不重算 |

象限閾值：緊急分 ≥ `U` 且重要分 ≥ `I` → Q1，依此類推。  
`ScoreReason` 必須能指出每一條加減分，UI 原文顯示（i18n key）。

### 7.3 Storage Schema（JSON）

路徑：`%LOCALAPPDATA%\OutlookAiHelper\`（不寫系統目錄；可改到使用者資料夾）

| 檔 | 內容 | 預設保存 |
|---|---|---|
| `settings.json` | AppSettings | 直到使用者刪除/重設 |
| `overrides.json` | Override[] | 直到刪除或「清除本地分類記憶」 |
| `todos.json` | TodoItem[] | 直到刪除或匯出後清除 |
| `cache\scan-index.json` | 上次掃描摘要（無內文） | 可清除；可設 TTL 30 天 |
| `secrets\refresh-<sourceId>.dat` | 雲端來源的 refresh token（DPAPI CurrentUser；永不進 settings.json 或 Log） | 直到登出或移除來源 |
| `logs\app-YYYYMMDD.log` | 技術 Log | 滾動保留 7 天 |
| `exports\` | 使用者匯出 | 使用者自管 |

每個 JSON 有 `schemaVersion`。破壞性變更走 `Migrate(from → to)`；無法讀取時把壞檔改名 `*.bad-YYYYMMDD` 再重建，不直接靜默覆寫。

**資料最小化對照**

| 資料 | 為何需要 | 用在哪 | 存多久 | 誰存取 | 可刪 |
|---|---|---|---|---|---|
| 主旨/寄件者/時間 | 分類與列表 | UI/規則 | 快取可清；核心不長期存內文 | 僅本機使用者 | 是 |
| 內文前綴 | 關鍵字 | 僅記憶體分類 | 不落盤（除非使用者開診斷且明示） | 僅本機 | 是 |
| Override | 你的更正 | 分類覆寫 | 使用者可控 | 僅本機 | 是 |
| Todo | 待回待辦 | 待辦頁 | 使用者可控 | 僅本機 | 是 |
| AI 請求 | 選擇性增強 | Provider | 不預設本地保存請求體 | 你 + Provider | 不適用（已送出） |

---

## 8. Directory Structure

```text
outlook-ai-helper/
  DESIGN.md
  README.md
  src/
    OutlookAiHelper/
      OutlookAiHelper.csproj
      App.xaml
      App.xaml.cs
      Core/
        Classification/
          Quadrant.cs
          ScoreReason.cs
          ClassificationResult.cs
          RuleEngine.cs
          RuleOptions.cs
        Models/
          MailSummary.cs
          TodoItem.cs
          Override.cs
          MailSourceProfile.cs
        Abstractions/
          IMailReader.cs
          IClassifier.cs
          IOverrideStore.cs
          ITodoStore.cs
          ISettingsStore.cs
          IAiCapability.cs
      Application/
        ScanMailUseCase.cs
        ClassifyUseCase.cs
        TodoUseCase.cs
        SettingsUseCase.cs
      Adapters/
        Outlook/
          OutlookComMailReader.cs
          OutlookAvailability.cs
        Mail/
          MailReaderFactory.cs
          MailSourceRuntime.cs
          GraphMailReader.cs（v1.2.0 雲端）
        Storage/
          JsonStore.cs
          JsonPaths.cs
          SchemaMigrations.cs
        Ai/
          NullAiCapability.cs
        Logging/
          FileLogger.cs
      UI/
        MainWindow.xaml
        Views/
        ViewModels/
        Controls/
        Converters/
        Themes/
          Tokens.xaml
          Light.xaml
          Dark.xaml
      Localization/
        Strings.zh-TW.resx
        Strings.zh-CN.resx
        Strings.en-US.resx
  tests/
    OutlookAiHelper.Tests/
      RuleEngineTests.cs
      TodoStoreTests.cs
      MigrationTests.cs
```

---

## 9. UI / UX Plan

### 9.1 布局草圖

```text
┌────────┬──────────────────────────────────────────────┐
│ 四象限  │ [資料夾 ▾][天數 ▾] [開始掃描] [重新整理]      │
│ 待回待辦├──────────────────────────────────────────────┤
│ 設定   │ [全部|Q1|Q2|Q3|Q4]                           │
│        ├────────────────────────────┬─────────────────┤
│        │ 郵件列表                    │ 詳情             │
│        │ · 主旨 / 寄件者 / 時間      │ 主旨            │
│        │ · 象限點                    │ 原因列表        │
│        │ · …                         │ [改象限][待辦]  │
└────────┴────────────────────────────┴─────────────────┘
```

### 9.2 主路徑（非技術使用者）

1. 開啟程式 → 看到四象限空狀態：「先掃描收件匣」  
2. 預設資料夾=收件匣、天數=30 → 開始掃描  
3. 依 Q1 先處理；點一筆看「為什麼」  
4. 改錯分類；加一條待回  
5. 到待回待辦勾完成  

首次啟動顯示一次唯讀聲明（可關閉，設定可再看）。

### 9.3 視覺層執行對照

- 全部顏色/字級只從 `Tokens.xaml` 來，頁面不自創 hex  
- Light/Dark 跟隨系統，設定可強制  
- 文案全部走 `t(key)`，含錯誤、空狀態  
- 對比與鍵盤焦點照 Apple HIG；點擊目標用**桌面**數值，不套觸控的 44px 規則。實測：側邊欄文字列 186×38、圖示鈕 28×28、設定齒輪 38×38（另給 `ToolTip` 與 `AutomationProperties.HelpText`），可點區域遠大於指標所需  
- 每個可操作控制項都有可讀名稱（`AutomationProperties.Name`）：圖示按鈕、清單列（列用 `UiKit.FlatListItemStyle("Content.Tag.…")` 綁定自身內容）、以及設定頁每個輸入框與下拉選單（`WithName(…)`）。少了它，輔助工具只唸出容器型別或一個空字串  

### 9.4 Design pass 摘要

```text
SUBJECT   本機 Outlook 四象限與待回清單 — 信任感與可掃讀
COLOR     --bg #F2F2F7 / --ink #1C1C1E / --accent #0A84FF
          象限色 q1 #FF453A q2 #FF9F0A q3 #0A84FF q4 #8E8E93
TYPE      Segoe UI Variable + Microsoft JhengHei UI（系統字）
          標題 22/600，本文 14/400，說明 12
LAYOUT    傳統三欄工作區（導覽 | 列表 | 詳情），密度偏工具不偏卡片牆
SIGNATURE 象限板上的玻璃面板 + 象限色點；列表列保持實心可讀
RISK      玻璃不進密列表列；不用漸層濫用、不用吉祥物/AI 對話殼
```

---

## 10. Test Strategy

| 層級 | 對象 | 方法 | 驗收 |
|---|---|---|---|
| 單元 | RuleEngine | 純 C#，表格驅動關鍵字/VIP/Override | 覆蓋邊界與空輸入 |
| 單元 | JSON 遷移 | 壞檔、舊版 schemaVersion | 不靜默遺失 |
| 單元 | TodoUseCase | 新增/完成/刪除 | 狀態正確 |
| 整合 | ScanMailUseCase | `IMailReader` 假實作 | 可取消、批次、錯誤 |
| 手動 | Outlook COM | 真實 2016/2019/2024/O365 至少兩種 | 唯讀、千封、關 Outlook |
| 手動 | UI 三態 | 斷網/空/錯 | 無空白頁 |
| 手動 | i18n | 切三語 | 錯誤與空狀態也翻 |
| 隱私 | Log/快取 | 搜尋內文關鍵字、Token | 應為 0 |

M1 不把「真實 Outlook 自動化測試」當閘門，先用 Fake 鎖規則，再手動驗證 COM。

---

## 11. Milestone Plan

| 里程碑 | 目標 | 含 | 不含 |
|---|---|---|---|
| **M1** | 唯讀掃描 + 可解釋四象限 | COM 讀取、RuleEngine、四象限 UI、三態、手動改象限、基礎 i18n | 待辦持久化、PST、AI、匯出 |
| **M2** | 待回待辦 + Override 持久化 | todos.json、覆寫持久化、從郵件加待辦 | 自動待辦生成 |
| **M3** | 設定、隱私、匯出/清除 | 設定頁、隱私面板、匯出 JSON | 云同步 |
| **M4** | 可選 AI 增強 | Provider Adapter、白名單預覽、可切換 | 建議回覆（FUTURE） |
| **M5**（v1.2.0） | 多來源：本機＋雲端 Microsoft 365 | 來源模型與接縫、設定頁「來源」區塊、首次啟動選來源、裝置碼登入、`Mail.ReadBasic` 唯讀 | 共用信箱、待辦跨機同步、寫入郵件 |

依序交付；每個里程碑都要能走完「進入 → 核心任務 → 結果 → 錯誤」。

---

## 12. 已對齊決策（2026-09-24）

| 決策 | 內容 | 標籤 |
|---|---|---|
| 關鍵字 | 以**標題（主旨）**比對為主；專案內建預設詞 + 使用者可設定關鍵詞辨識緊急/重要 | [DECISION] |
| 掃描天數 | 預設 30 天，使用者可調整 | [DECISION] |
| 建置環境 | 本機無 MSBuild、無 .NET SDK；有 .NET Framework 4.8 + `csc.exe`。以 `build.ps1` 呼叫 csc 編譯 | [DECISION] |
| UI 實作 | 因無 XAML 編譯管線，WPF 以程式碼建構視窗/元件（玻璃視覺仍依 Token） | [DECISION] |
| 序列化 | 不用第三方 JSON；採 BCL `DataContractJsonSerializer` | [DECISION] |
| 測試 | 不引 NUnit；輕量 `TestRunner` 命令列斷言，`test.ps1` 執行 | [DECISION] |
| 郵件清單列身 | 點列身＝右側面板顯示該列代表的那封信（收合的主旨串列＝串頭那封）；只有列尾箭頭展開／收合 | [DECISION] |
| 郵件排序 | 一組共用排序設定（郵件清單與待辦面板的相關郵件共用）：預設「最新在上」，另有最舊在上、依象限、依寄件者、依主旨 | [DECISION] |
| 相關的郵件 | 待辦面板列出同一主旨串（去掉 Re:/回覆：等前綴後同主旨）的所有信 | [DECISION] |

### 12.1 主旨關鍵字（可設定）

比對範圍：主旨（M1）。詞表可在設定檔 `RuleOptions` 調整。

- 緊急預設：`緊急` `急` `盡快` `立刻` `馬上` `今天` `ASAP` `urgent` `immediate` `deadline` `到期` `截止`  
- 重要預設：`專案` `project` `合約` `報價` `付款` `發票` `客訴` `投訴` `簽核` `核准` `會議` `合約` `報價` `付款` `invoice` `contract` `quote` `approval` `complaint`  

命中即加對應分；可多詞累加有上限。VIP 寄件者仍計入重要分。

### 12.2 設定頁區塊（v1.1.8，v1.2.0 新增「來源」）

設定頁不再是單欄堆疊的說明文字，改成八個卡片區塊，由上面下：來源、掃描範圍、顯示、自動刷新、關鍵字規則、AI 增強、隱私、軟體更新。每塊＝`UiKit.InsetGroup`（區塊標題＋圓角卡片），卡片內每一列是一個設定；欄寬上限 640px（`SettingsColumnWidth`），避免在寬視窗被拉成一行長文。

「來源」排在第一位，因為它決定其餘設定作用在哪一顆信箱上；掃描範圍、關鍵字規則、待辦都隸屬於作用中的來源。

| 區塊 | 標題 key | 內容 |
|---|---|---|
| 1 來源 | `settings.section.source` | 來源清單（本機／雲端）、切換作用中來源、新增／登入／移除 |
| 2 掃描範圍 | `settings.section.scan` | 掃描天數＋掃描範圍說明 |
| 3 顯示 | `settings.section.display` | 介面語言、郵件排序（清單與待辦面板的相關郵件共用同一設定） |
| 4 自動刷新 | `settings.section.refresh` | 新郵件輪詢間隔 |
| 5 關鍵字規則 | `settings.section.rules` | 緊急／重要關鍵字、VIP 寄件者 |
| 6 AI 增強 | `ai.title` | AI 開關、供應者列、供應者編輯器、金鑰落地說明 |
| 7 隱私 | `privacy.title` | 資料位置、匯出／開啟資料夾／清除本機資料 |
| 8 軟體更新 | `update.section` | 啟動時自動檢查更新與更新資訊 |

AI 供應者的編輯方式（設定頁第 5 區塊）：

| 決策 | 內容 | 標籤 |
|---|---|---|
| 供應者列表 | 每個供應者一列（名稱＋URL，右側編集／移除），取代原本的下拉選單；最後一個供應者不可移除 | [DECISION] |
| 編輯器 | URL／KEY／模型只在按「新增」或「編輯」時**就地展開**（非彈窗）。「完成」保留輸入並收合，「取消」丟棄；取消一個剛新增、尚未確認的供應者時，該筆一併移除 | [DECISION] |
| 金鑰欄 | 預設遮蔽（`UiKit.PasswordInput` → `PasswordBox`＋`GlassSecretTemplate`）；勾選顯示切換為明文欄位。收合編輯器時一律回到遮蔽，且只把使用者看得見的那一欄當作真值 | [DECISION] |
| 收合的意義 | 編輯器收合時不做回寫：已存的供應者設定才是真值，欄位只是它的鏡像 | [DECISION] |
| 金鑰落地 | 金鑰只寫入本機設定檔，不寫入 Log（與隱私區塊的說明一致） | [DECISION] |

### 12.3 多來源：本機＋雲端 Microsoft 365（v1.2.0）

需求：除了本機 Outlook，也能讀企業 Microsoft 365 或雲端 Microsoft 帳號；首次使用可設定來源；之後可在設定增減與管理；多台電腦可同時讀同一顆信箱。

#### 決策

| 決策 | 內容 | 標籤 |
|---|---|---|
| 來源模型 | 「來源清單＋同時只有一個作用中來源」。未設定任何來源時自動合成一個本機來源，本機行為與 v1.1.x 完全相同 | [DECISION] |
| 登入 | 自帶裝置碼流程（device code）。**不引 MSAL、不新增任何相依**，沿用既有的 `HttpWebRequest` | [DECISION] |
| Client ID | 內建一組預設註冊（多租戶＋個人帳號），裝好即可用；進階使用者可在設定覆寫成自己的註冊 | [DECISION] |
| 權限 | 唯讀 `Mail.ReadBasic`（官方排除 `body`／`bodyPreview`／`uniqueBody`／`attachments`／擴充屬性）。「管理」不含寫信、移動、刪除 | [DECISION] |
| 帳號類型 | 多租戶＋個人 Microsoft 帳戶，授權端點用 `common` | [DECISION] |
| 多端 | 多台電腦可讀同一顆雲端信箱；**設定與待辦仍為單機**，不做跨機同步 | [DECISION] |
| 共用信箱 | 本版不做 | [DECISION] |
| 憑證落地 | refresh token 以 DPAPI（`CurrentUser`）存 `secrets\refresh-<sourceId>.dat`；不進 `settings.json`、不進 Log | [DECISION] |
| Client ID 的機密性 | client ID／tenant id 是公開資訊、不是機密（公開用戶端沒有 client secret），可寫在程式與文件；只有 refresh token 需要保護 | [DECISION] |

#### 唯讀權限為何足夠

`MailSummary` 只有 10 個中繼資料欄位（`EntryId`、`FolderPath`、`Subject`、`FromName`、`FromAddress`、`ReceivedOn`、`IsRead`、`HasFlag`、`IsMeetingRequest`、`Importance`），全 repo 沒有任何 `Save`／`Move`／`Delete` 呼叫，App 從不讀取內文。因此 `Mail.ReadBasic` 的欄位排除清單不影響任何既有功能；申請較高權限只是擴大授權面。

#### 風險（屬於租用戶政策，實作層無法消除）

| 風險 | 說明 | 對策 |
|---|---|---|
| 租用戶封鎖裝置碼流程 | 條件式存取若擋下裝置碼流程，該路徑即失效 | 無程式面解法；退回本機模式（該機 Outlook 已登入即可讀同一顆信箱） |
| 使用者同意被停用 | `Mail.ReadBasic` 本身 `AdminConsentRequired = No`，但企業常關閉使用者自行同意，仍需 IT 按一次 | 訊息要說清楚這是租用戶政策，不是程式問題 |
| 來源切換與待辦的關聯 | 待辦與手動分類覆寫都以 `EntryId` 為鍵。換來源後舊待辦仍在，但其郵件不在作用中來源，待辦詳情的「相關郵件」會變空；覆寫因鍵不同不會誤套 | 切換來源時提示；`JsonTodoStore` 的遷移邏輯越晚處理越貴 |

#### 接縫（v1.2.0 已完成的部分）

- `IMailReader` 收納 `ProbeInbox`／`TryOpenInOutlook`，介面外不再有 COM 專屬呼叫；`OutlookComMailReader` 全 repo 只剩一處建構。
- `MailSourceProfile`：`Kind` 以**文字**儲存，舊版讀到不認識的種類不會誤判成本機。
- `MailReaderFactory`：未實作的種類回 `null` 而非丟例外；v1.2.0 起依來源與註冊建出雲端讀取器，沒有可用的註冊（或來源 id 不可能是一個檔名）一樣回 `null`，呼叫端一律退回本機。
- `MailSourceRuntime`：五條退化路徑一律退回本機（清單為空／找不到 id／種類未實作／設定為 null／雲端來源沒有可用註冊），因為「沒有可讀信箱」與「Outlook 壞了」在畫面上長得一樣；並記住已建好的讀取器，只在切換來源時重建（否則每次掃描都會重新登入，也會忘記登入曾被拒絕）。
- 設定檔新增 `MailSources`／`ActiveMailSourceId`／`SourceChoiceAsked`；舊設定檔可直接讀（新增欄位向後相容）。
- `Application/GraphSignIn`：裝置碼登入的編排——要裝置碼、輪詢權杖、把重新整理權杖交給加密儲存——已收斂成一個物件，畫面已接上（見下一條）。失敗一律是一個畫面能顯示的字，不是例外（在 UI 執行緒上為了一個可預期的狀況丟例外會讓它變成崩潰）；輪詢時的 HTTP 400 `authorization_pending` 是「繼續等」而不是失敗，`slow_down` 是等更久，只有 `expired_token`／`access_denied` 才是結束。預設 client ID 仍為空字串，所以 `IsImplemented(Graph)` 仍回 `false`，雲端來源還不會出現在畫面上（模組已暫時關閉，見下一條）。
- `Application/SignInReason` 與 `UI/MainWindow.SignIn`：編排器已接上畫面——字詞轉句子自成一塊（只認字詞、只吐句子，所以能在離線測試裡窮舉），畫面那塊只管裝置碼提示、輪詢節奏與登出，輪詢沿用既有的 `DispatcherTimer` 寫法；沒有內建 client ID 時那條來源根本沒有登入可給（`SignInFor` 回 `null`），這整塊不會出現在畫面上。
- 這一版把雲端模組暫時關閉：`MailReaderFactory.GraphReaderAvailable` 為 `false`，所以雲端來源這一版不提供、也不會出現在畫面上（本機來源照常，Outlook 桌面已登入的個人帳號本來就涵蓋在本機模式裡）。開關和它上面的說明放在同一行，要重新開放只改這一行；`tests/CloudModuleHiddenTests.cs` 把「只有註冊、模組關著時仍讀不到」與「這個版本還認得雲端這個種類」兩件事釘住，免得哪天被當成死碼刪掉。

#### 待驗（要有真實帳號才能量測，在此之前不得當成結果陳述）

- `Mail.ReadBasic` 在真實租用戶實際回傳的欄位集合。
- `IsMeetingRequest` 對應到 Graph 的哪個欄位。
- 該租用戶是否封鎖裝置碼流程。

### 12.4 右側郵件資訊面板：按類別分區塊（v1.3.0；v1.3.1 開啟動作移入郵件資訊卡）

右側面板原本是一路往下的欄位文字：郵件資訊與待辦資訊各被接成**一個**多行字串（`寄件者: …`／`時間: …`／`資料夾: …`，待辦則是 `狀態: …`／`象限: …`／`建立時間: …`），三顆動作按鈕散在字串後面。改成「一個問題一張卡」，由上而下：

| 卡片 | 內容 | 收合條件 |
|---|---|---|
| 郵件資訊 | 寄件者、時間、資料夾、象限、狀態（狀態以膠囊呈現，不是一句話），以及「在 Outlook 開啟」 | 選中的是待辦而信不在掃描裡時整塊收合；沒有信可開或與待辦卡片重複時，開啟動作收起 |
| 為什麼在這裡 | 命中原因 | 沒有原因時整塊收合（`SetReasonBlockVisible`） |
| 有關的信 | 同一主旨串的其他信，排序與郵件清單共用同一組設定 | 主題只有一封信時收合（`related.Count <= 1`），且只在有選中待辦時才有 |
| 待回待辦 | 狀態、到期、建立時間、更新時間、備註、處理狀態 | 選中的是純郵件時整塊收合 |
| 操作 | 手動改象限、加入待辦、AI 說明 | 依選取內容決定哪一顆可用 |

- 區塊建構子＝`DetailBlock(標題, 內容)`；面板整體包在 `ScrollViewer`（`_detailScroll`）裡，卡片比視窗高時才出現捲軸。
- 值為空的欄位不畫那一列；`_detailMeta` 與待辦的串接字串都已移除，改成左標籤右值各自一列。
- 動作按鈕由三顆散落的按鈕合成一張「操作」卡（`BuildActionBlock`）。
- 新增字串鍵：`detail.section.mail`（郵件資訊）、`detail.section.actions`（操作），三語齊備。
- 「在 Outlook 開啟」原本長在「操作」卡，v1.3.1 起移入郵件資訊卡（`BuildMailOpenButton`）：卡片顯示哪一封就開哪一封，關聯郵件被點擊後也跟著換到那一封。`UpdateMailOpenButton` 依 `_mailFactsEntryId` 決定可用與否，而當卡片顯示的正是這張待辦自己的來源信（`_detailTodo.SourceEntryId`）時，待辦卡片已經有同一顆按鈕，這顆就收起，不印兩顆做同一件事的按鈕。
- 這一版的面板程式碼放在新的 partial 檔 `UI/MainWindow.Detail.cs`；`build.ps1` 以萬用字元收源碼，所以不必動建置設定，測試數不變（209）。
- 面板 harness 進版控（這一節的驗證方式，任何人都能重跑）：`selftest/PanelHarness.cs` 是六次狀態的驅動程式，`selftest/panel/panel.txt` 是各次斷言的字串、`selftest/panel/fixture.txt` 是畫面文字（中文放資料檔，避免 .cs 直寫中文被編譯器用主控台編碼讀壞）；執行 `powershell -NoProfile -ExecutionPolicy Bypass -File selftest\run_panel.ps1` 不需參數，會先確認 `dist\OutlookAiHelper.exe` 不比原始碼舊、編譯 harness、跑完六種狀態，最後印出「N/M 條斷言在畫面上」與 exit code，並在 `selftest\` 留下 `panel-out.txt` 與八張 `panel-out-*.png`（輸出檔已列進 `.gitignore`）。

**驗證方式**：右側面板在沒有點選前列是收合的，而合成點擊需要前景——無人桌面給不了，所以那一條 smoke 檢查一直是 SKIP。改用**同處理序**（in-process）harness：把視窗開在同一個處理序裡，直接呼叫視窗自己的選取處理常式（`ShowTodoDetail` / `OnMailSelectedDetail` / `OnRelatedMailClick`），再走視覺樹讀回畫面上真正看得見的文字。六種狀態共 84 條斷言：待辦選取且信在掃描內、待辦選取但信不在掃描內、從郵件清單點一列（沒有待辦）、待辦的主題只有自己一封信（主題卡必須收合而不是空著）、待辦的主題中的另一封信被點擊（郵件資訊卡的那個開啟動作必須跟著換到那一封，harness 另外直接檢查它指向的那個 entry id）、第六種檢查兩個清單各自的排序下拉都在畫面上、選項與程式宣告的一致、且開在設定記住的那個排序（只讀：在 harness 裡改下拉會寫到使用者的真實設定檔）。同一輪也在 900px 最窄視窗下確認兩個排序下拉與「開始掃描」都仍在視窗內，並直接量到工具列由一列長成兩列（98px → 146px）——換行是量到的，不是假設的。

### 12.5 兩種清單的排序區塊、待回待辦的同主題整合（未發行）

排序引擎其實早就有了（`Application/MailOrdering.cs` 五種；`Core/Models/TodoOrdering.cs` 這一版補上「最新優先」共六種），郵件清單也一直都是「時間降序」——問題是使用者看不到也改不到：郵件清單的排序只長在設定頁，改完要按「儲存」才生效；待回待辦雖然有排序下拉，卻在工具列最左邊，而且每次重開都回到預設。這一版把排序搬到清單自己身上：

| 位置 | 控制項 | 預設 | 是否記住 |
|---|---|---|---|
| 郵件清單工具列（開始掃描／取消右側） | 排序下拉，自成一個 glass 群組 | 最新在上 | 是（`AppSettings.MailSort`，與設定頁共用同一個設定） |
| 待回待辦工具列（「全部」右側、新增待辦下方） | 排序下拉，自成一個 glass 群組 | 最新優先 | 是（`AppSettings.TodoSort`，這一版新增） |

- 兩個下拉各是一個 `UiKit.ToolGroup`（玻璃膠囊），與掃描群組、天數群組並排；郵件清單工具列改成**會換行**，視窗縮窄或右側面板打開時不再把「開始掃描」擠掉，也不再被裁掉。
- 待回待辦新增 `TodoSortMode.NewestFirst`（最新優先），`TodoOrdering.DefaultMode` 由「加入順序」改成它——所以「預設以最新時間降序」是預設值真的變了，不是只把下拉選到某個位置。舊設定檔沒有這個鍵時讀回 `AppSettings.DefaultTodoSort`，向後相容。
- 改排序就立刻重畫（郵件清單與其右側的相關郵件、待回待辦清單都一樣），不必重新掃描，也不必按儲存。

**同主題整合**（待回待辦）：`Application/TodoTopics.cs`

| 欄位 | 內容 |
|---|---|
| 主題的認定 | `MailThreads.KeyOfSubject`（去掉 RE:/FW: 這類回覆前綴，忽略大小寫與多餘空白）——**與郵件清單共用同一條規則**，所以「同一主題」在兩頁指的是同一件事，不是兩套定義 |
| 誰當代表 | 使用者選的排序中排最前面的那一項（`Lead`）；主題在清單裡的位置也由它決定，主題內的順序就是清單的順序 |
| 主題列 | 代表那列 ＋ 一條細線：「同主題 N 項」與「展開／收合」；展開後其餘成員各自成列並縮排，勾選、刪除、點擊預覽的行為與一般列相同 |
| 不整合的情形 | 標題無法構成主題（空白、只有一個字）→ 永遠自己一列；**只有一項的主題不算主題**（`IsTopic` ＝ 有鍵且成員 > 1），所以一般待辦不會長出「同主題 1 項」和一個沒東西可展開的按鈕 |
| 已完成與未完成 | 同一主旨的已完成與未完成共用一個鍵（一個主題被回掉了不代表要換一個主題），統計時才分開數 |
| 整合不改變資料 | 摺疊只在畫面上發生，不動待辦本身，也不動已存的來源信 entry id |

**小結行**（看數字就知道現在是什麼狀態）：

- 郵件清單下方：`共 {0} 封 · {1} 個主旨串 · {2} 封已加入待辦`（`mail.list.summary`）後面再接一段「· 排序：{0}」（`mail.list.order`，由 `MailOrdering.LabelKey` 取當前排序的名字）。
- 待回待辦下方：`共 {0} 項 · {1} 個同主題`（`todo.list.summary`），只算真的整合起來的主題，不會把單獨一項算成一個群組。

- 新字串鍵（三語齊備）：`todo.sort.newest`（最新優先）、`todo.topic.count`（同主題 {0} 項）、`todo.topic.more`／`todo.topic.less`（展開／收合）、`todo.list.summary`、`mail.list.order`。

**驗證方式**：`Application/TodoTopics.cs` 不碰 WPF，所以整合規則可以直接量：`tests/TodoTopicTests.cs` 十二條（同主題成一列、不同主題各一列、大小寫與空白不拆、只有一項不算主題、無法構成主題者永遠單獨、主導順序由使用者選的排序決定、已完成仍屬於它的主題且不被算成未完成、整合不增不減、統計只算真的整合、空清單不炸），註冊在 `tests/TestRunner.cs`，測試數 209 → 221。畫面仍由面板 harness 量：`selftest/run_panel.ps1` 六種狀態 84/84、exit 0。

### 12.6 清單列的版面、郵件摘要與同主題列的展開圖示（未發行）

- **排序落點（修正）**：`MailThreads` 的折疊沿用呼叫端給的順序，而 `MailOrdering.SortRows` 只決定「列」排在哪，不重排列裡的信；所以列的位置看最新一封、列首卻是原本排第一的那封。`MailOrdering.BuildRows` 現在一律先 `SortMails` 再折疊：時間模式下列首就是最新一封，象限模式仍由分數決定列首。
- **側欄寬度**：`SidebarWidth` 由 220pt 收成 200pt（欄位宣告與 `UpdatePanelWidths()` 共用同一個常數），清單因此多出 20pt；視窗最小寬度維持 900。
- **一列的結構**：主旨最多兩行（`TextWrapping=Wrap` 加高度上限）；其下是本機信的開頭（`MailSummary.Preview`，由 `OutlookComMailReader.MapItem` 取 `Body` 前段並壓成單行）；寄件者與時間收成一行。
- **郵件摘要的界線**：本機 COM 讀信時本來就拿得到 `Body`，所以摘要在本機來源不需要任何新權限；雲端 Graph 這條路只要求 `Mail.ReadBasic`，拿不到內文，因此雲端來源的列沒有摘要——要有的話得把範圍放寬，這一版沒有做。
- **同主題的展開圖示**：`BuildTodoTopicRow` 的主要信那一列改用 `UiKit.ChevronGlyph`（既有元件），放在「刪除」左側。

**驗證**：`tests/MailOrderingTests.cs` 改寫三條（列首在最新優先下是最新一封、同主題內不把舊信排在新信之上、非時間模式仍由分數決定列首），測試 221 → 223。面板 harness 多一組列斷言：郵件列看得到摘要、主旨可兩行、待辦列看得到摘要、展開圖示在刪除左側（4/4），原有六種狀態仍 84/84、exit 0。
