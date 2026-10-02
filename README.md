# Outlook AI 助手（OutlookAiHelper）

Windows 桌面小工具（WPF、.NET Framework 4.x，單一執行檔）。用可解釋的規則掃描**本機 Outlook** 的郵件，分成四個象限（急件／重要／待辦／其他），可選擇用 AI 產生摘要與分類建議，並內建自我更新。

- 只讀取郵件：不移動、不刪除、不標記、不寄送。
- 掃描預設**排除「刪除的郵件」**（含子資料夾）。
- AI 預設關閉；開啟後才把**你選取的那封信**送到**你自己設定的端點**（自帶 API key）。
- 設定、API key、待辦、覆寫紀錄只存在本機 `%LOCALAPPDATA%\OutlookAiHelper`，不在這個 repo 裡。

## 主要功能

| 頁面 | 內容 |
| --- | --- |
| 四象限 | 依規則分數列出郵件、顯示評分理由，可手動覆寫象限、開啟於 Outlook |
| 待辦 | 從郵件加入待辦、完成／刪除，資料存本機 JSON |
| 設定 | 掃描天數、關鍵字、VIP、語言（繁中／簡中／英文）、AI 供應商與模型、軟體更新 |

## 隱私邊界

1. Outlook 端只呼叫取得屬性與列舉資料夾的 COM API；唯讀原則見 `DESIGN.md` 6.2。
2. AI 為明示開啟＋逐封送出的設計；金鑰存放於使用者設定檔，不進版控。
3. 更新功能只對 GitHub 發出 GET（見 `DESIGN.md` 6.3），不挾帶任何本機資料。
4. `.gitignore` 排除 `dist/`、`obj/`、log、`settings.json` / `todos.json` / `overrides.json`、金鑰與憑證。

## 建置

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File build.ps1
```

需求：Windows 10/11、.NET Framework 4.x（使用系統內建的 `csc.exe`）、桌面版 Outlook。

輸出：

- `dist\OutlookAiHelper.exe` — 應用程式
- `dist\OutlookAiHelper.Tests.exe` — 測試主機，`build.ps1` 會直接執行並在失敗時中止建置

## 版本與發佈

`VERSION` 是唯一的版本來源（例：`1.0.0`）：

1. `build.ps1` 把版本蓋入組件的 `AssemblyInformationalVersion`，App 以 `AppInfo` 回報。
2. App 內「設定 → 軟體更新 → 檢查更新」會比對 GitHub Release 的最新標籤。
3. `release.ps1` 負責發佈：建置 → 建立 `v<版本>` 標籤 → `gh release create` 並附上 `OutlookAiHelper.exe`。

產物名稱固定為 `OutlookAiHelper.exe`，這是更新程式的白名單條件之一。

## 更新機制的安全設計

- 僅 HTTPS，且只信任 GitHub 網域（`github.com`、`api.github.com`、`objects.githubusercontent.com`）；重導後的最終網址需再通過同一檢查。
- 只接受檔名 `OutlookAiHelper.exe`、大小上限 200 MiB，並比對 Release 宣告的大小與 Windows 執行檔（MZ）檔頭。
- 只有按下確認框的「下載並更新」才會下載；下載到 `%TEMP%`，由批次腳本在程式結束後替換執行檔並重新啟動。
- 程式所在資料夾不可寫時不嘗試安裝，改為導向 Release 頁面。
- 「啟動時自動檢查更新」預設關閉；檢查失敗僅寫入 log，不彈錯。

## 檔案結構

```text
build.ps1                    建置 + 測試（含版本戳記）
release.ps1                  發佈：標籤 + GitHub Release 產物
VERSION                      版本單一來源
DESIGN.md                    規格與設計決策
DESIGN_NOTES.md              視覺走向筆記
assets/app-icon.png          圖示來源
tools/make_icon.py           產生 src/OutlookAiHelper/app.ico
src/OutlookAiHelper/
  Program.cs                 進入點
  Core/                      Models / Classification / Security / Abstractions / Application
  Application/               使用案例（掃描、分類、隱私、待辦）
  Adapters/                  Outlook COM、Storage(JSON)、Ai(HTTP)、Update(GitHub)、Logging
  UI/                        MainWindow、UiKit、Theme、UpdateSheet
  Localization/Strings.cs    繁中／簡中／英文
tests/                       建置後執行、無框架的測試（規則、儲存、掃描範圍、更新）
```

## 測試

```powershell
dist\OutlookAiHelper.Tests.exe
```

涵蓋：分類規則、待辦儲存、設定隱私與毀損隔離、掃描範圍（刪除的郵件排除）、更新版本比較與更新安全政策 / Release 解析。目前 49 項全部通過。

## 授權

MIT License，詳見 [LICENSE](LICENSE)。
