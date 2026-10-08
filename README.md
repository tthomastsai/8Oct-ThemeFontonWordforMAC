# 主題字型 — Word 增益集（Mac）

修改文件的佈景主題字型（Headings / Body，含中文字型），等同 Windows 版 Word 的「自訂字型」。
液態玻璃介面，字型欄位可直接輸入搜尋，每個選項用自己的字型顯示。

網頁由 GitHub Pages 提供：https://tthomastsai.github.io/8Oct-ThemeFontonWordforMAC/

## 安裝（一次性，不需要開終端機伺服器）

1. 下載本倉庫的 [`manifest.xml`](manifest.xml)（點進去後按右上角的下載鈕）
2. 放進 Word 的側載資料夾：
   ```bash
   mkdir -p ~/Library/Containers/com.microsoft.Word/Data/Documents/wef
   cp ~/Downloads/manifest.xml ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/
   ```
3. 重新啟動 Word，開啟任一文件：
   - 功能區 **Home（常用）** 最右側會出現「主題字型」按鈕；或
   - **Insert → Add-ins → My Add-ins**，在「Developer Add-ins」區找到「主題字型」。

## 使用

1. 點「主題字型」開啟側欄（會自動讀入目前的主題字型）
2. 在四個欄位挑選或輸入字型（輸入 `ming`、`hei`、`正黑`、`times` 即時篩選）
3. 按「套用到此文件」，完成後會重新讀回文件確認字型已更新

**第一次使用請先勾選「套用前先在新視窗開啟原文件備份」**，並可用 Cmd + Z 復原。

## 運作方式

Office.js 沒有直接修改主題字型的 API，所以流程是：

1. `getFileAsync` 取得整份文件（.docx）
2. 在側欄內用 JSZip 修改 `word/theme/theme1.xml`
3. `insertFileFromBase64("Replace", { importTheme: true, importStyles: false })` 以修改後的版本取代目前文件
4. 重新讀回文件，確認主題字型已改變

## 限制

- 需要 WordApi 1.5（Word 2019 / Microsoft 365 for Mac）。
- 「取代整份文件」比直接改檔案更動較大，建議先用備份文件試過。
- Office.js 無法列出電腦已安裝的字型，選單是常用字型清單；欄位仍可輸入任何字型名稱。
- 只有使用「主題字型」的文字會連動；被手動指定字型的文字，請全選後按 `Ctrl + Space` 清除。

## 檔案

| 檔案 | 用途 |
|---|---|
| `manifest.xml` | 增益集描述檔（指向 GitHub Pages），放進 Word 的 `wef` 資料夾 |
| `taskpane.html` / `taskpane.js` | 側欄介面與 Word 互動 |
| `core.js` | 主題字型修改邏輯 |
| `vendor/jszip.min.js` | JSZip |
| `assets/` | 功能區圖示 |
| `dev/` | 本機開發用：`server.py` + `manifest.localhost.xml`（`https://localhost:3000`），一般使用不需要 |

## 維護者：啟用 GitHub Pages

倉庫需為 Public → Settings → Pages → Source 選 **Deploy from a branch** → Branch 選 `main`、資料夾 `/ (root)` → Save。
約一分鐘後 `https://tthomastsai.github.io/8Oct-ThemeFontonWordforMAC/taskpane.html` 會可以開啟。
