# 主題字型 — Python 離線版

修改 Word 文件的佈景主題字型（Headings / Body，含中文字型），等同 Windows 版 Word 的「自訂字型」。
**完全離線**，直接改 .docx 內的主題檔，內文與其他部分原封不動，並可讀取本機所有字型。

> Word 增益集（線上版）在 `main` 分支。這個分支是獨立的 Python 版本。

## 三種輸出方式

| 按鈕 | 說明 |
|---|---|
| **覆寫檔案** | 直接修改選取的 .docx（先寫暫存檔，成功後才取代；若文件仍在 Word 中開啟會拒絕執行） |
| **建立新檔…** | 把修改後的副本另存成新檔，原檔不變 |
| **建立新文件…** | 不需要原檔，依設定的字型建立一份空白的新文件（A4，Title / Heading 1–3 走標題字型，內文走內文字型） |

字型欄位**留空 = 不修改該項**。欄位可直接輸入文字搜尋本機字型（例如 `ming`、`hei`、`times`）。

## 使用

```bash
# 基本介面（內建 Tk）
python3 theme_fonts.py

# 液態玻璃介面（建議）
pip3 install PySide6
python3 theme_fonts.py
```

指令模式：

```bash
# 覆寫
python3 theme_fonts.py 論文.docx --latin "Times New Roman" --ea "標楷體"
# 建立新檔
python3 theme_fonts.py 論文.docx --latin Arial --ea "PingFang TC" -o 新檔.docx
# 建立新文件
python3 theme_fonts.py --new-blank 空白.docx --major-latin Arial --minor-latin "Times New Roman"
```

標題與內文可分開設定：`--major-latin` / `--major-ea`（標題）、`--minor-latin` / `--minor-ea`（內文）。

## 注意

- 只有使用「主題字型」的文字會跟著改；被手動指定字型的文字，請在 Word 全選後按 `Ctrl + Space` 清除。
- 選單列出的是這台電腦的系統字型；Office 自己附帶的字型（如 PMingLiU）可能不在清單內，直接輸入名稱即可。
- 覆寫沒有備份可還原，重要文件請先複製一份，或改用「建立新檔」。
