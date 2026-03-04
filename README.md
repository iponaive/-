# 專案名稱：Word 批次轉 PDF 暨自動化頁面處理工具![PowerShell](https://img.shields.io/badge/PowerShell-%235391FE.svg?style=flat&logo=powershell&logoColor=white)


## 功能演示 (Demo)
<img width="1731" height="253" alt="image" src="https://github.com/user-attachments/assets/56d2c843-40b6-47da-8821-274b4f9717f0" />


## 專案簡介
本工具透過 PowerShell 驅動 Word COM 介面，自動化執行 .docx 檔案轉 PDF 之流程。處理過程包含文件清理（刪除註解與隱藏修訂）、首頁植入 PBC 印章以及動態頁尾生成。

## 核心功能
* **文件清理與標準化**：
    * 執行 `Comments.DeleteAll()` 刪除文件中所有註解。
    * 設定 `ShowRevisionsAndComments = $false` 隱藏所有修訂追蹤。
    * 自動設定頁面邊界（BottomMargin: 15, FooterDistance: 16）。
* **首頁印章植入**：
    * 於第一頁左上角插入 `pbc_stamp.png`（寬度設定為 3.0cm）。
    * 採用相對於頁面（Page-relative）的絕對座標定位（Top: 10, Left: 20）。
* **檔案解析與頁尾生成**：
    * **解析邏輯**：擷取原始檔名中底線（_）之前的字串作為標題。
    * **頁尾配置**：
        * 左側：固定文字 `By Ariel Lin`。
        * 右側：顯示 `[解析標題] P.[自動分頁碼]`。
        * 字體：Times New Roman，Size 6。

## 邏輯範例
* **輸入檔案**：`ProjectA_2026_v1.docx`
* **處理程序**：
    1. 刪除所有註解並隱藏修訂軌跡。
    2. 擷取 `ProjectA` 作為識別碼。
    3. 在首頁左上角空白處放置 PBC 印章。
    4. 插入頁尾：`By Ariel Lin` [Tab] `ProjectA P.1`。
* **輸出檔案**：`ProjectA_2026_v1.pdf`。

## 技術實作
* **開發語言**：PowerShell*  
* **核心技術點**：
    * **Word COM 物件**：使用 `New-Object -ComObject Word.Application` 進行背景處理。
    * **座標換算**：定義 `$cmToPoints = 28.35` 進行長度單位轉換。
    * **頁面對象模型**：遍歷 `Sections.Footers` 並操作 `Shapes.AddPicture` 進行圖文編排。
    * **欄位碼應用**：使用 `wdFieldPage` (33) 插入動態頁碼。

## 執行流程
1. 將 `pbc_stamp.png` 與腳本置於同一資料夾。
2. 將待處理的 Word 文件放入該資料夾。
3. 執行 PowerShell 腳本。
4. 於同路徑產出處理後的 PDF 檔案。
