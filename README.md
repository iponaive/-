# 專案名稱：Word 批次轉 PDF 暨自動化頁面處理工具 (PowerShell)

## 專案簡介
本工具透過 PowerShell 驅動 Microsoft Word COM 介面，實現大量文件的自動化處理。程式能精準控制文件排版，在批次轉檔 PDF 的過程中，自動完成隱藏修訂追蹤、首頁植入 PBC 印章以及動態頁尾生成，確保輸出文件符合正式歸檔與外部查閱標準。

## 核心功能與技術邏輯
* **環境自動化控制**：
    * 自動隱藏文件中的修訂與註解（Revisions/Comments），確保 PDF 呈現最終版本。
    * 精確控制頁面邊距（BottomMargin/FooterDistance），優化版面空間。
* **首頁印章精確定位**：
    * 透過「相對於頁面（Page-relative）」而非邊界的絕對座標定位。
    * 將 **PBC 印章** 自動放置於首頁左上角（Top: 10, Left: 20），確保不干擾正文內容。
* **智慧檔案解析與頁尾生成**：
    * **檔名解析**：自動擷取底線（_）前的字串作為文件識別標題。
    * **動態頁尾**：
        * 左側：標註 `By Ariel Lin`。
        * 右側：動態生成 `[解析標題] P.[自動分頁編號]`。
        * 字體設定：統一使用 Times New Roman，Size 6，確保專業度。

## 程式執行範例
* **原始檔名**：`ReportA_20260304.docx`
* **處理邏輯**：
    1. 擷取標題為 `ReportA`。
    2. 首頁左上角植入 3.0cm 寬度之 PBC 印章。
    3. 頁尾生成：`By Ariel Lin` [Tab] `ReportA P.1`。
* **輸出結果**：同路徑下之 `ReportA_20260304.pdf`。

## 技術實作環境
* **語言**：PowerShell
* **依賴組件**：Microsoft Word (New-Object -ComObject Word.Application)
* **核心技術**：
    * **Word DOM 控制**：操作 Sections, Footers, Shapes 等對象。
    * **單位換算**：精確執行公分（cm）與點數（Points）之換算（1 cm = 28.35 pt）。
    * **欄位程式碼（Field Codes）**：動態插入 `wdFieldPage` (33) 實現自動頁碼。

## 執行流程
1. 將 `pbc_stamp.png` 圖檔與腳本放置於同一資料夾。
2. 將待處理的 `.docx` 檔案放入該資料夾。
3. 執行腳本，程式將自動開啟背景 Word 引擎進行批次處理。
4. 完成後會輸出綠色成功訊息並產出 PDF 檔案。

