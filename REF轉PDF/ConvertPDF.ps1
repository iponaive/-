$word_app = New-Object -ComObject Word.Application
$word_app.Visible = $false

# 單位換算：1 公分 = 28.35 點 (Points)
$cmToPoints = 28.35

# 設定印章圖檔路徑
$stampPath = Join-Path $PSScriptRoot "pbc_stamp.png"

$files = Get-ChildItem -Path $PSScriptRoot -Filter *.docx

foreach ($file in $files) {
    $fileNamePrefix = $file.Name.Split('_')[0]
    Write-Host "正在處理: $fileNamePrefix ..." -ForegroundColor Cyan
    $doc = $word_app.Documents.Open($file.FullName)
    
    # 1. 顯示設定
    $word_app.ActiveWindow.View.ShowRevisionsAndComments = $false
    $word_app.ActiveWindow.View.RevisionsView = 0 

    # 2. 頁面邊界設定
    $section = $doc.Sections.Item(1)
    $section.PageSetup.BottomMargin = 15
    $section.PageSetup.FooterDistance = 16

    # --- 3. 加入 PBC 印章 (移動至紙張最頂端死角，不擋到文字) ---
    if (Test-Path $stampPath) {
        $shape = $doc.Shapes.AddPicture($stampPath)
        $shape.LockAspectRatio = $true
        
        # 設定寬度為 3.0cm
        $shape.Width = 3.0 * $cmToPoints 
        
        # 維持浮動於文字上方 (wdWrapFront = 3)
        $shape.WrapFormat.Type = 3 
        
        # 修改定位參考：相對於「整個頁面 (Page = 1)」而不是邊界
        $shape.RelativeHorizontalPosition = 1 # wdRelativeHorizontalPositionPage
        $shape.RelativeVerticalPosition = 1   # wdRelativeVerticalPositionPage
        
        # 設定絕對位置 (單位：點)
        # Left 20, Top 10 會讓它出現在紙張左上角的極邊緣空白處
        $shape.Left = 20  
        $shape.Top = 10  
        
    } else {
        Write-Host "找不到印章圖檔: $stampPath" -ForegroundColor Yellow
    }

    # --- 4. 設定頁尾 ---
    $footerTypes = @(1, 2) 
    foreach ($type in $footerTypes) {
        $footer = $section.Footers.Item($type)
        $footer.Range.Text = "" 
        $footer.Range.ParagraphFormat.TabStops.ClearAll()
        
        $rightPos = $section.PageSetup.PageWidth - $section.PageSetup.LeftMargin - $section.PageSetup.RightMargin
        $footer.Range.ParagraphFormat.TabStops.Add($rightPos, 2)
        
        $footer.Range.Font.Name = "Times New Roman"
        $footer.Range.Font.Size = 6
        $footer.Range.ParagraphFormat.Alignment = 0 
        $footer.Range.Text = "By Ariel Lin" + [char]9 + "$fileNamePrefix P."
        
        $rangePage = $footer.Range
        $rangePage.Collapse(0) 
        $doc.Fields.Add($rangePage, 33) 
    }

    $pdf_filename = $file.FullName -replace '\.docx$', '.pdf'
    $doc.SaveAs([ref] $pdf_filename, [ref] 17)
    $doc.Close($false)
}

$word_app.Quit()
Write-Host "完成！印章已放置於左上角空白處。" -ForegroundColor Green
Pause   