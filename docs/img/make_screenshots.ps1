# 用 Excel 從練習用範例檔拍手冊截圖（需要本機有 Excel，以及 Python 套件 PyMuPDF、Pillow）
# 用法：在 repo 根目錄執行 .\docs\img\make_screenshots.ps1
# 只標色、不存檔，範例檔本身不會被改到。範例檔改版後重跑一次即可更新截圖。
# 不用 CopyPicture：有些環境的 Excel 拿不到剪貼簿，改成匯出 PDF 再轉 PNG（還能帶出列號欄號）。

$ErrorActionPreference = "Stop"
$root = (Get-Item .).FullName
$ex = Join-Path $root "examples"
$out = Join-Path $root "docs\img"
$p1 = Join-Path $ex "練習用-第一期估驗計價.xlsx"
$p2 = Join-Path $ex "練習用-第二期估驗計價.xlsx"
$YELLOW = 65535

function Shot($wb, $sheet, $range, $highlight, $file, [bool]$formulas = $false, [string]$hideCols = "") {
    $ws = $wb.Worksheets.Item($sheet)
    $ws.Activate()
    $wb.Windows.Item(1).DisplayFormulas = $formulas
    if ($hideCols) { $ws.Range($hideCols).EntireColumn.Hidden = $true }
    foreach ($h in $highlight) { $ws.Range($h).Interior.Color = $YELLOW }
    $ps = $ws.PageSetup
    $ps.PrintArea = $range
    $ps.PrintHeadings = $true
    $ps.PrintGridlines = $true
    $ps.Orientation = 2          # 橫向
    $ps.Zoom = $false
    $ps.FitToPagesWide = 1
    $ps.FitToPagesTall = 1
    $pdf = Join-Path $out ("_tmp_" + [IO.Path]::GetFileNameWithoutExtension($file) + ".pdf")
    $ws.ExportAsFixedFormat(0, $pdf)
    foreach ($h in $highlight) { $ws.Range($h).Interior.ColorIndex = -4142 }
    if ($hideCols) { $ws.Range($hideCols).EntireColumn.Hidden = $false }
    $wb.Windows.Item(1).DisplayFormulas = $false
    $env:PYTHONIOENCODING = "utf-8"
    python -X utf8 (Join-Path $out "pdf_to_png.py") $pdf (Join-Path $out $file)
    Remove-Item $pdf
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $wb1 = $excel.Workbooks.Open($p1, 0, $true)
    $wb2 = $excel.Workbooks.Open($p2, 0, $true)

    # 範例長相
    Shot $wb1 "請款明細表" "A1:L12" @() "01-請款明細表.png"
    # 錯誤 A：前一期物調累計 12,000，本期前期累計還是 12,500
    Shot $wb1 "請款明細表" "A4:L12" @("L11", "L12") "02-A-第一期累計.png"
    Shot $wb2 "請款明細表" "A4:L12" @("H11", "H12") "03-A-第二期前期累計.png"
    # 錯誤 B：顯示公式，J7 是公式、J8 被打成固定值 30（隱藏 C~I 欄，公式才不會擠到看不清楚）
    Shot $wb1 "計價總表" "A4:K10" @("J8") "04-B-固定值誤植.png" $true "C:I"
    # 錯誤 C：同一個數字，第一期顯示 0.37、第二期顯示 0.367188
    Shot $wb1 "計價總表" "A4:K10" @("J7") "05-C-第一期顯示.png"
    Shot $wb2 "計價總表" "A4:K10" @("F7") "06-C-第二期顯示.png"

    $wb1.Close($false)
    $wb2.Close($false)
} finally {
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}

# 截圖更新了，單檔版的快速上手也要重新打包，不然傳出去的還是舊圖
python -X utf8 (Join-Path $root "docs\build_quickstart.py")
