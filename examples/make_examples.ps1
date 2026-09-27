# 重新產生練習用範例檔（需要本機有 Excel）
# 用法：在 repo 根目錄執行 .\examples\make_examples.ps1
# openpyxl 寫出的公式沒有計算結果，核算腳本會讀到空值，所以要用 Excel 開啟存檔一次。

$ErrorActionPreference = "Stop"
$dir = Join-Path (Get-Item .).FullName "examples"
$env:PYTHONIOENCODING = "utf-8"
python -X utf8 (Join-Path $dir "make_examples.py")
if ($LASTEXITCODE -ne 0) { exit 1 }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    # 檔名不寫在這裡：PowerShell 5.1 會把無 BOM 腳本裡的中文讀成亂碼
    foreach ($f in Get-ChildItem -Path $dir -Filter "*.xlsx") {
        $wb = $excel.Workbooks.Open($f.FullName)
        $excel.CalculateFull()
        $wb.Save()
        $wb.Close()
        Write-Host ("Recalculated with Excel: " + $f.Name)
    }
} finally {
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
