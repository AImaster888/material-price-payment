# material-price-payment skill - 一鍵安裝
# 用法：在 repo 根目錄執行 .\setup.ps1
#
# 本腳本不會修改 repo 內任何檔案（保留乾淨可發布狀態）。
# 只做兩件事：
#   1. 安裝 Python 依賴（openpyxl）
#   2. 把 SKILL.md 處理過的副本（<PROJECT_PATH> 換成實際路徑）安裝到
#      ~/.claude/skills/material-price-payment/

$ErrorActionPreference = "Stop"

$root = (Get-Item .).FullName
Write-Host ""
Write-Host "Repo 路徑：$root" -ForegroundColor Cyan

# 1. 安裝 Python 依賴
Write-Host ""
Write-Host "[1/2] 安裝 Python 依賴..." -ForegroundColor Yellow
pip install -r requirements.txt -q
if ($LASTEXITCODE -ne 0) {
    Write-Host "  x 失敗，請手動執行：pip install -r requirements.txt" -ForegroundColor Red
    exit 1
}
Write-Host "  v openpyxl 已安裝" -ForegroundColor Green

# 2. 跑測試，確保這份 clone 本身沒問題
Write-Host ""
Write-Host "[1.5/2] 跑內建測試..." -ForegroundColor Yellow
python tests\run_tests.py
if ($LASTEXITCODE -ne 0) {
    Write-Host "  x 測試沒過，先別裝 skill，回報這個訊息給維護者" -ForegroundColor Red
    exit 1
}
Write-Host "  v 測試全綠" -ForegroundColor Green

# 3. 安裝 Claude Code skill
Write-Host ""
Write-Host "[2/2] 安裝 Claude Code skill ..." -ForegroundColor Yellow
$skillDir = Join-Path $env:USERPROFILE ".claude\skills\material-price-payment"
New-Item -ItemType Directory -Force -Path $skillDir | Out-Null

# 讀 repo 的 SKILL.md，替換 <PROJECT_PATH>，寫到 ~/.claude/skills/
$skillContent = Get-Content "SKILL.md" -Raw -Encoding UTF8
$skillContent = $skillContent -replace [regex]::Escape("<PROJECT_PATH>"), $root
[System.IO.File]::WriteAllText(
    (Join-Path $skillDir "SKILL.md"),
    $skillContent,
    (New-Object System.Text.UTF8Encoding $false)
)
Write-Host "  v 已安裝到 $skillDir\SKILL.md" -ForegroundColor Green

Write-Host ""
Write-Host "安裝完成！重啟 Claude Code，輸入「核對估驗計價」即可使用。" -ForegroundColor Green
Write-Host ""
Write-Host "若要 fork / PR：本腳本不會動到 repo 內任何檔案，可放心 git commit / push。" -ForegroundColor Cyan
Write-Host ""
