# NavisworksMCP Add-in 安裝腳本
# 將編譯好的 DLL 複製到 Navisworks Plugins 資料夾

param(
    [string]$NavisworksVersion = "2025",
    [string]$NavisworksPath = ""
)

$ErrorActionPreference = "Stop"

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  NavisworksMCP Add-in 安裝" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

# 偵測 Navisworks 安裝路徑
if (-not $NavisworksPath) {
    $NavisworksPath = "C:\Program Files\Autodesk\Navisworks Manage $NavisworksVersion"
}

if (-not (Test-Path $NavisworksPath)) {
    Write-Host "[錯誤] 找不到 Navisworks: $NavisworksPath" -ForegroundColor Red
    Write-Host "請使用 -NavisworksPath 參數指定正確路徑" -ForegroundColor Yellow
    exit 1
}

Write-Host "[OK] Navisworks 路徑: $NavisworksPath" -ForegroundColor Green

# 建立 Plugins 目標資料夾
$pluginDir = Join-Path $NavisworksPath "Plugins\NavisworksMCP"
if (-not (Test-Path $pluginDir)) {
    New-Item -ItemType Directory -Path $pluginDir -Force | Out-Null
    Write-Host "[OK] 已建立 Plugin 資料夾: $pluginDir" -ForegroundColor Green
}

# 複製 DLL
$buildDir = Join-Path $PSScriptRoot "..\MCP\bin\Release"
if (-not (Test-Path $buildDir)) {
    $buildDir = Join-Path $PSScriptRoot "..\MCP\bin\Debug"
}

if (-not (Test-Path $buildDir)) {
    Write-Host "[錯誤] 找不到編譯輸出，請先在 Visual Studio 中編譯專案" -ForegroundColor Red
    exit 1
}

$files = @("NavisworksMCP.dll", "NavisworksMCP.pdb", "Newtonsoft.Json.dll")
foreach ($file in $files) {
    $src = Join-Path $buildDir $file
    if (Test-Path $src) {
        Copy-Item $src $pluginDir -Force
        Write-Host "[OK] 已複製: $file" -ForegroundColor Green
    }
}

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  安裝完成！" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "請重新啟動 Navisworks 以載入 Plugin" -ForegroundColor Yellow
