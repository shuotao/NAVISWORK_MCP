@echo off
chcp 65001 >nul
echo ============================================
echo   NavisworksMCP - 一鍵安裝腳本
echo ============================================
echo.

:: 檢查 Node.js
where node >nul 2>&1
if %errorlevel% neq 0 (
    echo [錯誤] 找不到 Node.js，請先安裝 Node.js LTS
    echo 下載: https://nodejs.org/
    pause
    exit /b 1
)
echo [OK] Node.js:
node --version

:: 安裝 MCP Server 依賴
echo.
echo [步驟 1/2] 安裝 MCP Server 依賴...
cd /d "%~dp0..\MCP-Server"
call npm install
if %errorlevel% neq 0 (
    echo [錯誤] npm install 失敗
    pause
    exit /b 1
)

:: 編譯 TypeScript
echo.
echo [步驟 2/2] 編譯 MCP Server...
call npm run build
if %errorlevel% neq 0 (
    echo [錯誤] TypeScript 編譯失敗
    pause
    exit /b 1
)

echo.
echo ============================================
echo   安裝完成！
echo ============================================
echo.
echo 使用方法:
echo   1. 在 Visual Studio 中開啟 NavisworksMCP.sln
echo   2. 設定 NavisworksPath 指向你的 Navisworks 安裝路徑
echo   3. 編譯 C# 專案 (DLL 會自動複製到 Plugins 資料夾)
echo   4. 開啟 Navisworks，點擊 "MCP 服務 (開/關)" 啟動服務
echo   5. AI 平台即可透過 MCP Server 控制 Navisworks
echo.
echo WebSocket 端口: 2233
echo.
pause
