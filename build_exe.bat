@echo off
chcp 65001
echo ========================================
echo   PT Report to CSV 執行檔打包工具
echo ========================================
echo.

echo [1/4] 檢查 Python 環境...
python --version >nul 2>&1
if errorlevel 1 (
    echo 錯誤：找不到 Python，請確認 Python 已安裝並加入 PATH
    pause
    exit /b 1
)
python --version

echo.
echo [2/4] 安裝 PyInstaller...
pip install pyinstaller
if errorlevel 1 (
    echo 錯誤：PyInstaller 安裝失敗
    pause
    exit /b 1
)

echo.
echo [3/4] 安裝專案依賴套件...
pip install pandas openpyxl
if errorlevel 1 (
    echo 錯誤：依賴套件安裝失敗
    pause
    exit /b 1
)

echo.
echo [4/4] 使用 PyInstaller 打包執行檔...
pyinstaller main.spec --clean
if errorlevel 1 (
    echo 錯誤：執行檔打包失敗
    pause
    exit /b 1
)

echo.
echo ========================================
echo        打包完成！
echo ========================================
echo 執行檔位置：dist\pt-report-to-csv.exe
echo.
echo 使用說明：
echo 1. 將 "Portfolio Performance Report.xlsx" 放在執行檔同一資料夾
echo 2. 將 "symbol.csv" 放在執行檔同一資料夾
echo 3. 雙擊執行 pt-report-to-csv.exe
echo 4. 程式會產生 "portfolio_trades.csv" 輸出檔
echo.
pause