@echo off
chcp 65001 >nul
echo 啟動公文序號改名工具...
python "%~dp0pdf_renamer.py"
if errorlevel 1 (
    echo.
    echo 發生錯誤，請確認已安裝 Python 及相關套件
    echo 請參考「安裝與使用說明.txt」
    pause
)
