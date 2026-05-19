@echo off
cd /d "%~dp0"
echo 正在啟動商品上架結構表產生器...
echo.
echo 請用瀏覽器開啟: http://localhost:5000
echo 關閉此視窗將停止服務
echo.
python app.py
pause