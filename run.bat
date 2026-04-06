@echo off
cd /d "%~dp0"
echo StockOps 시작 중...
echo 브라우저에서 http://localhost:5000 으로 접속하세요
start "" "http://localhost:5000"
python app.py
pause
