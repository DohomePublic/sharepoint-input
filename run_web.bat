@echo off
REM เปิดหน้าเว็บเพิ่มข้อมูล  ใช้:  run_web.bat   หรือ  run_web.bat --dry-run
cd /d "%~dp0"
pip install -r requirements.txt --quiet
python webapp\server.py --port 8000 %*
pause
