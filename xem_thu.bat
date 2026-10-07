@echo off
REM Tao du lieu va mo trang tren may (khong day len GitHub) de xem truoc
chcp 65001 >nul
set PYTHONUTF8=1
cd /d "%~dp0"
set PY=python
where python >nul 2>nul || set PY=py
%PY% tools\build_data.py && start "" "docs\index.html"
if exist tin_nhan_zalo.txt (
  powershell -NoProfile -Command "Get-Content -Raw -Encoding UTF8 'tin_nhan_zalo.txt' | Set-Clipboard"
  echo.
  echo [OK] Tin nhan Zalo da duoc chep san - mo Zalo va nhan Ctrl+V de dan
)
pause
