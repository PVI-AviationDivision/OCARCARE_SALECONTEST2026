@echo off
REM Tao du lieu va mo trang tren may (khong day len GitHub) de xem truoc
chcp 65001 >nul
set PYTHONUTF8=1
cd /d "%~dp0"
set PY=python
where python >nul 2>nul || set PY=py
%PY% tools\build_data.py && start "" "docs\index.html"
pause
