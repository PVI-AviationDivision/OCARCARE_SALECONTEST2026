@echo off
chcp 65001 >nul
set PYTHONUTF8=1
cd /d "%~dp0"
echo ==============================================
echo   OCARCARE Grand Prix - Cap nhat dashboard
echo ==============================================

set PY=python
where python >nul 2>nul || set PY=py

%PY% tools\build_data.py
if errorlevel 1 (
  echo.
  echo [LOI] Khong tao duoc du lieu. Kiem tra file trong thu muc input\ va mapping\
  if "%1"=="" pause
  exit /b 1
)

git add docs
git diff --cached --quiet
if not errorlevel 1 (
  echo Khong co thay doi so voi lan cap nhat truoc.
  goto end
)
for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyy-MM-dd_HHmm"') do set TS=%%i
git commit -m "Cap nhat du lieu %TS%"
git push
if errorlevel 1 (
  echo [LOI] Day len GitHub that bai. Kiem tra mang / dang nhap GitHub.
  if "%1"=="" pause
  exit /b 1
)
echo.
echo [OK] Da day len GitHub. Trang web se cap nhat sau khoang 1 phut.

:end
if "%1"=="" pause
