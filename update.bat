@echo off
chcp 65001 >nul
set PYTHONUTF8=1
cd /d "%~dp0"
echo ==============================================
echo   OCARCARE Grand Prix - Cap nhat dashboard
echo ==============================================

set PY=python
where python >nul 2>nul || set PY=py

echo Dong bo du lieu moi nhat tu GitHub...
git pull --rebase --autostash -q origin main
if errorlevel 1 echo [CANH BAO] Chua dong bo duoc voi GitHub - kiem tra mang, van tiep tuc cap nhat
echo.

%PY% tools\build_data.py
if errorlevel 1 (
  echo.
  echo [LOI] Khong tao duoc du lieu. Kiem tra file trong thu muc input\ va mapping\
  if "%1"=="" pause
  exit /b 1
)

if exist tin_nhan_zalo.txt (
  powershell -NoProfile -Command "Get-Content -Raw -Encoding UTF8 'tin_nhan_zalo.txt' | Set-Clipboard"
  echo.
  echo [OK] Tin nhan Zalo da duoc chep san - mo Zalo va nhan Ctrl+V de dan
)

git add docs tools update.bat xem_thu.bat
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
