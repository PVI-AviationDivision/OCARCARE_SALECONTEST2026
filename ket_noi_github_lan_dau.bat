@echo off
chcp 65001 >nul
set PYTHONUTF8=1
cd /d "%~dp0"
echo ==========================================================
echo   KET NOI THU MUC NAY VOI REPO GITHUB (chi chay 1 lan)
echo ==========================================================
set REPO=https://github.com/PVI-AviationDivision/OCARCARE_SALECONTEST2026.git

where git >nul 2>nul || (echo [LOI] May chua cai Git: https://git-scm.com/download/win & pause & exit /b 1)
set PY=python
where python >nul 2>nul || set PY=py
%PY% --version || (echo [LOI] May chua cai Python: https://www.python.org/downloads/ & pause & exit /b 1)

echo.
echo [1/4] Cai thu vien doc Excel...
%PY% -m pip install --quiet openpyxl

echo [2/4] Tao du lieu tu file bao cao trong input\ ...
%PY% tools\build_data.py || (pause & exit /b 1)

echo [3/4] Ket noi repo GitHub...
if not exist ".git" git init -q
git config user.email >nul 2>nul || git config user.email "ocarcare-dashboard@users.noreply.github.com"
git config user.name >nul 2>nul || git config user.name "OCarCare Dashboard"
git remote remove origin >nul 2>nul
git remote add origin %REPO%
git fetch -q origin || (echo [LOI] Khong ket noi duoc GitHub & pause & exit /b 1)
git reset -q origin/main
git branch -M main

echo [4/4] Day len GitHub (lan dau se hien cua so dang nhap GitHub)...
git add -A
git status --short | findstr /i ".xls" && (echo [DUNG] Phat hien file Excel - khong day len! & pause & exit /b 1)
git commit -q -m "Chuyen sang cau truc docs/ + cap nhat tu dong" || (echo [LOI] Khong tao duoc commit & pause & exit /b 1)
git push -u origin main || (echo [LOI] Day len that bai - kiem tra quyen ghi vao repo & pause & exit /b 1)

echo.
echo [OK] Xong. Buoc cuoi: vao GitHub - Settings - Pages - chon Branch "main", thu muc "/docs" - Save.
echo Link: https://pvi-aviationdivision.github.io/OCARCARE_SALECONTEST2026/
pause
