@echo off
chcp 65001 >nul
REM Tao lich chay update.bat luc 17:30 hang ngay (chay 1 lan la du)
schtasks /create /f /sc daily /st 17:30 /tn "OCarCare Dashboard Update" /tr "\"%~dp0update.bat\" auto"
echo Da tao lich: 17:30 hang ngay. Xem/sua trong Task Scheduler.
pause
