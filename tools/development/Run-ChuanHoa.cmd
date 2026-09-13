@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Run-ChuanHoa.ps1" -OpenAdmin
if errorlevel 1 pause
endlocal
