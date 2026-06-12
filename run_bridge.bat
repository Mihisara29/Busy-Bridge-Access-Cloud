@echo off
echo ========================================
echo   BUSY 21 Bridge v2.0 Launcher
echo   Using 32-bit PowerShell (SysWOW64)
echo ========================================
echo.
C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe ^
    -ExecutionPolicy Bypass ^
    -File "%~dp0busy_api.ps1" ^
    -Port 8081
pause