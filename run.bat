@echo off
title BUSY Permissions Discovery
echo Starting...

C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe ^
    -ExecutionPolicy Bypass ^
    -NoProfile ^
    -NoExit ^
    -Command "& { . 'C:\busy-bridge - Local\permissions.ps1' }"

echo.
echo Script finished.
pause