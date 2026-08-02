@echo off
setlocal
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0SqlConnectionPoc.ps1"
set "pocExitCode=%ERRORLEVEL%"
endlocal & exit /b %pocExitCode%
