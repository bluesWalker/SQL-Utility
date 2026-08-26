@echo off
setlocal
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0SqlUtility.ps1" -VerifyCatalog
set "sqlUtilityExitCode=%ERRORLEVEL%"
endlocal & exit /b %sqlUtilityExitCode%
