@echo off
setlocal
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -Command "$memberDefinition = '[System.Runtime.InteropServices.DllImport({0}kernel32.dll{0})] public static extern System.IntPtr GetConsoleWindow(); [System.Runtime.InteropServices.DllImport({0}user32.dll{0})] public static extern bool ShowWindow(System.IntPtr windowHandle, int command);' -f [char] 34; Add-Type -MemberDefinition $memberDefinition -Name NativeMethods -Namespace SqlUtilityLauncher; [void] [SqlUtilityLauncher.NativeMethods]::ShowWindow([SqlUtilityLauncher.NativeMethods]::GetConsoleWindow(), 0); & powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File \"%~dp0SqlUtility.ps1\" -VerifyCatalog; exit $LASTEXITCODE"
set "sqlUtilityExitCode=%ERRORLEVEL%"
endlocal & exit /b %sqlUtilityExitCode%
