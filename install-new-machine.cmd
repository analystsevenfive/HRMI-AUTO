@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\install-new-machine.ps1"
set "INSTALL_EXIT=%ERRORLEVEL%"
echo.
if not "%INSTALL_EXIT%"=="0" (
  echo Installation failed. Review the message above.
) else (
  echo Installation completed.
)
pause
exit /b %INSTALL_EXIT%
