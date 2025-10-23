@echo off
setlocal
set "PS1=%~dp0WahWahButton.ps1"
if not exist "%PS1%" (
  echo PowerShell script not found: %PS1%
  exit /b 1
)
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%PS1%"
endlocal
