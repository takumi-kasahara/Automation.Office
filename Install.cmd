@echo off
setlocal enableextensions disabledelayedexpansion
cd /D "%~dp0"

:begin

:process
powershell -NoLogo -NoProfile -NonInteractive -ExecutionPolicy RemoteSigned -File "Install.Modules.ps1"
powershell -NoLogo -NoProfile -NonInteractive -ExecutionPolicy RemoteSigned -File "Install.Scripts.ps1"

:end
exit /B %errorlevel%
