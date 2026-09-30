@echo off
cd /d "%~dp0"
setlocal

:begin

:process
powershell -NoLogo -NoProfile -ExecutionPolicy RemoteSigned -File "Install.Modules.ps1"
powershell -NoLogo -NoProfile -ExecutionPolicy RemoteSigned -File "Install.Scripts.ps1"

:end
exit /b %errorlevel%
