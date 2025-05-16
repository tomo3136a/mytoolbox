@echo off
set /p pkg=アドイン名？：
pushd %~dp0
if not exist %pkg% mkdir %pkg%
set OPT=-Sta -NoProfile -NoLogo -ExecutionPolicy RemoteSigned
powershell.exe %OPT% ./tools/add-ribbonbas.ps1 .\%pkg% 
powershell.exe %OPT% ./tools/new-xlam.ps1 .\%pkg%
powershell.exe %OPT% ./tools/add-customui.ps1 .\%pkg%
popd
pause
