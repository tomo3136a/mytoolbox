@echo off
set /p pkg=アドイン名？：
pushd %~dp0
if not exist %pkg% mkdir %pkg%
set OPT=-Sta -NoProfile -NoLogo -ExecutionPolicy RemoteSigned
powershell.exe %OPT% lib/set-ribbonbas.ps1 .\%pkg% 
powershell.exe %OPT% lib/new-xlam.ps1 .\%pkg%
powershell.exe %OPT% lib/add-customui.ps1 .\%pkg% -ctmenu
popd
pause
