@echo off
set pkg=mydesigner
pushd %~dp0
set OPT=-Sta -NoProfile -NoLogo -ExecutionPolicy RemoteSigned
powershell.exe %OPT% lib/new-xlam.ps1 .\%pkg%
powershell.exe %OPT% lib/add-customui.ps1 .\%pkg%
popd
