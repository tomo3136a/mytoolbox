@echo off
set pkg=AddinDev
pushd %~dp0
set OPT=-Sta -NoProfile -NoLogo -ExecutionPolicy RemoteSigned
powershell.exe %OPT% ./tools/new-xlam.ps1 .\%pkg%
powershell.exe %OPT% ./tools/add-customui.ps1 .\%pkg%
popd
