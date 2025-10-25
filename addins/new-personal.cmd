@echo off
pushd %~dp0
set OPT=-Sta -NoProfile -NoLogo -ExecutionPolicy RemoteSigned
powershell.exe %OPT% lib/%~n0.ps1 %*
popd
pause
