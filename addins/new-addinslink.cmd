@echo off
pushd %~dp0
set OPT=-Sta -NonInteractive -NoProfile -NoLogo -ExecutionPolicy RemoteSigned
powershell.exe %OPT% ./tools/%~n0.ps1 %*
popd
