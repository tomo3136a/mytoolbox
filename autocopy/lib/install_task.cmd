@echo off
pushd %~dp0
set OPT=-Sta -NonInteractive -NoProfile -NoLogo -ExecutionPolicy RemoteSigned
set PS1=./%~n0.ps1
if exist ./lib/%~n0.ps1 set PS1=./lib/%~n0.ps1
powershell.exe %OPT% %PS1% %*
popd
