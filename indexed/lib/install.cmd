@echo off
pushd %~dp0
echo installing...
if exist ..\bin\indexed.exe (
  ..\bin\indexed.exe -u
)
popd
pause
