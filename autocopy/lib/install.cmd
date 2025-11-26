@echo off
pushd %~dp0
echo installing...
if exist ..\bin\autocopy.exe (
  ..\bin\autocopy.exe -u
)
popd
pause
