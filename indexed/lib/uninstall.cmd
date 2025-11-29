@echo off
pushd %~dp0
set APP=indexed

echo uninstalling...
if exist install_task_%APP%.cmd (
  call install_task_%APP%.cmd -Clean -Pass
)
if exist ..\bin\%APP%.exe (
  ..\bin\%APP%.exe -u1
  timeout /t 2 >nul
  del ..\bin\%APP%.exe
)
popd
pause
