@echo off
pushd %~dp0
echo uninstalling...
if exist install_task.cmd (
  call install_task.cmd -Clean -Pass
)
if exist ..\bin\indexed.exe (
  ..\bin\indexed.exe -u1
  timeout /t 2 >nul
  del ..\bin\indexed.exe
)
popd
pause
