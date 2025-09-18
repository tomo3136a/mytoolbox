@echo off
pushd %~dp0
echo uninstalling...
set p=c:\opt\bin
if exist lib\install_task.cmd (
    call lib\install_task.cmd -Clean -Pass
)
if exist %p%\indexed.exe (
%p%\indexed.exe -u1
del %p%\indexed.exe
)
popd
pause
