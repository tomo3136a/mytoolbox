@echo off
pushd %~dp0
echo uninstalling...
set p=c:\opt\mtb\bin
if exist lib\install_task.cmd (
  call lib\install_task.cmd -Clean -Pass
)
if exist lib\indexed.exe (
  lib\indexed.exe -u1
  timeout /t 2 >nul
)

if exist %p%\mkfolder.exe ( del %p%\mkfolder.exe )
if exist %p%\indexed.exe ( del %p%\indexed.exe )
if exist %p%\files.exe ( del %p%\files.exe )
if exist %p%\setr.exe ( del %p%\setr.exe )
if exist %p%\..\lib\install_indexed.cmd ( del %p%\..\lib\install_indexed.cmd )
if exist %p%\..\lib\install_task.cmd ( del %p%\..\lib\install_task.cmd )
if exist %p%\..\lib\install_task.ps1 ( del %p%\..\lib\install_task.ps1 )
if exist %p%\..\lib\uninstall_indexed.cmd ( del %p%\..\lib\uninstall_indexed.cmd )
if exist %p%\..\lib\setpath.ps1 ( del %p%\..\lib\setpath.ps1 )

popd
pause
