@echo off
pushd %~dp0

set ODIR=%CD%\package

if not exist %ODIR%\bin mkdir %ODIR%\bin
if not exist %ODIR%\lib mkdir %ODIR%\lib

if exist setr\build.cmd (
  call setr\build.cmd -pass %ODIR%\bin
  del %ODIR%\bin\install.cmd
  copy setr\setpath.cmd %ODIR%
  copy setr\lib\setpath.ps1 %ODIR%\lib
)

if exist indexed\build.cmd (
  call indexed\build.cmd -pass %ODIR%\bin
  move %ODIR%\bin\install.cmd %ODIR%
  move %ODIR%\bin\uninstall.cmd %ODIR%
  move %ODIR%\bin\install_task.* %ODIR%\lib
)

copy lib\install.cmd %ODIR%
copy lib\uninstall.cmd %ODIR%

popd
pause
