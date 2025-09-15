@echo off
pushd %~dp0

set ODIR=%CD%\package

if not exist %ODIR%\bin mkdir %ODIR%\bin
if not exist %ODIR%\lib mkdir %ODIR%\lib

if exist setr\build.cmd (
  call setr\build.cmd -pass %ODIR%\bin
  copy setr\setpath.cmd %ODIR%
  copy setr\lib\setpath.ps1 %ODIR%\lib
)

if exist indexed\build.cmd (
  call indexed\build.cmd -pass %ODIR%\bin
  copy indexed\lib\install.cmd %ODIR%
  copy indexed\lib\uninstall.cmd %ODIR%
  copy indexed\lib\install_task.cmd %ODIR%
  copy indexed\lib\install_task.ps1 %ODIR%
)

popd
pause
