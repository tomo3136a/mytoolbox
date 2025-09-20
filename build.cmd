@echo off
pushd %~dp0

set PDIR=%~dp0\package
set ODIR=%~dp0\output
set SZIP="c:\Program Files\7-Zip\7z.exe"

if not exist %PDIR%\bin mkdir %PDIR%\bin
if not exist %PDIR%\lib mkdir %PDIR%\lib

if exist setr\build.cmd (
  call setr\build.cmd -pass %PDIR%\bin
  del %PDIR%\bin\install.cmd
  copy setr\setpath.cmd %PDIR%
  copy setr\lib\setpath.ps1 %PDIR%\lib
)

if exist indexed\build.cmd (
  call indexed\build.cmd -pass %PDIR%\bin
  move %PDIR%\bin\install.cmd %PDIR%
  move %PDIR%\bin\uninstall.cmd %PDIR%
  move %PDIR%\bin\install_task.* %PDIR%\lib
)

copy lib\install.cmd %PDIR%
copy lib\uninstall.cmd %PDIR%

if not exist %SZIP% goto eof
if not exist %PDIR% goto eof
if not exist %ODIR% mkdir %ODIR%

pushd %PDIR%
%SZIP% a -tzip %ODIR%/mtb *
%SZIP% a -sfx7z.sfx %ODIR%/mtb *
popd

:eof
popd
pause
