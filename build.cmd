@echo off
setlocal enabledelayedexpansion
pushd %~dp0

set PKG=mtb
set ODIR=%~dp0\package
set PDIR=%ODIR%\%PKG%
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

if exist files\build.cmd (
  call files\build.cmd -pass %PDIR%\bin
)

copy lib\install.cmd %PDIR%
copy lib\uninstall.cmd %PDIR%

if not exist %SZIP% goto eof
if not exist %PDIR% goto eof
if not exist %ODIR% mkdir %ODIR%

%SZIP% a -tzip %ODIR%/mtb %PDIR%\*
%SZIP% a -sfx7z.sfx %ODIR%/mtb %PDIR%

set PTN1=abcdefghijklmnopqrstuvwxyz
set PTN2=0123456789%PTN1%

set /a N=1%DATE:~5,2%-100
set CMD=echo %%PTN2:~%N%,1%%
for /f "usebackq tokens=*" %%a in (`!CMD!`) do (set DT1=%%a)

set /a N=1%DATE:~8,2%-100
set CMD=echo %%PTN2:~%N%,1%%
for /f "usebackq tokens=*" %%a in (`!CMD!`) do (set DT2=%%a)

set /a N=%TIME:~0,2%
set CMD=echo %%PTN1:~%N%,1%%
for /f "usebackq tokens=*" %%a in (`!CMD!`) do (set TM=%%a)

set REV=%DATE:~0,4%%DT1%%DT2%%TM%

move %ODIR%\mtb.zip %ODIR%\mtb_%REV%.zip
move %ODIR%\mtb.exe %ODIR%\mtb_%REV%.exe

:eof
popd
pause
