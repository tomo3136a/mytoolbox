@echo off
setlocal enabledelayedexpansion
pushd %~dp0

set PKG=mtb
set LST=setr indexed files mkfolder

set ODIR=%~dp0\package
set PDIR=%ODIR%\%PKG%
set SZIP="c:\Program Files\7-Zip\7z.exe"

if not exist %PDIR%\bin mkdir %PDIR%\bin
if not exist %PDIR%\lib mkdir %PDIR%\lib

if exist setpath (
  xcopy setpath %PDIR% /s /e /q
)

for %%f in (%LST%) do (
  if exist %%f\build.cmd (
    call %%f\build.cmd -pass %PDIR%\bin
  )
  if exist %PDIR%\lib\install.cmd (
    move %PDIR%\lib\install.cmd %PDIR%\lib\install_%%f.cmd
  )
  if exist %PDIR%\lib\uninstall.cmd (
    move %PDIR%\lib\uninstall.cmd %PDIR%\lib\uninstall_%%f.cmd
  )
)

copy lib\install.cmd %PDIR%
copy lib\uninstall.cmd %PDIR%

if exist custom (
  xcopy custom %PDIR%\lib /s /e /q
)

if not exist %SZIP% goto eof
if not exist %PDIR% goto eof
if not exist %ODIR% mkdir %ODIR%

%SZIP% a -tzip %ODIR%/mtb %PDIR%\* >nul
%SZIP% a -sfx7z.sfx %ODIR%/mtb %PDIR% >nul
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
