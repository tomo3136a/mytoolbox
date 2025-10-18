@echo off
setlocal enabledelayedexpansion
pushd %~dp0

set PKG=mtb
set LST=setr setpath indexed files mkfolder

set ODIR=%~dp0\package
set PDIR=%ODIR%\%PKG%
set SZIP="c:\Program Files\7-Zip\7z.exe"

if not exist %PDIR%\bin mkdir %PDIR%\bin
if not exist %PDIR%\lib mkdir %PDIR%\lib

for %%f in (%LST%) do (
  if exist %%f\build.cmd (
    call %%f\build.cmd -pass %PDIR%\bin
  )
  if exist %PDIR%\bin\install.cmd (
    move %PDIR%\bin\install.cmd %PDIR%\lib\install_%%f.cmd
  )
  if exist %PDIR%\bin\uninstall.cmd (
    move %PDIR%\bin\uninstall.cmd %PDIR%\lib\uninstall_%%f.cmd
  )
)


rem if exist setr\build.cmd (
rem   call setr\build.cmd -pass %PDIR%\bin
rem   move %PDIR%\bin\install.cmd %PDIR%\lib\install_setr.cmd
rem   copy setr\setpath.cmd %PDIR%
rem   copy setr\lib\setpath.ps1 %PDIR%\lib
rem )

rem if exist indexed\build.cmd (
rem   call indexed\build.cmd -pass %PDIR%\bin
rem   move %PDIR%\bin\install.cmd %PDIR%\lib\install_indexed_menu.cmd
rem   move %PDIR%\bin\uninstall.cmd %PDIR%\lib\uninstall_indexed.cmd
rem   move %PDIR%\bin\install_task.* %PDIR%\lib
rem )

rem if exist files\build.cmd (
rem   call files\build.cmd -pass %PDIR%\bin
rem   move %PDIR%\bin\install.cmd %PDIR%\lib\install_files.cmd
rem )

rem if exist mkfolder\build.cmd (
rem   call mkfolder\build.cmd -pass %PDIR%\bin
rem )

if exist custom (
  xcopy custom %PDIR%\lib /s /e /q
)

copy lib\install.cmd %PDIR%
copy lib\uninstall.cmd %PDIR%

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
