@echo off
pushd %~dp0
set SRC=lib\addins
set DST=%APPDATA%\Microsoft\Addins
copy %SRC%\my*.xlam %DST%
if "%1"=="dev" (
    copy %SRC%\adindev*.xlam %DST%
)
popd
pause
