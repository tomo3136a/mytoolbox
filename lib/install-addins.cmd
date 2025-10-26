@echo off
pushd %~dp0
set SRC=lib\addins
set DST=%APPDATA%\Microsoft\Addins
copy %SRC%\my*.xlam %DST%
popd
pause
