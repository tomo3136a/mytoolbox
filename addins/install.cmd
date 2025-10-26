@echo off
pushd %~dp0
set SRC=addins
set DST=%APPDATA%\Microsoft\Addins
copy %SRC%\*.* %DST%
popd
