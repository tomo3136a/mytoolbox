@echo off
pushd %~dp0
echo installing...
set p=c:\opt

if not exist %p%\bin mkdir %p%\bin
if not exist %p%\lib mkdir %p%\lib

copy bin\*.* %p%\bin
copy lib\*.* %p%\lib

popd
pause
