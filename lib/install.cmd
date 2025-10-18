@echo off
pushd %~dp0
echo installing...
set p=c:\opt\mtb

if not exist %p%\bin mkdir %p%\bin
if not exist %p%\lib mkdir %p%\lib

xcopy bin\*.* %p%\bin /s /e /q
xcopy lib\*.* %p%\lib /s /e /q

call setpath.cmd %p%\bin -pass

popd
pause
