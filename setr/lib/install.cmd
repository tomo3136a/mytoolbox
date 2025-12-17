@echo off
pushd %~dp0
set INST=c:\opt\bin
set APP=setr

echo installing...
if exist ..\bin\%APP%.exe (
  if not exist %INST% mkdir %INST% 
  copy ..\bin\%APP%.exe %INST%
)

popd
pause
