
@echo off
cls
setr -i %0 -u
pause
goto :eof

rem #選択
rem * -f AAA
rem * -f BBB=c:\work\test1.ps1
rem * -f CCC=c:\work\test.txt
rem * -f DDD=c:\work\test.txt テキスト *.txt ログ *.log
rem * -f EEE=c:\work\test.txt PDF *.pdf テキスト *.txt ログ *.log