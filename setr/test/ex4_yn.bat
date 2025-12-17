@echo off
cls
setr -i %0 -u
pause
goto :eof

rem #yes/no判定
rem * -y AAA
rem * -y BBB 123
rem * -y CCC 456 789
rem * -y DDD=0
rem * -y EEE=1
rem * -y FFF=11 11 22
rem * -y GGG=22 11 22
rem * -y HHH=33 11 22
rem * -y III=111 111 222 -m テスト1
rem * -y JJJ=222 111 222 -m "テスト2  テスト2  テスト2"
rem * -m "テスト3  テスト3  テスト3"
rem * -y KKK 111 222
rem * -y LLL 111 222
