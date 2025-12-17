@echo off
cls
setr -i %0 -u
pause
goto :eof

rem #メッセージ表示
rem * -b AAA
rem * -b BBB 123
rem * -b CCC 456 789
rem * -b DDD=0
rem * -b EEE=1
rem * -b FFF=11 11 22
rem * -b GGG=22 11 22
rem * -b HHH=33 11 22

rem * -b III=111 111 222 -m テスト1
rem * -b JJJ=222 111 222 -m "テスト2  テスト2  テスト2"
rem * -m "テスト3  テスト3  テスト3"
rem * -b KKK 111 222
rem * -b LLL 111 222
