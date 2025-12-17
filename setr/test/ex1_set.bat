@echo on
cls
setr -i %0 -u
pause
goto :eof

rem コメント行
rem　"#" 文字以降はコメント行を作成
rem #test comment       1
rem #test      comment #2
rem #     test comment  3   
rem #test4^
rem #test5
rem #
rem
rem 設定行
rem * AAA
rem * BBB=123^
rem * AAA
rem * CCC=123^
rem *AAA
rem * DDD="123 456 789"
rem * EEE=123 FFF=456
rem * GGG="    123    456   "
