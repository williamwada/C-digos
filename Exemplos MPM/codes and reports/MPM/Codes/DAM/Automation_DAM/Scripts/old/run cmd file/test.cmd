echo off
echo test
copy C:\test.txt C:\test\crabon
echo %ERRORLEVEL%
IF %ERRORLEVEL% GEQ 1 EXIT /B 2
pause
