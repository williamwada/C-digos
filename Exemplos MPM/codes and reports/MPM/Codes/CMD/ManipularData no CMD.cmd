@echo off
cls
echo Date format = %date%
echo dd = %date:~8,2%
echo mm = %date:~5,2%
echo yyyy = %date:~0,4%
echo.
echo Time format = %time%
echo hh = %time:~0,2%
echo mm = %time:~3,2%
echo ss = %time:~6,2%
echo.
echo FinalData = %date:~0,4%-%date:~5,2%-%date:~8,2%-%time:~0,2%-%time:~3,2%-%time:~6,2%
pause