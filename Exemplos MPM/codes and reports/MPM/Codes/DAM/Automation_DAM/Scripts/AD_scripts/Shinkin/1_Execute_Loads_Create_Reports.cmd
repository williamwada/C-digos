set _LogName=%date:~0,4%%date:~5,2%%date:~8,2%_%time:~0,2%%time:~3,2%
set LogName=%_LogName: =%

sqlcmd -S ADCF-PCKT\ADCFSQL -d shinkin -U adcf_user -P adcf_password -i "S:\Daily opreation\auto\main.sql" -o logs\LOG_Shinkin_%LogName%.txt
pause
