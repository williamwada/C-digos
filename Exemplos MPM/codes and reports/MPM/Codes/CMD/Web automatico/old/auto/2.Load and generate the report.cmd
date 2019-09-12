set _LogName=%date:~0,4%%date:~5,2%%date:~8,2%_%time:~0,2%%time:~3,2%
set LogName=%_LogName: =%

sqlcmd -S ADCF-PCKT\ADCFSQL -d EPOS -U adcf_user -P adcf_password -i "Y:\Daily opreation\WebEpos\auto\main.sql" -o logs\LOG_web_epos_%LogName%.txt
pause