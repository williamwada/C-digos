::set date time to log file name
set _LogName=%date:~0,4%%date:~5,2%%date:~8,2%_%time:~0,2%%time:~3,2%
set LogName=%_LogName: =%

::Execute the SQL scripts to load and create the report
sqlcmd -S ADCF-PCKT\ADCFSQL -d 7Card -U adcf_user -P adcf_password -i "Z:\Daily opreation\DAM\auto\main.sql" -o logs\LOG_web_7Card_%LogName%.txt
