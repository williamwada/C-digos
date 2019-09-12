
::-----------------------------------------------------------------------
@ECHO ON
::-----------------------------------------------------------------------
set YYYYMMDD=%date:~6,4%%date:~0,2%%date:~3,2%

::EXECUTE 1.GET FILES
CALL "D:\AD_EPOS_PRD\DAM_Automations\Monday\1.DAM_Get_Ftp.cmd" :: 1>> log\log_%YYYYMMDD%.log 2>> log\error_%YYYYMMDD%.log

::EXECUTE 2.LOAD AND CREATE REPORTS
cscript.exe "D:\AD_EPOS_PRD\DAM_Automations\Monday\2_Copy_load_Create_Report.vbs" :: 1>> log\log_%YYYYMMDD%.log 2>> log\error_%YYYYMMDD%.log

::EXECUTE 3.ZIP FILES
cscript.exe "D:\AD_EPOS_PRD\DAM_Automations\Monday\3.zip_Files.vbs" :: 1>> log\log_%YYYYMMDD%.log 2>> log\error_%YYYYMMDD%.log

::EXECUTE 4.Send Emails
cscript.exe "D:\AD_EPOS_PRD\DAM_Automations\Monday\4.Send_Emails.wsf" :: 1>> log\log_%YYYYMMDD%.log 2>> log\error_%YYYYMMDD%.log

