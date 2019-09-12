
'
'MOVE THE FILES TO DESTINATION FOLDERS
'CREATE BY WW  2017/03/09
'
'set yesterday date
d = date() - 1
dateFinal= Cstr(year(d) * 10000 + month(d) * 100 + day(d))

'
'LOAD YESFILE, F13, LOAD FILES
'CREATE BY WILLIAM 20170309
'

' Globals
Dim gsLibDir : gsLibDir = ".\"
Dim goFS     : Set goFS = CreateObject("Scripting.FileSystemObject")

' LibraryInclude
ExecuteGlobal goFS.OpenTextFile(goFS.BuildPath(gsLibDir, "Class_load.vbs")).ReadAll()



'------EPOS-------------------------------------------------------------------

'set paths---
'F13
EP_F13_SFile   = "F:\EPOS_AD\f13\*.xls"
Const EP_F13_SFolder = "F:\EPOS_AD\f13\"
Const EP_F13_DestFolder = "Y:\DATA_AREA\STAGING_AREA\WEB_F13\"
const EP_F13_DestUSBFolder = "I:\DailyReports\"

Dim o : Set o = New Class_Load
o.SourceFolder = EP_F13_SFolder
o.DestFolder = EP_F13_DestFolder
o.SourceFile =EP_F13_SFile
o.Load()

'LOG can
EP_LOG_CAN_SFile   = "F:\EPOS_AD\logfile\adcan\*.csv"
Const EP_Log_SFolder = "F:\EPOS_AD\logfile\adcan\"
Const EP_Log_CAN_DestFolder = "Y:\DATA_AREA\STAGING_AREA\WEB_log\ADCAN\"

'LOG mgi
EP_LOG_MGI_SFile   = "F:\EPOS_AD\logfile\mgi\*.csv"
Const EP_Log_MGI_SFolder = "F:\EPOS_AD\logfile\mgi\"
Const EP_Log_MGI_DestFolder = "Y:\DATA_AREA\STAGING_AREA\WEB_log\ADMGI\"


'YES Files can
EP_YES_CAN_SFile = "F:\EPOS_AD\yesfile\adcan\*.txt"
Const Yes_SourceFolder = "F:\EPOS_AD\yesfile\adcan\"
Const Yes_DestFolder = "Y:\DATA_AREA\STAGING_AREA\WEB_YES_File\"

'YES Files mgi
EP_YES_MGI_SFile = "F:\EPOS_AD\yesfile\mgi\*.txt"
Const Yes_SourceFolder = "F:\EPOS_AD\yesfile\mgi\"
Const Yes_DestFolder = "Y:\DATA_AREA\STAGING_AREA\WEB_YES_File\"

'End Set Paths---


'F13 MOVE the FILES---
'set objects
Dim Class_load
 




















'//End























'------7CARD-----
'F13
seven_F13_SFile   = "F:\7CARD_AD\f13\*.xls"
Const seven_F13_SFolder = "F:\7CARD_AD\f13\"
Const seven_F13_DestFolder = "Y:\DATA_AREA\STAGING_AREA\WEB_F13\"
const seven_F13_DestUSBFolder = "I:\DailyRsevenorts\"

'LOG can
seven_LOG_CAN_SFile   = "F:\7CARD_AD\logfile\adcan\*.csv"
Const seven_Log_SFolder = "F:\7CARD_AD\logfile\adcan\"
Const seven_Log_CAN_DestFolder = "Y:\DATA_AREA\STAGING_AREA\WEB_log\ADCAN\"

'LOG mgi
seven_LOG_MGI_SFile   = "F:\7CARD_AD\logfile\mgi\*.csv"
Const seven_Log_MGI_SFolder = "F:\7CARD_AD\logfile\mgi\"
Const seven_Log_MGI_DestFolder = "Y:\DATA_AREA\STAGING_AREA\WEB_log\ADMGI\"


'YES Files can
seven_YES_CAN_SFile = "F:\7CARD_AD\yesfile\adcan\*.txt"
Const Yes_SourceFolder = "F:\7CARD_AD\yesfile\adcan\"
Const Yes_DestFolder = "Y:\DATA_AREA\STAGING_AREA\WEB_YES_File\"

'YES Files mgi
seven_YES_MGI_SFile = "F:\7CARD_AD\yesfile\mgi\*.txt"
Const Yes_SourceFolder = "F:\7CARD_AD\yesfile\mgi\"
Const Yes_DestFolder = "Y:\DATA_AREA\STAGING_AREA\WEB_YES_File\"

'//End









    'Check to see if the file already exists in the original folder
'    If fso.FileExists(SourceFolder) Then
        
'		fso.moveFile SourceFolder, DestFolder
'		msgbox"File moved!",64,"Success"
'    Else
'        msgbox"File not found!",vbError,"Error"
'    End If

'Clean up! 
Set fso = Nothing
Set objFolder = Nothing
Set objFile = Nothing

'COPY  F:\EPOS_AD\f13\*.xls I:\DailyReports
'move F:\EPOS_AD\f13\*.xls Y:\DATA_AREA\STAGING_AREA\WEB_F13\
'move F:\EPOS_AD\logfile\adcan\*.csv Y:\DATA_AREA\STAGING_AREA\WEB_log\ADCAN\
'move F:\EPOS_AD\logfile\mgi\*.csv Y:\DATA_AREA\STAGING_AREA\WEB_log\ADMGI\
'move F:\EPOS_AD\yesfile\adcan\*.txt Y:\DATA_AREA\STAGING_AREA\WEB_YES_File\
'move F:\EPOS_AD\yesfile\mgi\*.txt Y:\DATA_AREA\STAGING_AREA\WEB_YES_File\