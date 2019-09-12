'#############################################################################
'Description: MOVE THE FILES TO DESTINATION FOLDERS YESFILE, F13, LOAD FILES
'CREATE BY:   WW  2017/03/09
'#############################################################################


' Start the instance of class 
Dim o : Set o = New moveFiles
Dim o1 : Set o1 = New copyFiles


'Move the files to destination folders: Yesfile, F13, Load files------------------------

'------EPOS
'F13
EP_F13_SFile   = "F:\EPOS_AD\f13\*.xls"
Const EP_F13_SFolder = "F:\EPOS_AD\f13\"
Const EP_F13_DestFolder = "Y:\DATA_AREA\STAGING_AREA\WEB_F13\"
const EP_F13_DestUSBFolder = "I:\DailyReports\"

'Call the Class Copy File
o1.CSourceFolder = EP_F13_SFolder
o1.CDestFolder = EP_F13_DestUSBFolder
o1.CSourceFile =EP_F13_SFile
o1.Load()



'Call the CLASS MOVE File
o.SourceFolder = EP_F13_SFolder
o.DestFolder = EP_F13_DestFolder
o.SourceFile =EP_F13_SFile
o.Load()


'LOG can
EP_LOG_CAN_SFile   = "F:\EPOS_AD\logfile\adcan\*.csv"
Const EP_Log_SFolder = "F:\EPOS_AD\logfile\adcan\"
Const EP_Log_CAN_DestFolder = "Y:\DATA_AREA\STAGING_AREA\WEB_log\ADCAN\"

'Call the CLASS MOVE File
o.SourceFolder = EP_Log_SFolder
o.DestFolder = EP_Log_CAN_DestFolder
o.SourceFile =EP_LOG_CAN_SFile
o.Load()


'LOG mgi
EP_LOG_MGI_SFile   = "F:\EPOS_AD\logfile\mgi\*.csv"
Const EP_Log_MGI_SFolder = "F:\EPOS_AD\logfile\mgi\"
Const EP_Log_MGI_DestFolder = "Y:\DATA_AREA\STAGING_AREA\WEB_log\ADMGI\"

'Call the CLASS MOVE File
o.SourceFolder = EP_Log_MGI_SFolder
o.DestFolder = EP_Log_MGI_DestFolder
o.SourceFile =EP_LOG_MGI_SFile
o.Load()

'YES Files can
EP_YES_CAN_SFile = "F:\EPOS_AD\yesfile\adcan\*.txt"
Const Yes_can_SourceFolder = "F:\EPOS_AD\yesfile\adcan\"
Const Yes_can_DestFolder = "Y:\DATA_AREA\STAGING_AREA\WEB_YES_File\"

'Call the CLASS MOVE File
o.SourceFolder = Yes_can_SourceFolder
o.DestFolder = Yes_can_DestFolder
o.SourceFile =EP_YES_CAN_SFile
o.Load()

'YES Files mgi
EP_YES_MGI_SFile = "F:\EPOS_AD\yesfile\mgi\*.txt"
Const Yes_mgi_SourceFolder = "F:\EPOS_AD\yesfile\mgi\"
Const Yes_mgi_DestFolder = "Y:\DATA_AREA\STAGING_AREA\WEB_YES_File\"

'Call the CLASS MOVE File
o.SourceFolder = Yes_mgi_SourceFolder
o.DestFolder = Yes_mgi_DestFolder
o.SourceFile =EP_YES_MGI_SFile
o.Load()

'------7CARD

'F13
seven_F13_SFile   = "F:\7CARD_AD\f13\*.xls"
Const seven_F13_SFolder = "F:\7CARD_AD\f13\"
Const seven_F13_DestFolder = "Z:\DATA_AREA\STAGING_AREA\WEB_F13\"
const seven_F13_DestUSBFolder = "I:\DailyReports\"

'Call the Class Copy File
o1.CSourceFolder = seven_F13_SFolder
o1.CDestFolder = seven_F13_DestUSBFolder
o1.CSourceFile =seven_F13_SFile
o1.Load()


'Call the CLASS MOVE File
o.SourceFolder = seven_F13_SFolder
o.DestFolder = seven_F13_DestFolder
o.SourceFile =seven_F13_SFile
o.Load()


'LOG can
seven_LOG_CAN_SFile   = "F:\7CARD_AD\logfile\adcan\*.csv"
Const seven_Log_SFolder = "F:\7CARD_AD\logfile\adcan\"
Const seven_Log_CAN_DestFolder = "Z:\DATA_AREA\STAGING_AREA\WEB_log\ADCAN\"

'Call the CLASS MOVE File
o.SourceFolder = seven_Log_SFolder
o.DestFolder = seven_Log_CAN_DestFolder
o.SourceFile =seven_LOG_CAN_SFile
o.Load()


'LOG mgi
seven_LOG_MGI_SFile   = "F:\7CARD_AD\logfile\mgi\*.csv"
Const seven_Log_MGI_SFolder = "F:\7CARD_AD\logfile\mgi\"
Const seven_Log_MGI_DestFolder = "Z:\DATA_AREA\STAGING_AREA\WEB_log\ADMGI\"

'Call the CLASS MOVE File
o.SourceFolder = seven_Log_MGI_SFolder
o.DestFolder = seven_Log_MGI_DestFolder
o.SourceFile =seven_LOG_MGI_SFile
o.Load()

'YES Files can
seven_YES_CAN_SFile = "F:\7CARD_AD\yesfile\adcan\*.txt"
Const seven_Yes_can_SFolder = "F:\7CARD_AD\yesfile\adcan\"
Const seven_Yes_can_DestFolder = "Z:\DATA_AREA\STAGING_AREA\WEB_YES_File\"

'Call the CLASS MOVE File
o.SourceFolder = seven_Yes_can_SFolder
o.DestFolder = seven_Yes_can_DestFolder
o.SourceFile =seven_YES_CAN_SFile
o.Load()

'YES Files mgi
seven_YES_MGI_SFile = "F:\7CARD_AD\yesfile\mgi\*.txt"
Const seven_Yes_mgi_SFolder = "F:\7CARD_AD\yesfile\mgi\"
Const seven_Yes_mgi_DestFolder = "Z:\DATA_AREA\STAGING_AREA\WEB_YES_File\"

'Call the CLASS MOVE File
o.SourceFolder = seven_Yes_mgi_SFolder
o.DestFolder = seven_Yes_mgi_DestFolder
o.SourceFile =seven_YES_MGI_SFile
o.Load()

If Err.Number = 0 Then
	msgbox"File moved!",64,"Success"
ELSE
	WScript.Echo "File Move Failed: " & Err.Description
End If

'Clean up! 
Set o = Nothing
'END Move the files to destination folders: Yesfile, F13, Load files--------


'---------------------------------------------------------------------------
'--Start the procedure to load the files and create the reports.------------
'---------------------------------------------------------------------------

Set objWshShell = CreateObject("Wscript.Shell")

 errorCode = objWshShell.Run("C:\Users\williamwada\Desktop\test.cmd", 1,true)
WScript.Echo "ErrorCode" & errorCode 

If errorCode <> 0 Then
    WScript.Echo "Error: " & errorCode    
else
 WScript.Echo "Sucessful run!" & errorCode 
End If

Set objWshShell = Nothing




'END --Start the procedure to load the files and create the reports.--------



'---------------------------------------------------------------------------
'COPY THE REPORTS TO USB ---------------------------------------------------
'---------------------------------------------------------------------------

'set yesterday date
d = date() - 1
dateFinal= Cstr(year(d) * 10000 + month(d) * 100 + day(d))



'-- EPOS REPORTS

epos_Rep_File   = "EPOS_WEB_Sales_"&dateFinal&".xls"
epos_Rep_SFolder = "Y:\Reports\WebEpos\EPOS_WEB_Sales_"&dateFinal&"\"
const epos_Rep_DestUSBFolder = "I:\DailyReports\"

'Call the Class Copy File
o1.CSourceFolder = epos_Rep_SFolder
o1.CDestFolder = epos_Rep_DestUSBFolder
o1.CSourceFile =epos_Rep_SFolder&epos_Rep_File
o1.Load()



'-- 7CARD REPORTS

seven_Rep_File   = "7Card_WEB_Sales_"&dateFinal&".xls"
seven_Rep_SFolder = "Z:\Reports\DAM_7Card\7Card_WEB_Sales_"&dateFinal&"\"
const seven_Rep_DestUSBFolder = "I:\DailyReports\"

'Call the Class Copy File
o1.CSourceFolder = seven_Rep_SFolder
o1.CDestFolder = seven_Rep_DestUSBFolder
o1.CSourceFile =seven_Rep_SFolder&seven_Rep_File
o1.Load()




'Clean up! 
Set o1 = Nothing

'END COPY THE REPORTS TO USB ----------------------------------------------


'---------------------------------------------------------------------------
' CLASS MOVE File
'---------------------------------------------------------------------------
Class moveFiles
   
   Public SourceFolder
   Public DestFolder
   Public SourceFile
 
    Public sub Load()

	  'set objects
	  Dim objFolder 
	  Dim objFile 
	  Set fso = CreateObject("Scripting.FileSystemObject")
	  Set objFolder = fso.GetFolder(SourceFolder) 
	  Set DataFiles = objFolder.Files
	  NumberOfFiles = DataFiles.Count

	    ' If it doesn't have any files display a message
		If NumberOfFiles= 0 Then
		msgbox"File not found!",vbError,"Error"
		Else
			'Loop through the Files collection 
			For Each objFile in objFolder.Files 
				
				' Check the files type
				'if right(fso.GetFileName(objFile),3) = "csv" then 
					fso.moveFile SourceFile, DestFolder
						
				'end if 
			Next 
		
		end if
    End sub
 
End Class

'END CLASS MOVE File 



'---------------------------------------------------------------------------
' CLASS COPY File
'---------------------------------------------------------------------------
Class copyFiles
   
   Public CSourceFolder
   Public CDestFolder
   Public CSourceFile
 
    Public sub Load()

	  'set objects
	  Dim CobjFolder 
	  Dim CobjFile 
	  Set Cfso = CreateObject("Scripting.FileSystemObject")
	  Set CobjFolder = Cfso.GetFolder(CSourceFolder) 
	  Set CDataFiles = CobjFolder.Files
	  CNumberOfFiles = CDataFiles.Count

	    ' If it doesn't have any files display a message
		If CNumberOfFiles= 0 Then
		msgbox"File not found!",vbError,"Error"
		Else
			'Loop through the Files collection 
			For Each CobjFile in CobjFolder.Files 
				
				' Check the files type
				'if right(Cfso.GetFileName(CobjFile),3) = "xls" then 
					Cfso.CopyFile CSourceFile, CDestFolder
					'msgbox"File copied!",64,"Success"	
			'	end if 
			Next 
		
		end if
    End sub
 
End Class

'END CLASS Copy File 




'//End



















