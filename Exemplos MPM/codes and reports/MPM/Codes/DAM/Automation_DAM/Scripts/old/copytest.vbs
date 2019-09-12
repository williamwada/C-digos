DIM o1: Set o1 = New copyFiles

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


