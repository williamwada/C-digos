'########################################################################
'Description: Move and load the F13 files from AP server to Epos DB server
'Create by  : William Wada 2017/04/10
'########################################################################



'set objects
'Dim move : Set move = New moveFiles
Dim copy : Set copy = New copyFiles
Dim objFolder 
Dim objFile 
Dim Desc, CDesc

d = date() - 1
dateFinal= Cstr(year(d) * 10000 + month(d) * 100 + day(d))

'------EPOS
'F13
EP_F13_SFile   = "\\172.31.20.38\ad_epos_prd\DATA_AREA\IN\DAM\f13\*.xls"

Const EP_F13_SFolder = "\\172.31.20.38\ad_epos_prd\DATA_AREA\IN\DAM\f13\"
Const EP_F13_DestFolder = "\\172.31.31.159\shared\DAM_EPOS\F13\"
Const EP_F13_Dest_Hist_Folder = "\\172.31.31.159\shared\DAM_EPOS\F13\HISTORIC\"
Const EP_F13_S_H_Folder = "\\172.31.31.159\shared\DAM_EPOS\F13\"
'Copy File 
copy.CSourceFolder = EP_F13_SFolder
copy.CDestFolder = EP_F13_DestFolder
copy.CSourceFile =EP_F13_SFile
copy.CDesc = "F13"
copy.Load()



'LOAD F13##############################

Set fso = CreateObject("Scripting.FileSystemObject")
Set objFolder = fso.GetFolder(EP_F13_DestFolder) 
Set DataFiles = objFolder.Files
NumberOfFiles = DataFiles.Count


If NumberOfFiles= 0 Then
'msgbox"File not found!",vbError,"Error"
Wscript.Echo "File F13 not found! "&vbError
Else
	'Loop through the Files collection 
	For Each objFile in objFolder.Files 
		Wscript.Echo "File name: "+objFile.Name
		'Call Procedure to load
		Call ExecuteSql( "EXEC Load_F13_DAM '"+objFile.Name+"'" )
	Next 

end if


'Move File to Historic
'		Dim moveH : Set moveH = New moveFiles
'		moveH.SourceFolder = EP_F13_S_H_Folder
'		moveH.DestFolder = EP_F13_Dest_Hist_Folder
'		moveH.SourceFile =EP_F13_S_H_Folder &"*.xls"
'		moveH.Load()

'Delete the file after load
 Const DeleteReadOnly = True
 Set obj = CreateObject("Scripting.FileSystemObject")
 obj.DeleteFile(EP_F13_S_H_Folder &"*.xls"),DeleteReadOnly
	
'Clean up! 
'Set moveH = Nothing

'CREATE THE REPORT###############################################

Call ExecuteSql( "EXEC DAM_EposDaily_Report "+CStr(NumberOfFiles) )
	
	

'Move File to 38

'EP_Report_File   = "D:\EPOS_PRD\REPORTS\DAM_Reports\WEB_EPOS_Sales_"&dateFinal&"\WEB_EPOS_Sales_"&dateFinal&".xls"
'Const EP_Report_Dest_Folder = "D:\shared\DAM_EPOS\reports\"
'EP_Rep_S_Folder    = "D:\EPOS_PRD\REPORTS\DAM_Reports\WEB_EPOS_Sales_"&dateFinal&"\"
'Desc = "Report"

'		move.SourceFolder = EP_Rep_S_Folder
'		move.DestFolder = EP_Report_Dest_Folder
'		move.SourceFile = EP_Report_File   
'		move.desc = Desc
'		move.Load()

'Clean up! 
'Set move      = Nothing
Set obj       = Nothing
Set objFile   = Nothing
Set objFolder = Nothing
Set copy      = Nothing
	
If Err.Number = 0 Then
	WScript.Echo "2.Copy_load_Create_Report Finished Sucessfull!"
	 Wscript.Quit
ELSE
	WScript.Echo "2.Copy_load_Create_Report Failed! Error: " & Err.Description
	 Wscript.Quit
End If	
	
'---------------------------------------------------------------------------
' CLASS MOVE File
'---------------------------------------------------------------------------
Class moveFiles
   
   Public SourceFolder
   Public DestFolder
   Public SourceFile
   Public Desc
 
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
		
		Wscript.Echo "File "&Desc&" not found! "&vbError
		Else
			'Loop through the Files collection 
			For Each objFile in objFolder.Files 
				
				' Check the files type
				'if right(fso.GetFileName(objFile),3) = "csv" then 
					fso.moveFile SourceFile, DestFolder
						
				'end if 
			Next 
		
		end if
		
'Clean up! 
Set objFile   = Nothing
Set objFolder = Nothing
Set fso       = Nothing
Set DataFiles = Nothing
		
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
   Public CDesc
 
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
		Wscript.Echo "File "&Desc&" not found! "&vbError
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
		
		'Clean up! 
		Set CobjFile   = Nothing
		Set CobjFolder = Nothing
		Set Cfso       = Nothing
		Set CDataFiles = Nothing
		
    End sub
 
End Class

'END CLASS Copy File 


'---------------------------------------------------------------------------
' FUNCTION Execute SQL
'---------------------------------------------------------------------------

Public Sub ExecuteSql( sqlString )

Dim sServer, sConn, oConn, sDatabaseName, sUser, sPassword
sDatabaseName="DDS_EPOS_PRD"
sServer="172.31.31.159"
sUser="ad_epos_user"
sPassword="ad_epos_159"
sConn="provider=sqloledb;data source=" & sServer & ";initial catalog=" & sDatabaseName

Set oConn = CreateObject("ADODB.Connection")
oConn.Open sConn, sUser, sPassword
oConn.Execute sqlString

WScript.Echo "executed "&sqlString
oConn.Close

'Clean up! 
Set oConn = Nothing

End Sub