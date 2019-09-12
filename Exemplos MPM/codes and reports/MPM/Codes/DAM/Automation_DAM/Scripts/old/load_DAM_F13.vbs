'########################################################################
'Description: Move and load the F13 files from AP server to Epos DB server
'Create by  : William Wada 2017/04/10
'########################################################################



'set objects
Dim move : Set move = New moveFiles
'Dim copy : Set copy = New copyFiles
Dim objFolder 
Dim objFile 

Set fso = CreateObject("Scripting.FileSystemObject")
Set objFolder = fso.GetFolder(EP_F13_DestFolder) 
Set DataFiles = objFolder.Files
NumberOfFiles = DataFiles.Count


'------EPOS
'F13
EP_F13_SFile   = "\\172.31.20.38\ad_epos_prd\DATA_AREA\IN\DAM\f13\*.xls"

Const EP_F13_SFolder = "\\172.31.20.38\ad_epos_prd\DATA_AREA\IN\DAM\f13\"
Const EP_F13_DestFolder = "D:\EPOS_PRD\DATA_AREA\IN\DAM\F13\"
Const EP_F13_Dest_Hist_Folder = "D:\EPOS_PRD\DATA_AREA\IN\DAM\F13\HISTORIC\"
Const EP_F13_S_H_Folder = "D:\EPOS_PRD\DATA_AREA\IN\DAM\F13\"
'Move File 
move.SourceFolder = EP_F13_SFolder
move.DestFolder = EP_F13_DestFolder
move.SourceFile =EP_F13_SFile
move.Load()

'Clean up! 
Set move = Nothing

'LOAD F13##############################

'Set DataFiles = objFolder.Files

If NumberOfFiles= 0 Then
'msgbox"File not found!",vbError,"Error"
Wscript.Echo "File not found! "&vbError
Else
	'Loop through the Files collection 
	For Each objFile in objFolder.Files 
		Wscript.Echo "File name! "+objFile.Name
		'Call Procedure to load
		Call ExecuteSql( "EXEC Load_F13_DAM '"+objFile.Name+"'" )
	Next 

end if

'Clean up! 
Set objFile = Nothing

'Move File to Historic
		Dim moveH : Set moveH = New moveFiles
		moveH.SourceFolder = EP_F13_S_H_Folder
		moveH.DestFolder = EP_F13_Dest_Hist_Folder
		moveH.SourceFile =EP_F13_S_H_Folder &"*.xls"
		moveH.Load()
	
'Clean up! 
Set moveH = Nothing

'CREATE THE REPORT###############################################

Call ExecuteSql( "EXEC DAM_EposDaily_Report" )
	
	
'ZIP THE F13 and REPORT##########################################




	
	
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
		'msgbox"File not found!",vbError,"Error"
		Wscript.Echo "File not found! "&vbError
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


'---------------------------------------------------------------------------
' FUNCTION Execute SQL
'---------------------------------------------------------------------------

Public Sub ExecuteSql( sqlString )

Dim sServer, sConn, oConn, sDatabaseName, sUser, sPassword
sDatabaseName="DDS_EPOS_PRD"
sServer="CFDBSVR-EPOS"
sUser="ad_epos_user"
sPassword="ad_epos_159"
sConn="provider=sqloledb;data source=" & sServer & ";initial catalog=" & sDatabaseName

Set oConn = CreateObject("ADODB.Connection")
oConn.Open sConn, sUser, sPassword
oConn.Execute sqlString

WScript.Echo "executed "&sqlString
oConn.Close
Set oConn = Nothing
End Sub