
'------EPOS-------------------------------------------------------------------

'set paths---
'F13
EP_F13_SFile   = "C:\test\test1\*.csv"
Const EP_F13_SFolder = "C:\test\test1\"
Const EP_F13_DestFolder = "C:\test\test2\"

Dim o : Set o = New moveFiles
o.SourceFolder = EP_F13_SFolder
o.DestFolder = EP_F13_DestFolder
o.SourceFile =EP_F13_SFile
msgbox"Chegou aqui",64,"Success"
o.Load()


EP_SFile   = "C:\test\test3\*.csv"
Const EP_SFolder = "C:\test\test3\"
Const EP_DestFolder = "C:\test\test2\"

o.SourceFolder = EP_SFolder
o.DestFolder = EP_DestFolder
o.SourceFile =EP_SFile
msgbox"Chegou aqui2",64,"Success"
o.Load()

'Clean up! 
Set o = Nothing







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
				if right(fso.GetFileName(objFile),3) = "csv" then 
					fso.moveFile SourceFile, DestFolder
					msgbox"File moved!",64,"Success"	
				end if 
			Next 
		
		end if
    End sub
 
End Class