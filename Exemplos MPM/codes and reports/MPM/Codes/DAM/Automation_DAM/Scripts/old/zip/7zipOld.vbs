
' Zip folder using 7-Zip
' Written by William Wada

Dim fso,shell
 
' File System Object
Set fso = CreateObject("Scripting.FileSystemObject")
' Shell
Set shell = CreateObject("WScript.Shell")


srcFile="C:\test"		    		 'Source Folder
dstFolder="C:\test2"            		 'Destination folder
FileExt = ".zip"                		 ' File extension
password ="test"						 ' Password
ZipedFile = dstFolder&"\ziped"&FileExt   ' Zip file extension
zipEXE="C:\Program Files\7-Zip\7z.exe"  '7zip program path


 'If Not fso.FileExists(srcFile) Then 
'	 MsgBox "The " & srcFile & " doesn't exist ",48,zipEXE 
 'ELSE
 
	''''Verify if 7zip exist
	 If Not fso.FileExists(zipEXE) Then 
	 MsgBox "The " & zipEXE & " doesn't exist ",48,zipEXE 
	 ELSE
	 
		'''' Zip Source Folder
		' Create shell command /if you dont want to set password comment the [-p"&password]
		
		' Change to source directory
		shell.CurrentDirectory = srcFile & "/"
		
		shCommand = """" & zipEXE & """ a -r """ & ZipedFile & """-p """&password& """"
		
		' Run 7-Zip in shell
		shVal = shell.Run(shCommand,4,true)
		 
		' Check 7-Zip exit code
		If shVal > 1 Then
			Wscript.Echo "7-Zip failed with error code: " & shVal
			Wscript.Quit
		Else
			Wscript.Echo "7-Zip Success zipped!"
			Wscript.Quit
		End If
	 End if
' End if
 