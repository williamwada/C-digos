'=================================================================
' ==============Zip folder using 7-Zip
' Written by William Wada
'=================================================================

'set yesterday date
d = date() - 1
dateFinal= Cstr(year(d) * 10000 + month(d) * 100 + day(d))
dateEmail= month(d)&"/"&day(d)

'--EPOS_F13
 srcFolder="D:\EPOS_PRD\DATA_AREA\IN\DAM\F13\HISTORIC"		    		     				 	'Source Folder
 srcFile = "D:\EPOS_PRD\DATA_AREA\IN\DAM\F13\HISTORIC\AF13_EPOS_WEB_SALES_"&dateFinal&".xls"	'Source File
 dstFolder="D:\EPOS_PRD\DATA_AREA\IN\DAM\F13\HISTORIC\"            		 						'Destination folder
 FileExt = ".zip"                									 							' File extension
 password ="0101"						 														' Password
 

 Call zipFileWithPass(srcFolder,srcFile,dstFolder,FileExt,password)

 'Delete the file after zip
 Const DeleteReadOnly = True
 Set obj = CreateObject("Scripting.FileSystemObject")
 obj.DeleteFile(srcFile),DeleteReadOnly
 
 

'--EPOS_Report
 srcFolder="D:\EPOS_PRD\REPORTS\DAM_Reports"		    		     										'Source Folder
 srcFile = "D:\EPOS_PRD\REPORTS\DAM_Reports\WEB_EPOS_Sales_"&dateFinal&"\WEB_EPOS_Sales_"&dateFinal&".xls"  'Sourcer File
 dstFolder="D:\EPOS_PRD\REPORTS\DAM_Reports\"            		 											'Destination folder
 FileExt = ".zip"                		 																	' File extension
 
 Call zipFile(srcFolder,srcFile,dstFolder,FileExt)
 
 
'----------------------------------------------------------------
'ZIP WIth PASS FUNCTION
'----------------------------------------------------------------

Function zipFileWithPass(srcFolder,srcFile,dstFolder, fileExt,password)

Dim fso,shell

 zipEXE="C:\Program Files\7-Zip\7z.exe"  '7zip program path

 
' File System Object
Set fso = CreateObject("Scripting.FileSystemObject")
' Shell
Set shell = CreateObject("WScript.Shell")


	If Not fso.FileExists(srcFile) Then 
		 MsgBox "The " & srcFile & " doesn't exist ",48,zipEXE 
	 ELSE
	 
		filename = fso.GetBaseName(srcFile)
		ZipedFile = dstFolder&Filename&fileExt   ' Zip file extension
	    
		''''Verify if 7zip exist
		 If Not fso.FileExists(zipEXE) Then 
		 MsgBox "The " & zipEXE & " doesn't exist ",48,zipEXE 
		 ELSE
		 
			'''' Zip Source Folder
								
			' Change to source directory
			shell.CurrentDirectory = srcFolder & "/"
			shCommand = """" & zipEXE & """ a -r """ & ZipedFile & """ -p""" & password & """ """ & srcFile & """"
			
			' Run 7-Zip in shell
			shVal = shell.Run(shCommand,4,true)
			 
			' Check 7-Zip exit code
			If shVal > 1 Then
				Wscript.Echo "7-Zip failed with error code: " & shVal
				Wscript.Quit
			'Else
			'	Wscript.Echo "7-Zip Success zipped!"
				
			End If
		 End if
	 End if
 
End Function

'----------------------------------------------------------------
'ZIP WITHOUT PASS FUNCTION
'----------------------------------------------------------------

Function zipFile(srcFolder,srcFile,dstFolder, fileExt)

Dim fso,shell

 zipEXE="C:\Program Files\7-Zip\7z.exe"  '7zip program path

 
' File System Object
Set fso = CreateObject("Scripting.FileSystemObject")
' Shell
Set shell = CreateObject("WScript.Shell")


	If Not fso.FileExists(srcFile) Then 
		 MsgBox "The " & srcFile & " doesn't exist ",48,zipEXE 
	 ELSE
	 
		filename = fso.GetBaseName(srcFile)
		ZipedFile = dstFolder&Filename&fileExt   ' Zip file extension
	    
		''''Verify if 7zip exist
		 If Not fso.FileExists(zipEXE) Then 
		 MsgBox "The " & zipEXE & " doesn't exist ",48,zipEXE 
		 ELSE
		 
			'''' Zip Source Folder
								
			' Change to source directory
			shell.CurrentDirectory = srcFolder & "/"
			shCommand = """" & zipEXE & """ a -r """ & ZipedFile & """ """ & srcFile & """"
		
			' Run 7-Zip in shell
			shVal = shell.Run(shCommand,4,true)
			 
			' Check 7-Zip exit code
			If shVal > 1 Then
				Wscript.Echo "7-Zip failed with error code: " & shVal
				Wscript.Quit
				
			End If
		 End if
	 End if
 
End Function
