'=================================================================
' ==============Zip folder using 7-Zip
' Written by William Wada
'=================================================================


Dim a : a=3
Const DeleteReadOnly = True
    For i=1 to a Step 1 'i is the counter variable and it is incremented by 1
     
		'set date
		d = date() - i
		dateFinal= Cstr(year(d) * 10000 + month(d) * 100 + day(d))
		dateEmail= month(d)&"/"&day(d)

		'--EPOS_F13
		 srcFolder="D:\AD_EPOS_PRD\DATA_AREA\IN\DAM\f13"		    		     				 	'Source Folder
		 srcFile = "D:\AD_EPOS_PRD\DATA_AREA\IN\DAM\f13\AF13_EPOS_WEB_SALES_"&dateFinal&".xls"	'Source File
		 dstFolder="D:\AD_EPOS_PRD\DATA_AREA\IN\DAM\f13\"            		 						'Destination folder
		 FileExt = ".zip"                									 							' File extension
		 password ="0101"						 														' Password
			
			Wscript.echo "FILE: "& srcFile
		
		 Call zipFileWithPass(srcFolder,srcFile,dstFolder,FileExt,password)

		 'Delete the file after zip
		 
		 Set obj = CreateObject("Scripting.FileSystemObject")
		 obj.DeleteFile(srcFile),DeleteReadOnly

 	Next
 
 'Clean up! 
Set obj = Nothing

d = date()-1
dateFinal= Cstr(year(d) * 10000 + month(d) * 100 + day(d))

'--EPOS_Report
 srcFolder="\\172.31.31.159\shared\DAM_EPOS\reports\WEB_EPOS_Sales_"&dateFinal		    		     										'Source Folder
 srcFile = "\\172.31.31.159\shared\DAM_EPOS\reports\\WEB_EPOS_Sales_"&dateFinal&"\WEB_EPOS_Sales_"&dateFinal&".xls"  'Sourcer File
 dstFolder="\\172.31.20.38\ad_epos_prd\Reports\DAM\Daily_Reports\"            		 											'Destination folder
 FileExt = ".zip"                		 																	' File extension
 Wscript.echo  "REPORT_FILE: "& srcFile
 Call zipFile(srcFolder,srcFile,dstFolder,FileExt)
 


 'Delete the file after zip
 Set obj = CreateObject("Scripting.FileSystemObject")
 obj.DeleteFile(srcFile),DeleteReadOnly
  
 'Clean up! 
 Set obj = Nothing

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
		 Wscript.Echo "The " & srcFile & " doesn't exist " 
	 ELSE
	 
		filename = fso.GetBaseName(srcFile)
		ZipedFile = dstFolder&Filename&fileExt   ' Zip file extension
	    
		''''Verify if 7zip exist
		 If Not fso.FileExists(zipEXE) Then 
		 Wscript.Echo "The " & zipEXE & " doesn't exist " 
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
		 Wscript.Echo "The " & srcFile & " doesn't exist " 
	 ELSE
	 
		filename = fso.GetBaseName(srcFile)
		ZipedFile = dstFolder&Filename&fileExt   ' Zip file extension
	    
		''''Verify if 7zip exist
		 If Not fso.FileExists(zipEXE) Then 
		 Wscript.Echo "The " & zipEXE & " doesn't exist "
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
