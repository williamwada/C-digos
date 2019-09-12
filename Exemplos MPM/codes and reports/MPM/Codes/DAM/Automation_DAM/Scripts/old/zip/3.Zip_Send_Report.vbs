'=================================================================
' ==============Zip folder using 7-Zip
' Written by William Wada
'=================================================================

'set yesterday date
d = date() - 1
dateFinal= Cstr(year(d) * 10000 + month(d) * 100 + day(d))


'--EPOS_F13
 srcFolder="H:\DailyReports"		    		     'Source Folder
 srcFile = "H:\DailyReports\AF13_EPOS_WEB_SALES_"&dateFinal&".xls"
 dstFolder="H:\DailyReports\"            		 'Destination folder
 FileExt = ".zip"                		 ' File extension
 password ="0101"						 ' Password
 

 Call zipFileWithPass(srcFolder,srcFile,dstFolder,FileExt,password)
 
 
'--7CARD_F13
 srcFolder="H:\DailyReports"		    		     'Source Folder
 srcFile = "H:\DailyReports\AF13_7CARD_WEB_SALES_"&dateFinal&".xls"
 dstFolder="H:\DailyReports\"            		 'Destination folder
 FileExt = ".zip"                		 ' File extension
 password ="WPSw#@XsT"						 ' Password

 Call zipFileWithPass(srcFolder,srcFile,dstFolder,FileExt,password)

'--EPOS_Report
 srcFolder="H:\DailyReports"		    		     'Source Folder
 srcFile = "H:\DailyReports\EPOS_WEB_Sales_"&dateFinal&".xls"
 dstFolder="H:\DailyReports\"            		 'Destination folder
 FileExt = ".zip"                		 ' File extension
 
 Call zipFile(srcFolder,srcFile,dstFolder,FileExt)
 
 '--7Card_Report
 srcFolder="H:\DailyReports"		    		     'Source Folder
 srcFile = "H:\DailyReports\7Card_WEB_Sales_"&dateFinal&".xls"
 dstFolder="H:\DailyReports\"            		 'Destination folder
 FileExt = ".zip"                		 ' File extension
 
 Call zipFile(srcFolder,srcFile,dstFolder,FileExt)
 





'=================================================================
' ==============send Email
' Written by William Wada
'=================================================================

'=======================
Dim xstrAll(), strMessage,intupper()
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Send the mail
'SentMail

 
 
 Sub SentMail

  strTo= "williampswada@gmail.com" 
  strFrom="william@mpmjapan.com"
  strSubject="Automated Alert Message: Server Notification"
  strMessage =  strMessage & "Epos"& vbCr & vbCr & dateFinal &"email test"& vbCr
  'strMessage =  strMessage & "株式会社エポスカード様"& vbCr & vbCr &"大変お世話になっております。"& vbCr
  'strMessage =  strMessage & "MPMジャパンのウィリアムです。"& vbCr & dateFinal & " F13データをお送り致します。" & vbCr  ' 本文を指定  
  'strMessage =  strMessage & "ご確認いただきますようお願いいたします。"& vbCr & vbCr &
  'strMessage =  strMessage & "※解凍PWは別途お送り致します。"& vbCr & vbCr &
  'strMessage =  strMessage & "以上、よろしくお願い致します。"& vbCr & vbCr &
  'strMessage =  strMessage & "William"& vbCr & vbCr &
  
  
  strAccountID="william@mpmjapan.com"
  strPassword="JETSTre@m110910"
  'strSMTPServer="smtp.office365.com"
  strSMTPServer="pod51010.outlook.com"
  strAttach = "H:\DailyReports\AF13_EPOS_WEB_SALES_"&dateFinal&".zip"
  
   SendMail strFrom,strTo,strSubject,strMessage,strAccountID,strPassword,strSMTPServer,strAttach
  If Err.Number = 0 Then
	msgbox"Files zipped and sent!",64,"Success"
	 Wscript.Quit
ELSE
	WScript.Echo "Failed: " & Err.Description
	 Wscript.Quit
End If
 

 End Sub


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

'=======================FUNCTIONS =======================
 Function SendMail( strFrom, strSendTo, strSubject, strMessage , strUser, strPassword, strSMTP , strAttach )

     Set oEmail = CreateObject("CDO.Message")
    
    'configure message
    With oEmail.Configuration.Fields
          .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
          .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSMTP
          .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = true
		  .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") =25     
		  .item("http://schemas.microsoft.com/cdo/configuration/StartTLS") = true   
          .item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 
          .item("http://schemas.microsoft.com/cdo/configuration/sendusername") = strUser
          .item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = strPassword          
          
		  .Update
    End With
    
    ' build messaged
    With oEmail
         .From           = strFrom
         .To             = strSendTo 
         .Subject        = strSubject
         .TextBody       = strMessage
		 .AddAttachment strAttach
    End With
	
	wscript.echo "Att: "& strAttach
    ' send message
    On Error Resume Next
    oEmail.Send	 
    If Err Then
         WScript.Echo "SendMail Failed:" & Err.Description
	Else
	WScript.Echo "Sent Sucessfull!"
    End If
    
 End Function
 '====END FUNCTIONS 


 