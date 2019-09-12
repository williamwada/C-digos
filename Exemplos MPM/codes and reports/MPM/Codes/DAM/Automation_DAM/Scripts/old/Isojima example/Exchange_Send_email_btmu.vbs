'=======================
Dim xstrAll(), strMessage,intupper()
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Send the mail
SentMail

 
 
 Sub SentMail

'===========================================================================
 'CHANGE DATA HERE ->Email Address (from and to), Subject, Message, Account ID, Password
  strTo= "tsukasa@mpmjapan.com" 
  strFrom="tsukasa@mpmjapan.com"
  'strSubject="Automated Alert Message: Server Notification"
  strMessage =  strMessage & "Dear All,"& vbCr & vbCr &"Please find the attached daily BCCF Report."& vbCr & vbCr & "Best regards,"& vbCr &"Isojima"   ' 本文を指定 
  strAccountID="tsukasa@mpmjapan.com"
  strPassword="Mktt0376"
'==========================================================================

  'strSMTPServer="smtp.office365.com"
  strSMTPServer="pod51010.outlook.com"
  
  
  SendMail strFrom,strTo,strSubject,strMessage,strAccountID,strPassword,strSMTPServer
 
 End Sub

 

'=======================FUNCTIONS =======================
 Function SendMail( strFrom, strSendTo, strSubject, strMessage , strUser, strPassword, strSMTP )
	 
	'Take the file info to the folder
	Set objFolder = objFSO.GetFolder("\\172.31.24.179\reports")
	 
	'Take file obeject from  folder object of files property
	For Each objFile In objFolder.Files


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
    
    ' build message
    With oEmail
         .From = strFrom
         .To = strSendTo 
         .Subject = objFSO.getBaseName(objFile.Name)
         .TextBody = strMessage
         .AddAttachment objFolder&"\"&objFile.Name 
    End With
    
    ' send message
    On Error Resume Next
	wscript.echo "Att:"&objFolder&"\"&objFile.Name 
  '  oEmail.Send
    Dim strSrcFile      ' 移動対象ファイル
    Dim strDstFile      ' 移動先ファイル名

    strSrcFile = objFSO.GetFile(objFolder&"\"&objFile.Name)
    strDstFile = objFolder&"\Archive\"&objFile.Name

    If Err.Number = 0 Then
       objFSO.MoveFile strSrcFile, strDstFile
    If Err.Number = 0 Then
       WScript.Echo strSrcFile & " を " & _
            strDstFile & " に移動しました。"
    Else
        WScript.Echo "エラー: " & Err.Description
    End If
    
    Else
    WScript.Echo "エラー: " & Err.Description
    End If

    Set objFSO = Nothing
    
    If Err Then
         WScript.Echo "SendMail Failed:" & Err.Description
         With oEmail
         .From = strFrom
         .To = strSendTo 
         .Subject = "Error Message"&objFSO.getBaseName(objFile.Name)
         .TextBody = "エラー: " & Err.Description
         .send
    End With
        
    End If

   Next
 End Function
 '====END FUNCTIONS 
