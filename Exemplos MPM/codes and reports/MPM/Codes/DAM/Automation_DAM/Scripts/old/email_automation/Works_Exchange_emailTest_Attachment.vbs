'=======================
Dim xstrAll(), strMessage,intupper()
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Send the mail
SentMail

 
 
 Sub SentMail

  strTo= "williampswada@gmail.com" 
  strFrom="william@mpmjapan.com"
  strSubject="Automated Alert Message: Server Notification"
  strMessage =  strMessage & "Your other message..."  
  strAttach = "C:\Users\williamwada\Desktop\msgdiego.txt"
  strAccountID="william@mpmjapan.com"
  strPassword="JETStre@m110910"
  'strSMTPServer="smtp.office365.com"
  strSMTPServer="pod51010.outlook.com"
  
  
  SendMail strFrom,strTo,strSubject,strMessage,strAccountID,strPassword,strSMTPServer,strAttach
 
 End Sub

 

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
