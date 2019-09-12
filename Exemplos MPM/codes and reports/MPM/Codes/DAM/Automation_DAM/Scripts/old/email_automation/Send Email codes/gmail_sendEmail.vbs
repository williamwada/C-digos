'=======================
Dim xstrAll(), strMessage,intupper()
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Send the mail
SentMail

 
 '=======================SUB =======================
 Sub SentMail

 '===========================================================================
 'CHANGE DATA HERE ->Email Address (from and to), Subject, Message, Account ID, Password
  strMessage =  strMessage & "Your other message..."
  strFrom="xxxxxxxx@gmail.com"
  strTo= "william.wada@mpmjapan.com" 
  strSubject="[TEST]Automated Alert Message: Server Notification"
  strAccountID="xxxxxxxx@gmail.com"
  strPassword="xxxxxxxxx"
'==========================================================================

  strSMTPServer="smtp.gmail.com"
 SendMail strFrom,strTo,strSubject,strMessage,strAccountID,strPassword,strSMTPServer 

 End Sub
 '===========End SUB

 
 '=======================FUNCTIONS =======================
  Function SendMail( strFrom, strSendTo, strSubject, strMessage , strUser, strPassword, strSMTP )

     Set oEmail = CreateObject("CDO.Message")
    
    'configure message
    With oEmail.Configuration.Fields
          .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
          .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSMTP
          .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = true
		  .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") =465     
		'  .item("http://schemas.microsoft.com/cdo/configuration/StartTLS") = true   '===>This commented line maybe necessary in some instances together with "smtpusessl"
          .item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 
          .item("http://schemas.microsoft.com/cdo/configuration/sendusername") = strUser
          .item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = strPassword
          
          .Update
    End With
    
    ' build message
    With oEmail
         .From = strFrom
         .To = strSendTo 
         .Subject = strSubject
         .TextBody = strMessage
    End With
    
    ' send message
    On Error Resume Next
    oEmail.Send
    If Err Then
         WScript.Echo "SendMail Failed:" & Err.Description
    End If
        
End Function
'=========END FUNCTIONS