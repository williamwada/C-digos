' VBScript to Send Email Notification
' Author: https://helloacm.com
' Usage: cscript.exe sendemail.vbs email subject text
' 23/Dec/2014
 
Sub SendEmail(ToAddress, Subject, Text)
    Dim iMsg 
    Dim iConf
    Dim Flds
	
	 
    Set iMsg = CreateObject("CDO.Message")
    Set iConf = CreateObject("CDO.Configuration")
 
    iConf.Load -1
    Set Flds = iConf.Fields
    
    With Flds
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "william@mpmjapan.com"
        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "JETStre@m110910"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.office365.com" 'smtp mail server
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 'stmp server
        .Update
    End With
 
    With iMsg
        Set .Configuration = iConf
        .To = ToAddress
        .From = "william.wada@mpmjapan.com"
        .Subject = Subject
        .TextBody = Text
        .Send
    End With
 
    Set iMsg = Nothing
    Set iConf = Nothing
End Sub
 
If WScript.Arguments.Count <> 3 Then
    WScript.Echo "Usage: cscript.exe " & WScript.ScriptFullName & " email subject text"
Else 
    SendEmail WScript.Arguments(0), WScript.Arguments(1), WScript.Arguments(2)
End If