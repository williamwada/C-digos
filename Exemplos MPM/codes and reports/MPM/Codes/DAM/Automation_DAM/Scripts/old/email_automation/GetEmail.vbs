
ReceiveMail "pod51010.outlook.com","william@mpmjapan.com","JETSTre@m110910",true





Sub ReceiveMail( _
ByVal sServer, _
ByVal sUserName, _
ByVal sPassword, _
ByVal bSSLConnection)
    
    Const MailServerPop3 = 0
    Const MailServerImap4 = 1
    Const MailServerEWS = 2
    Const MailServerDAV = 3

    'For evaluation usage, please use "TryIt" as the license code, otherwise the
    '"invalid license code" exception will be thrown. However, the object will expire in 1-2 months, then
    '"trial version expired" exception will be thrown.
    Dim oClient
    Set oClient = CreateObject("EAGetMailObj.MailClient")
    oClient.LicenseCode = "TryIt"

    'To receive email from imap4 server, please change
    'MailServerPop3 to MailServerImap4 to MailServer.Protocol

    'To receive email with Exchange Web Service, please change
    'MailServerPop3 to MailServerEWS to MailServer.Protocol

    'To receive email with Exchange WebDAV, please change
    'MailServerPop3 to MailServerDAV to MailServer.Protocol

    'Exchange Server supports POP3/IMAP4 protocol as well, but in Exchange 2007
    'or later version, POP3/IMAP4 service is disabled by default. If you don't want to use POP3/IMAP4
    'to download email from Exchange Server, you can use Exchange Web Service(Exchange 2007/2010 or
    'later version) or WebDAV(Exchange 2000/2003) protocol.
    Dim oServer
    Set oServer = CreateObject("EAGetMailObj.MailServer")
    oServer.Server = sServer
    oServer.User = sUserName
    oServer.Password = sPassword
    oServer.SSLConnection = bSSLConnection
    oServer.Protocol =  MailServerEWS
    
    ''by default, the pop3 port is 110, imap4 port is 143,
    'the pop3 ssl port is 995, imap4 ssl port is 993
    'you can also change the port like this
    'oServer.Port = 110
    If oServer.Protocol = MailServerImap4 Then
        If oServer.SSLConnection Then
            oServer.Port = 993 'SSL IMAP4
        Else
            oServer.Port = 143 'IMAP4 normal
        End If
    Else
        If oServer.SSLConnection Then
            oServer.Port = 995 'SSL POP3
        Else
            oServer.Port = 110 'POP3 normal
        End If
    End If
    

    oClient.Connect oServer
    Dim infos
    infos = oClient.GetMailInfos()
    
    Dim i, Count
    Count = UBound(infos)
    For i = LBound(infos) To Count
        Dim info
        Set info = infos(i)
        
        MsgBox "UIDL: " & info.UIDL
        MsgBox "Index: " & info.Index
        MsgBox "Size: " & info.Size
        'For POP3/Exchange Web Service/WebDAV, the IMAP4MailFlags is meaningless.
        MsgBox "Flags: " & info.IMAP4Flags 
        'For POP3, the Read is meaningless.
        MsgBox "Read: " & info.Read
        MsgBox "Deleted: " & info.Deleted
                    
        Dim oMail
        Set oMail = oClient.GetMail(info)
        'Save mail to local
        oMail.SaveAs "c:\tempfolder\" & i & ".eml", True
    Next

    For i = LBound(infos) To Count
        oClient.Delete (infos(i))
    Next

    '' Delete method just mark the email as deleted,
    ' Quit method purge the emails from server exactly.
    oClient.Quit
End Sub
