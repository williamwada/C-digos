
Set objWshShell = CreateObject("Wscript.Shell")

 'objWshShell.run "C:\Users\williamwada\Desktop\test.cmd", 1

 errorCode = objWshShell.Run("C:\Users\williamwada\Desktop\test.cmd", 1,true)
WScript.Echo "ErrorCode" & errorCode 

If errorCode <> 0 Then
    WScript.Echo "Error: " & errorCode
      
else
 WScript.Echo "Sucessful run!" & errorCode 
End If



Set objWshShell = Nothing