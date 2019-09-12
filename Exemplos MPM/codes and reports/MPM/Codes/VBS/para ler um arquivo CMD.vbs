Dim wshshell 
intOpcao = msgbox("Deseja que o windows abra o painel de controle?",vbyesno,"Windows") 
if intOpcao = vbyes then 
     Set WshShell = WScript.CreateObject("WScript.Shell") 
  WshShell.Run("%systemroot%\system32\control.exe")