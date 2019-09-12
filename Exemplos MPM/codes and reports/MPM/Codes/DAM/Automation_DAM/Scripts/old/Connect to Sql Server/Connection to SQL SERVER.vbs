Dim sServer,sDataBaseName ,sConn,oConn,oRS
Dim sUserName,sPassWord 
 
sServer = "WILLIAM\MPMJAPAN"
sDataBaseName = "Northwind"
sUserName = "sa"
sPassWord = "vaga1109"

'====CONNECTION 
sConn="DRIVER={SQL Server};SERVER=" & sServer & ";DATABASE=" & sDataBaseName & ";Encrypt=No;"
Set oConn = CreateObject("ADODB.Connection")
'oConn.CommandTimeout = 36000
oConn.Open sConn, sUserName, sPassWord
'END=====


Set FetchData = CreateObject("ADODB.Recordset")
FetchData.open "SELECT COUNT (*) as UserName FROM [Northwind].[dbo].[Customers]", oconn
While Not FetchData.eof 
          MsgBox FetchData.Fields(0) 
		  FetchData.movenext
Wend

'Clean up! 
Set oConn = Nothing