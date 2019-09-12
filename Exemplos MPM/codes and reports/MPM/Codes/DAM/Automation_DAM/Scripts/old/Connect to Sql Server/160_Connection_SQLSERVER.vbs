Dim sServer,sDataBaseName ,sConn,oConn,oRS
Dim sUserName,sPassWord 
 
sServer = "CFDBSVR-7FI"
sDataBaseName = "DDS_7FI_PRD"
sUserName = "ad_7fi_user"
sPassWord = "ad_7fi_160"

'====CONNECTION 
sConn="DRIVER={SQL Server};SERVER=" & sServer & ";DATABASE=" & sDataBaseName & ";Encrypt=No;"
Set oConn = CreateObject("ADODB.Connection")
'oConn.CommandTimeout = 36000
oConn.Open sConn, sUserName, sPassWord
'END=====


Set FetchData = CreateObject("ADODB.Recordset")
FetchData.open "select count(*) from dbo.DDS_CDR", oconn
While Not FetchData.eof 
          MsgBox FetchData.Fields(0) 
		  FetchData.movenext
Wend

'Clean up! 
Set oConn = Nothing