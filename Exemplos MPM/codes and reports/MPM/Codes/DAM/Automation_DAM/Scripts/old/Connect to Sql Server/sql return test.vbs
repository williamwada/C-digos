'CREATE THE REPORT###############################################

Call ExecuteSql( "select Sales_Count_CAN,Sales_Count_MGI from dbo.TM_WebEposReport where convert(date,sales_date) = convert(date,GETDATE()-1)" )

'---------------------------------------------------------------------------
' FUNCTION Execute SQL
'---------------------------------------------------------------------------

Public Sub ExecuteSql( sqlString )

Dim sServer, sConn, oConn, sDatabaseName, sUser, sPassword , field_can,field_mgi,Recordset
sDatabaseName="DDS_EPOS_PRD"
sServer="CFDBSVR-EPOS"
sUser="ad_epos_user"
sPassword="ad_epos_159"
sConn="provider=sqloledb;data source=" & sServer & ";initial catalog=" & sDatabaseName

Set oConn = CreateObject("ADODB.Connection")
Set Recordset = CreateObject("ADODB.Recordset")

oConn.Open sConn, sUser, sPassword
Recordset.Open sqlString,oConn

'first of all determine whether there are any records 

If Recordset.EOF Then 

wscript.echo "There are no records to retrieve; Check that you have the correct job number."

Else 

'if there are records then loop through the fields 

Do While NOT Recordset.Eof   

 

field_can = Recordset("Sales_Count_CAN")
field_mgi = Recordset("Sales_Count_MGI")
 

if field_can <> "" then

wscript.echo field_can
wscript.echo field_mgi

end if

 

Recordset.MoveNext     

Loop

End If

 

'close the connection 

Recordset.Close
Set Recordset=nothing

oConn.Close
Set oConn=nothing

End Sub