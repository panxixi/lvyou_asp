<!--#include file="const.asp"-->
<%
'On Error Resume Next
If data_type="mssql" Then
set conn=server.createobject("adodb.Connection")
conn.Open"Driver={SQL Server};Server="&mssql_ip&";UID="&mssql_user&";PWD="&mssql_pwd&";database="&mssql_data_name&";"
Else
Set conn=Server.CreateObject("ADODB.Connection")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath(install&DBPath) 
End If
  	If Err Then
		
		err.descrption
		Set Conn = Nothing
		Response.Write "数据库连接出错"
		Response.End
	End If
%> 

