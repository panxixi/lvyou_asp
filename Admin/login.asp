<!--#include file="../inc/utf.asp"-->
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/sys.asp"-->
<!--#include file="../inc/s.asp"-->
<!--#include file="md5.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=<%=q_Charset%>" />
</head>

<body>
<%
set cn=new config

admin_name=trim(request.Form("admin_name"))
admin_password=trim(request.Form("admin_password"))
admin_yzm=trim(request.Form("admin_yzm"))
admin_password=md5(admin_password)
if admin_yzm=yzcode then
if admin_name="" then
session("admin_name")=""
%>
<SCRIPT LANGUAGE="javascript">
window.alert('请输入用户名!');window.self.location.href="index.asp"
</SCRIPT>
<%
else
if admin_password="" then
session("admin_name")=""
%>
<SCRIPT LANGUAGE="javascript">
window.alert('请输入密码!');window.self.location.href="index.asp"
</SCRIPT>
<%
else
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql="select * from admin where admin_name='"&admin_name&"' and admin_password='"&admin_password&"'"
rs.open sql
if rs.eof or rs.bof then
session("admin_name")=""
%>
<SCRIPT LANGUAGE="javascript">
window.alert('用户名或者密码错误!');window.self.location.href="index.asp"
</SCRIPT>
<%
else
'把用户名密码放进session
session("admin_name")=admin_name
session("admin_password")=admin_password
session("id")=rs("id")
%>
<SCRIPT LANGUAGE="javascript">
window.self.location.href="manage.asp"
</SCRIPT>
<%
end if
end if
end if
else
%>
<SCRIPT LANGUAGE="javascript">
window.alert('验证失败!');window.self.location.href="index.asp"
</SCRIPT>
<%
end if
%>
</body>
</html>