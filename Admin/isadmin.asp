<%
'取得用户名密码
admin_name=trim(session("admin_name"))
admin_password=trim(session("admin_password"))
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql="select * from admin where admin_name='"&admin_name&"' and admin_password='"&admin_password&"'"
rs.open sql
if rs.eof then
session("admin_name")=""
session("admin_pwd")=""
%>
<SCRIPT LANGUAGE="javascript">
window.alert('请登陆后台!');
</SCRIPT>
<%
response.Redirect("err.asp")
end if
%>