<!--#include file="inc/utf.asp"-->
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/s.asp"-->
<!--#include file="inc/news.asp"-->
<!--#include file="inc/category.asp"-->
<!--#include file="inc/sys.asp"-->
<!--#include file="inc/tfunction.asp"-->
<!--#include file="admin/md5.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=<%=q_Charset%>" />
<title>无标题文档</title>
<style type="text/css">
<!--
body{
	font-size:12px;
	text-decoration:none;
	background-color:#FFF;
	margin:0px;
}
.clear {
	clear: both;
}
a{
	color:#999999;
	text-decoration:none;
	font-size:12px;
}
a:hover{
	text-decoration:underline;
	color: #FF6600;
}
td{
	line-height:120%
}
td{
	color:#999999;
	text-decoration:none;
	font-size:12px;
}
.inp {
	background-color:#FFF;
	PADDING-RIGHT: 2px;
	PADDING-LEFT: 2px;
	PADDING-BOTTOM: 1px;
	WIDTH: 100px;
	PADDING-TOP: 1px;
	border: 1px solid #404040;
	color: #FFFFFF;
}
.admin_input_active {
	background-color:#EFF5FB;
	PADDING-RIGHT: 2px;
	PADDING-LEFT: 2px;
	PADDING-BOTTOM: 1px;
	WIDTH: 100px;
	PADDING-TOP: 1px;
	border: 1px solid #6392C8;
	color: #FFFFFF;
}
body,td,th {
	color: #000;
}
-->
</style>

</head>
<body>
<%
If session("username")="" then
%>
<table cellspacing="1" cellpadding="5" border="0" width="178">
						<form action="login.asp?act=login" method="post" name="form1" onSubmit="return form_onsubmit()">
						<tr>
							<td width="40%" align="right">用户：</td>
							<td width="60%"><input name="username" value="" class="inp" onfocus="this.className='admin_input_active';" onblur="this.className='inp'" /></td>
						</tr>
						<tr>
							<td align="right">密码：</td>
							<td><input type="password" name="password" class="inp"  onFocus="this.className='admin_input_active';" onBlur="this.className='inp'" /></td>
						</tr>
						<tr>
							<td valign="top" align="center" colspan="2">
							<input type="submit" value=" 登 录 " style="border:solid 1px #3C3C3C; background-color:#cccccc; color:#000000; padding-top:2px;">　
							<input type="button" value=" 注 册 " onclick="window.open('../reg.asp','_parent')" style="border:solid 1px #3C3C3C; background-color:#cccccc; color:#000000; padding-top:2px;"></td>
						</tr>
                        <tr>
						  <td valign="top" align="center" colspan="2">&nbsp;</td>
						</tr>
                        </form>
					</table>

<%
					else
					%>
					<table cellspacing="1" cellpadding="5" border="0" width="178">
						<tr>
							<td align="right">您好！<b style="color:#FF6600"><%=session("username")%></b></td>
						</tr>
						<tr>
							<td align="right">欢迎您的光临！<a href="login.asp?act=logout"><font color="#91BE19">[ 安全退出 ]</font></a></td>
						</tr>
						<tr>
							<td valign="top" align="center">感谢您注册本网站的会员，希望我们的产品或信息能为您带来帮助！</td>
						</tr>
                        <tr>
						  <td valign="top" align="center">&nbsp;</td>
						</tr>
					</table>
					<%
					End if
					%>
</body>
</html>
<%
act=request.querystring("act")
If act="login" then
siteout()
username=trim(request.Form("username"))
userpwd=md5(trim(request.Form("password")))
if username="" then
session("username")=""
%>
<SCRIPT LANGUAGE="javascript">
window.alert('请输入用户名!');window.self.location.href="login.asp"
</SCRIPT>
<%
else
if userpwd="" then
session("username")=""
%>
<SCRIPT LANGUAGE="javascript">
window.alert('请输入密码!');window.self.location.href="login.asp"
</SCRIPT>
<%
else
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql="select * from users where username='"&username&"' and userpwd='"&userpwd&"'"
rs.open sql
if rs.eof or rs.bof then
session("username")=""
%>
<SCRIPT LANGUAGE="javascript">
window.alert('用户名或者密码错误!');window.self.location.href="login.asp"
</SCRIPT>
<%
else
'把用户名密码放进session
session("username")=username
session("userpwd")=userpwd
session("id")=rs("id")
response.write "<script>alert('登陆成功');window.location.href='/login.asp';</script>"
end if
end if
end If
ElseIf act="logout" Then
session("username")=""
session("userpwd")=""
response.write "<script>alert('安全退出！');window.location.href='/login.asp';</script>"
end if
%>