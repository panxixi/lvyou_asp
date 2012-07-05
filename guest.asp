<!--#include file="inc/utf.asp"-->
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/s.asp"-->
<!--#include file="inc/news.asp"-->
<!--#include file="inc/category.asp"-->
<!--#include file="inc/sys.asp"-->
<!--#include file="inc/tfunction.asp"-->

<%
ly=request.QueryString("active")
if ly="do" Then
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=<%=q_Charset%>" />
<title>无标题文档</title>
</head>

<body>
<%
If Trim(request.Form("code")) <>CStr(Session("GetCode")) then
  Response.Write("<script>alert('验证码有误，重新输入！');history.back();</script>")
  Response.End
End If
'留言接收
set cn=new config
cn.title=nohtml(request.Form("title"))
cn.name=nohtml(request.Form("name"))
cn.tel=nohtml(request.Form("tel"))
cn.email=nohtml(request.Form("email"))
cn.qq=nohtml(request.Form("qq"))
cn.content=nohtml(request.Form("content"))
cn.times=now()
cn.ip=request.ServerVariables("REMOTE_ADDR")
If message="true" then
cn.addguest()
xs_err("<font color='red'>留言失败！</font>")
jump=zx_url("留言成功！","guest.asp")
Else
If session("username")="" then
jump=zx_url("只有登陆后才能留言！","guest.asp")
Else
cn.addguest()
xs_err("<font color='red'>留言失败！</font>")
End If
jump=zx_url("留言成功！","guest.asp")
End If
%>
</body>
</html>
<%
else
response.Write guest_temp()
end if

%>
