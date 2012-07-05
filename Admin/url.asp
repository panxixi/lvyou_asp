<!--#include file="../inc/utf.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/news.asp"-->
<!--#include file="../inc/sys.asp"-->
<!--#include file="../inc/category.asp"-->
<!--#include file="../inc/do.asp"-->
<!--#include file="isadmin.asp"-->
<html>
<head>
<link rel=stylesheet href="styles/advanced/style.css" />
<meta http-equiv="Content-Type" content="text/html; charset=<%=q_Charset%>">
</head>
<body>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">分类管理</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<div class="table2">
<ul>
<li>重置路径，请耐心等待……
<%
Response.Buffer = true
Server.ScriptTimeOut=999
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
act=request.querystring("active")
If act="class" then
Set ct=new category
rs=ct.no_diy_types()
Do While Not ct.rs.eof
kk=class_arclist(ct.rs("cid"))
ll=ct.rs("cid")
strsql="update category set curl='"&kk&"' where cid="&ll
conn.execute(strsql)
response.write "<li>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;重置分类路径：cid="&ct.rs("cid")&"<br>"
response.Flush()
ct.rs.movenext()
loop
ElseIf act="cont" then
Set ns=new news
rs=ns.allnews()
Do While Not ns.rs.eof
id=ns.rs("newsid")
kk=cont_html_url(id)
strsql="update news set pinyin='"&kk&"' where newsid="&ns.rs("newsid")
conn.execute(strsql)
response.write "<li>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;重置内容路径ID："&ns.rs("newsid")&"<br>"
response.Flush()
ns.rs.movenext()
loop
End if
response.write "<font color='#ff0000'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;重置内容路径完成</font>"
%>
</div>
<br />
<br />
<table width="100%" border="0" cellpadding="0" cellspacing="0" align="center" class="table">
  <tr>
     <td align="center" height="40">Powered By <a href="http://www.q-cms.cn"><%=Version%></a>&nbsp;&nbsp;<%=UCase(q_Charset)%>&nbsp;&nbsp;<%=UCase(data_type)%> &copy; 2009 All Rights Reserved  <div class="no_view" style="display:none;"><script src="http://s87.cnzz.com/stat.php?id=1707573&web_id=1707573&show=pic" language="JavaScript" charset="gb2312"></script></div>
</td>
</table>
</body>
</html>