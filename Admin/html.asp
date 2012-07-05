<!--#include file="../inc/utf.asp"-->
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/news.asp"-->
<!--#include file="../inc/users.asp"-->
<!--#include file="../inc/category.asp"-->
<!--#include file="../inc/huodong.asp"-->
<!--#include file="../inc/sys.asp"-->
<!--#include file="../inc/review.asp"-->
<!--#include file="../inc/tfunction.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="isadmin.asp"-->
<html>
<head>
<link rel=stylesheet href="styles/advanced/style.css" />
<meta http-equiv="Content-Type" content="text/html; charset=<%=q_Charset%>">
</head>
<body>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">复制模版</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<div class="table2">
<ul>
<%
Response.Buffer = true
Server.ScriptTimeOut=999
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
act=request.querystring("active")
'生成首页
If act="index" Then
If html<>1 Then
response.write "<li>未开启生成静态"
else
Call do_html(index_temp(),Server.MapPath(install&"index.html"),q_Charset)
response.write "<li>生成成功！"
End if
'生成分类页
ElseIf act="list" Then
If html<>1 Then
response.write "<li>未开启生成静态"
else
response.Write("<li>生成过程比较慢，请耐心等待……<br>")
Set ctx=new category
rs=ctx.getlistdir()
Do While Not ctx.rs.eof
MakeNewsDir(Server.MapPath("/Html"))
classid=ctx.rs("cid")
pnum=pagenum(install&ctx.rs("ctemp"),classid)
for k =1 to pnum
html_page=k
Call do_html(list_temp(classid),Server.MapPath("/Html/"&"list-"&k&".html"),q_Charset)
Next
Set fso=Server.CreateObject("Scripting.FileSystemObject")
fso.CopyFile Server.MapPath("/Html/"&"list-1.html"),Server.MapPath("/Html/"&"index.html")
fso.DeleteFile(Server.MapPath("/Html/"&"list-1.html"))
Set fso = nothing
Set Engine = nothing
response.write "<li>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;分类<font color='#ff0000'>"&classid&"</font>已经生成<br>"
response.Flush()
ctx.rs.movenext()
Loop
response.write "<li>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;已生成全部分类！"
End if
'生成内容页
ElseIf act="cont" Then
If html<>1 Then
response.write "<li>未开启生成静态"
else
response.Write("<li>生成过程比较慢，请耐心等待……<br>")
set nst=new news
rs=nst.allnews()
do while not nst.rs.eof
newsid=nst.rs("newsid")

Call do_html(view_temp(newsid),Server.MapPath(nst.rs("curl")&nst.rs("pinyin")&".html"),q_Charset)
response.write "<li>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;已经生成ID"&nst.rs("newsid")&"<br>"

nst.rs.movenext()
response.Flush()
Loop
response.write "<li>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;已生成全部内容！"
End if
'生成GOOGLE地图
ElseIf act="google" Then
Call do_html(read_temp(Server.MapPath(install&"templist/google_map.qcms")),Server.MapPath(install&"google_sitemap.xml"),q_Charset)
response.write "<li>生成Google地图成功！"
'生成BAIDU地图
ElseIf act="baidu" Then
Call do_html(read_temp(Server.MapPath(install&"templist/baidu_map.qcms")),Server.MapPath(install&"baidu_sitemap.xml"),q_Charset)
response.write "<li>生成Baidu地图成功！"
'生成总RSS
ElseIf act="rss" then
Call do_html(read_temp(Server.MapPath(install&"templist/rss.qcms")),Server.MapPath(install&"rss_sitemap.xml"),q_Charset)
response.write "<li>生成RSS地图成功！"
End if

%>
</ul>
</div>
</div>
<br>
<br>
<table width="100%" border="0" cellpadding="0" cellspacing="0" align="center" class="table">
  <tr>
    <td align="center" height="40">Powered By <a href="http://www.q-cms.cn"><%=Version%></a>&nbsp;&nbsp;<%=UCase(q_Charset)%>&nbsp;&nbsp;<%=UCase(data_type)%> &copy; 2009 All Rights Reserved  <div class="no_view" style="display:none;"><script src="http://s87.cnzz.com/stat.php?id=1707573&web_id=1707573&show=pic" language="JavaScript" charset="gb2312"></script></div>
</td>
</table>

</body>
</html>