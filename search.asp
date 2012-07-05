<!--#include file="inc/utf.asp"-->
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/s.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>QCMS搜索页</title>
<style type="text/css">
<!--
.STYLE1 {
	color: #FF0000;
	font-weight: bold;
}
body,td,th {
	font-size: 12px;
}
a:link {
	text-decoration: none;
	color: #333333;
}
a:visited {
	text-decoration: none;
	color: #333333;
}
a:hover {
	text-decoration: none;
	color: #FF6600;
}
a:active {
	text-decoration: none;
	color: #333333;
}
-->
</style>
</head>

<body>
<table width="600" border="0" align="center">
  <tr>
    <td><form id="form1" name="form1" method="post" action="search.asp?do=1">
    关键字：<input name="keyword" type="text" id="keyword" size="30" />
    <label>
    <input type="submit" name="button" id="button" value="搜 索" />
    </label>
    </form>
    </td>
  </tr>
  <tr>
    <td></td>
  </tr>
</table>
<%
caozuo=trim(request.QueryString("do"))
if caozuo="1" then
if not isnumeric(caozuo) then response.Redirect(install&"fail.asp")
%>
<table width="600" border="0" align="center" bgcolor="#666666" cellpadding="5" cellspacing="1">
  <tr bgcolor="#FFFFFF">
    <td><span class="STYLE1">搜索结果：</span></td>
  </tr>
  <%
keyword=request.Form("keyword")
If keyword="" Then
response.write "请输入搜索关键字"
else
fenlei=request.Form("fenlei")
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
sql="select * from news where ntitle like '%"&keyword&"%'"
rs.cursortype=3
rs.open(sql)
%>
  <%
    do while not rs.eof
  %>
    <tr bgcolor="#FFFFFF">
   <td><a href='view.asp?newsid=<%=rs("newsid")%>' target="_blank"><%=rs("ntitle")%></a></td>
  </tr>
  <%
  rs.movenext
  loop
  end If
  End if
  %>
</table>
</body>
</html>
