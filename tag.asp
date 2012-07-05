<!--#include file="inc/utf.asp"-->
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/s.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>TAG页面</title>
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
<table width="600" border="0" align="center" bgcolor="#666666" cellpadding="5" cellspacing="1">
  <tr bgcolor="#FFFFFF">
    <td><span class="STYLE1">TAG相关内容：</span></td>
  </tr>
<%
tag_id=request.QueryString("id")
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
sql="select * from news where nkeyword like '%{"&tag_id&"}%'"
rs.cursortype=3
rs.open(sql)
%>
<%
    do while not rs.eof
  %>
    <tr bgcolor="#FFFFFF">
  <%
    if html=0 then
  %>
    <td><a href='<%=install%>view.asp?newsid=<%=rs("newsid")%>' target="_blank"><%=rs("ntitle")%></a></td>
  <%
	elseif html=1 then
  %>
    <td><a href='<%=install%><%=rs("cid")%>/view-<%=rs("newsid")%>.html' target="_blank"><%=rs("ntitle")%></a></td>
  <%
  end If
  %>
  </tr>
  <%
  rs.movenext
  loop
  %>
</table>
</body>
</html>
