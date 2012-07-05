<%@ CODEPAGE=65001 %>
<%Response.Charset="UTF-8"%>
<!--#include file="conn.asp"-->
<!--#include file="news.asp"-->
<%
newsid=request.QueryString("newsid")
if not isnumeric(newsid) then response.Redirect(install&"fail.asp")
set ns=new news
ns.getnewsinfo(newsid)
ns.updatecount(newsid)
response.Write "document.writeln('"&ns.rs("readcount")&"')"
%>