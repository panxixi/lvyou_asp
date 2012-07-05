<!--#include file="inc/utf.asp"-->
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/s.asp"-->
<!--#include file="inc/review.asp"-->
<!--#include file="inc/tfunction.asp"-->
<%
Set rv=new review
id=Trim(request.querystring("id"))
act=request.querystring("act")
rv.did_get(id)
If rv.rs.eof Then
jj="<tr><td>暂时无任何回复</td></tr>"
else
rv.rs.PageSize = 2
Page = CLng(Request("Page"))
If Page < 1 Or page="" Then Page = 1
If Page > rv.rs.PageCount Then Page = rv.rs.PageCount
rv.rs.AbsolutePage = Page
for i=1 to rv.rs.PageSize
if rv.rs.EOF then Exit For
kk="<tr><td>用户："&rv.rs("poster")&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;发布时间:"&rv.rs("posttime")&"</td></tr><tr><td>"&rv.rs("dcontent")&"</td></tr>"
jj=jj+kk
rv.rs.movenext()
Next
if page=1 then
re_page="<tr><td>首页&nbsp;上一页&nbsp;第"&page&"页&nbsp;<a href='review.asp?id="&id&"&page="&page+1&"'>下一页</a>&nbsp;<a href='review.asp?id="&id&"&page="&rv.rs.pagecount&"'>尾页</a></td></tr>"
elseif page=rv.rs.pagecount then
re_page="<tr><td><a href='review.asp?id="&id&"&page=1'>首页</a>&nbsp;<a href='review.asp?id="&id&"&page="&page-1&"'>上一页</a>&nbsp;第"&page&"页&nbsp;下一页&nbsp;尾页</td></tr>"
else
re_page="<tr><td><a href='review.asp?id="&id&"&page=1'>首页</a>&nbsp;<a href='review.asp?id="&id&"&page="&page-1&"'>上一页</a>&nbsp;第"&page&"页&nbsp;<a href='review.asp?id="&id&"&page="&page+1&"'>下一页</a>&nbsp;<a href='review.asp?id="&id&"&page="&rv.rs.pagecount&"'>尾页</a></td></tr>"
end If
End if
If act="js" then
ll=dm_to_js("<table width='100%' border='1' cellspacing='0' >"&jj&"</table>")
response.write "document.writeln('"&ll&"');"
Else
ll="<table>"&jj&re_page&"</table>"
response.write ll
End If

%>
