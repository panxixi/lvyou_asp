<!--#include file="utf.asp"-->
<!--#include file="conn.asp"-->
<!--#include file="s.asp"-->
<!--#include file="review.asp"-->
<!--#include file="tfunction.asp"-->
<%
If Trim(request.Form("code")) <>CStr(Session("GetCode")) then
  Response.Write("<script>alert('验证码有误，重新输入！');history.back();</script>")
  Response.End
End If
Set rv=new review
rv.poster=nohtml(request.Form("poster"))
rv.dcontent=nohtml(request.Form("dcontent"))
rv.newsid=nohtml(Trim(request.Form("newsid")))
rv.posttime=Now()
rv.did_add()
xs_err("<font color='red'>回复失败，请联系管理员！</font>")
jump=zx_url("回复成功",Trim(Request.ServerVariables("HTTP_REFERER")))
%>

