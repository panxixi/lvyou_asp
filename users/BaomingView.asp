<!--#include file="../inc/utf.asp"-->
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/s.asp"-->
<!--#include file="../inc/users.asp"-->
<!--#include file="../inc/baoming.asp"-->
<!--#include file="../inc/huodong.asp"-->
<!--#include file="../inc/category.asp"-->
<!--#include file="../inc/sys.asp"-->
<!--#include file="../inc/tfunction.asp"-->
<!--#include file="check.asp"-->
<%
Dim Action,login
baoMingId = Request.QueryString("baoMingId")
userName = vbsUnEscape(Replace(Request.Cookies(CookieName)("UserName"),"'","''"))
echo userIndex_Baomingstr("templist/users/BaomingView.html")
%>