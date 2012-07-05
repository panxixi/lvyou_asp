<!--#include file="inc/utf.asp"-->
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/s.asp"-->
<!--#include file="inc/users.asp"-->
<!--#include file="inc/category.asp"-->
<!--#include file="inc/HuoDong.asp"-->
<!--#include file="inc/Baoming.asp"-->
<!--#include file="inc/sys.asp"-->
<!--#include file="inc/tfunction.asp"-->
<%
huoDongId = Clng(Request.QueryString("huoDongId"))
If huoDongId = "" or huoDongId = 0 Then ToUrl("?Action=Step1")
echo isBindQQBackContent(userBaoming_str("templist/mingdan.html"))
%>