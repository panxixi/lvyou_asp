<!--#include file="inc/utf.asp"-->
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/s.asp"-->
<!--#include file="inc/users.asp"-->
<!--#include file="inc/news.asp"-->
<!--#include file="inc/category.asp"-->
<!--#include file="inc/carLocation.asp"-->
<!--#include file="inc/HuoDong.asp"-->
<!--#include file="inc/baoming.asp"-->
<!--#include file="inc/sys.asp"-->
<!--#include file="inc/tfunction.asp"-->
<%
huoDongId = Request.QueryString("huoDongId")
echo huoDong_temp()
%>