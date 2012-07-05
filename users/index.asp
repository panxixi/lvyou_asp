<!--#include file="../inc/utf.asp"-->
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/s.asp"-->
<!--#include file="../inc/users.asp"-->
<!--#include file="../inc/Baoming.asp"-->
<!--#include file="../inc/category.asp"-->
<!--#include file="../inc/sys.asp"-->
<!--#include file="../inc/tfunction.asp"-->
<!--#include file="check.asp"-->
<%
	dim body
	if not isBindQQ() then
		'body = replace(userIndex_temp,"{旅游:提示语句}","<script>alert('驴窝最新添加了绑定QQ功能!您的QQ还未在本站绑定过情使用本站的一键QQ绑定功能吧！');location.href='/txapi/QQbind.asp?apiType=qq';</script>")
		'die "<script>alert('驴窝最新添加了绑定QQ功能!您的QQ还未在本站绑定过,请使用本站的一键QQ绑定功能吧！');location.href='/txapi/QQbind.asp?apiType=qq';</script>"
		body = replace(userIndex_temp,"{旅游:提示语句}","<a href='/txapi/QQbind.asp?apiType=qq'><img src='/images/qq_bind_small.gif' /></a>")
	else
		body = replace(userIndex_temp,"{旅游:提示语句}","")
	end if
	echo body
%>