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
Action = Request.QueryString("Action")
userName = vbsUnEscape(Replace(Request.Cookies(CookieName)("UserName"),"'","''"))
Select Case Action
	Case "ZXin"
		echo isBindQQBackContent(userTp_str("templist/users/ViewHuodongs.html"))
	Case "Ybao"
		echo isBindQQBackContent(userIndex_str("templist/users/YibaoHuodong.html"))
	Case "CGuo"
		echo isBindQQBackContent(userIndex_str("templist/users/CuoguoHuodong.html"))
	Case "JiaJianRen"
		baomingId = Request.QueryString("baomingId")
		huodongId = Request.QueryString("huodongId")
		echo isBindQQBackContent(userBaoming_str("templist/users/JiaJianRen.html"))
	Case Else
		echo isBindQQBackContent(userIndex_str("templist/users/ViewHuodongs.html"))
end Select
%>