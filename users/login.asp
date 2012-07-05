<!--#include file="../inc/utf.asp"-->
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/s.asp"-->
<!--#include file="../inc/users.asp"-->
<!--#include file="../inc/category.asp"-->
<!--#include file="../inc/sys.asp"-->
<!--#include file="../inc/tfunction.asp"-->
<%
Dim Action,login
Action = Request.QueryString("Action")
Select Case Action
	Case "LoginAction"
		userName = Request.Form("userName")
		password = Request.Form("password")
		set u=new Users
		u.siteout()
		if userName <> "" or password <> "" then
		'login = uc_user_login(userName,password,"0","0","","")
			'if cint(login(0)) > 0 then
				u.getuser(userName)
				'if u.rs.eof then
				'	Response.cookies(CookieName)("Uid") = login(0)
				'	Response.cookies(CookieName)("UserName") = vbsEscape(userName)
				'	Response.cookies(CookieName)("password") = php_md5(password)
				'	Response.cookies(CookieName)("isUserLogin") = "True"
				'	echo "<script>alert('报名系统没有您的资料，需要您几分钟填写相关信息激活报名系统！')</script>"
				'	die userTp_str("templist/users/FromBBsReg.html")
				'end if
				if not u.rs.eof then
					if u.rs("password") <> MD5(password) then die "<script>alert('密码错误！请联系橡皮为您更改密码！');location.href='/';</script>"
					if u.rs("userName")=userName and u.rs("password")=MD5(password) then
						'Response.cookies(CookieName)("Uid") = login(0)
						Response.cookies(CookieName)("UserName") = userName
						Response.cookies(CookieName)("password") = MD5(password)
						Response.cookies(CookieName)("isUserLogin") = "True"
						isUserLogin = true
						Tourl("./Index.asp")
					else
						Tourl("./Reg.asp")
					end if
				else
					Tourl("./Reg.asp")
				end if
			'end if
			'if login(0) = "-1" then die "用户不存在，或者被删除 <a href='javascript:history.go(-1);'>返回</a>"
			'if login(0) = "-2" then die "论坛密码错误！ <a href='javascript:history.go(-1);'>返回</a>"
		else
			die "抱歉，必须把表格填写完整才能继续! <a href='javascript:history.go(-1);'>返回</a>"
		end if
	Case "Logout"
		Response.cookies(CookieName)("UserName") = ""
		Response.cookies(CookieName)("password") = ""
		Response.cookies(CookieName)("isUserLogin") = ""
		'echo uc_user_synlogout()
		Tourl("/Index.asp")
	Case Else
		echo userLogin_temp
end Select
%>