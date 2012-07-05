 <!--#include file="../inc/utf.asp"-->
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/users.asp"-->
<!--#include file="../inc/category.asp"-->
<!--#include file="../inc/sys.asp"-->
<!--#include file="../inc/huodong.asp"-->
<!--#include file="../inc/tfunction.asp"-->
<%
Dim Action,login,register
Action = Request.QueryString("Action")
Select Case Action
	Case "BindQQ"
		userName = Replace(Request.Form("userName"),"'","''")
		password = Request.Form("password")
		set u=new Users
		u.siteout()
		if userName <> "" or password <> "" then
				u.getuser(userName)
				if not u.rs.eof then
					if u.rs("password") <> MD5(password) then toBackurl()
					if u.rs("userName")=userName and u.rs("password")=MD5(password) then
						qqid = session("qqId")
						u.bindQQ u.rs("userName"),qqid
						Response.cookies(CookieName)("UserName") = userName
						Response.cookies(CookieName)("password") = MD5(password)
						'Response.cookies(CookieName)("Uid") = register
						Response.cookies(CookieName)("isUserLogin") = "True"
						isUserLogin = true
						Tourl("./Index.asp")
					else
						Tobackurl()
					end if
				else
					Tobackurl()
				end if
			'end if
			'if login(0) = "-1" then die "用户不存在，或者被删除 <a href='javascript:history.go(-1);'>返回</a>"
			'if login(0) = "-2" then die "论坛密码错误！ <a href='javascript:history.go(-1);'>返回</a>"
		else
			die "抱歉，必须把表格填写完整才能继续! <a href='javascript:history.go(-1);'>返回</a>"
		end if
	Case "QQreg"
		userName = Replace(Request.Form("userName"),"'","''")
		RealName = Replace(Request.Form("RealName"),"'","''")
		password = Replace(Request.Form("password"),"'","''")
		Email = Replace(Request.Form("Email"),"'","''")
		sex = Replace(Request.Form("sex"),"'","''")
		phone = Replace(Request.Form("phone"),"'","''")
		idcard = Replace(Request.Form("idcard"),"'","''")
		qqid = session("qqId")
		set u=new Users
		if userName <> "" and RealName <> "" and password <> "" and sex <> "" and phone <> "" and idcard <> "" then
			u.getuser(userName)
			if not u.rs.eof then die "<script>alert('-_-! Mingzi ChongFu!');history.go(-1);</script>"
			'register = uc_user_register(userName,password,Email,"","")
			Response.cookies(CookieName)("UserName") = userName
			Response.cookies(CookieName)("password") = MD5(password)
			Response.cookies(CookieName)("Uid") = register
			Response.cookies(CookieName)("isUserLogin") = "True"
			u.siteout()
			u.userName = userName
			u.RealName = RealName
			u.password = MD5(password)
			u.sex = sex
			u.phone = phone
			u.idcard = idcard
			u.jifen = ChushiJifen
			u.joinTime = now()
			u.QQHASH = qqid
			u.addQQuser()

			isUserLogin = true
			Tourl("./Index.asp")
		else
			echo "抱歉，必须把表格填写完整才能继续注册! <a href='javascript:history.go(-1);'>返回</a>"
		end if
'		u.rs.close
		set u.rs = nothing
		set u = nothing
	Case Else
		'die session("QQtx")
		body = replace(userQQReg_temp,"{旅游:QQ头像}","<img src="""&session("QQtx")&">")
		body = replace(body,"{旅游:QQ昵称}",session("qqNick"))
		echo body
end Select
%>