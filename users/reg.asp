 <!--#include file="../inc/utf.asp"-->
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/users.asp"-->
<!--#include file="../inc/category.asp"-->
<!--#include file="../inc/sys.asp"-->
<!--#include file="../inc/tfunction.asp"-->
<%
Dim Action,login,register
Action = Request.QueryString("Action")
Select Case Action
	Case "FromBBsAct"
		RealName = Replace(Request.Form("RealName"),"'","''")
		Email = Replace(Request.Form("Email"),"'","''")
		sex = Replace(Request.Form("sex"),"'","''")
		phone = Replace(Request.Form("phone"),"'","''")
		idcard = Replace(Request.Form("idcard"),"'","''")
		set u=new Users
		if userName <> "" and RealName <> "" and password <> "" and sex <> "" and phone <> "" and idcard <> "" then
			u.siteout()
			u.userName = userName
			u.RealName = RealName
			u.password = password
			u.sex = sex
			u.phone = phone
			u.idcard = idcard
			u.jifen = "1"
			u.joinTime = now()
			u.adduser()
			isUserLogin = true
			Tourl("./Index.asp")
		end if
	Case "RegAction"
		userName = Replace(Request.Form("userName"),"'","''")
		RealName = Replace(Request.Form("RealName"),"'","''")
		password = Replace(Request.Form("password"),"'","''")
		Email = Replace(Request.Form("Email"),"'","''")
		sex = Replace(Request.Form("sex"),"'","''")
		phone = Replace(Request.Form("phone"),"'","''")
		idcard = Replace(Request.Form("idcard"),"'","''")
		set u=new Users
		if userName <> "" and RealName <> "" and password <> "" and sex <> "" and phone <> "" and idcard <> "" then
			u.getuser(userName)
			if not u.rs.eof then die "<script>alert('-_-! Mingzi ChongFu!');history.go(-1);</script>"
			'register = uc_user_register(userName,password,Email,"","")
			Response.cookies(CookieName)("UserName") = userName
			Response.cookies(CookieName)("password") = MD5(password)
			Response.cookies(CookieName)("Uid") = register
			Response.cookies(CookieName)("isUserLogin") = "True"
			'u.siteout()
			'If register > "0" Then
	    	'	echo "注册成功" 
			'ElseIf register = "-1" Then 
			'    echo "用户名不合法" 
			'ElseIf register = "-2" Then 
			'    echo "包含要允许注册的词语" 
			'ElseIf register = "-3" Then 
			'    echo "用户名已经存在" 
			'ElseIf register = "-4" Then 
			'    echo "Email 格式有误" 
			'ElseIf register = "-5" Then 
			'    echo "Email 不允许注册" 
			'ElseIf register = "-6" Then 
			'    echo "该 Email 已经被注册" 
			'ElseIf register = "-7" Then 
			'    echo "注册信息填写不全" 
			'Else
			'    echo "未定义" 
			'End If
			u.userName = userName
			u.RealName = RealName
			u.password = MD5(password)
			u.sex = sex
			u.phone = phone
			u.idcard = idcard
			u.jifen = ChushiJifen
			u.joinTime = now()
			u.adduser()

			isUserLogin = true
			Tourl("./Index.asp")
		else
			echo "抱歉，必须把表格填写完整才能继续注册! <a href='javascript:history.go(-1);'>返回</a>"
		end if
'		u.rs.close
		set u.rs = nothing
		set u = nothing
	Case Else
		echo userReg_temp
end Select
%>