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
Select Case Action
	Case "Head"
		echo userIndex_str("templist/users/UserOptionHead.html")
	Case "Password"
		echo userIndex_str("templist/users/UserOptionPassword.html")
	Case "EditUserOption"
		call EditUserOption()
	Case "EditPassword"
		call EditPassword()
	Case Else
		echo userIndex_str("templist/users/UserOption.html")
end Select

Sub EditUserOption()
	Set u = new Users
	RealName = Replace(Request.Form("RealName"),"'","''")
	Sex = Replace(Request.Form("Sex"),"'","''")
	Phone = Replace(Request.Form("phone"),"'","''")
	Idcard = Replace(Request.Form("Idcard"),"'","''")
	if RealName = "" or Sex = "" or Phone = "" or Idcard = "" then die "<script>alert('资料没有填写完整！');history.go(-1);</script>"
	u.userName = userName
	u.RealName = RealName
	u.sex = Sex
	u.phone = Phone
	u.idcard = Idcard
	u.updateUserOption()
	die "<script>alert('修改成功！');location.href='./UserOption.asp';</script>"
	set u = nothing
end Sub

Sub EditPassword()
	Set u = new Users
	oldPassword = Replace(Request.Form("oldPassword"),"'","''")
	newPassword = Replace(Request.Form("newPassword"),"'","''")
	Call u.getuser(userName)
	if not u.rs.eof then 
		if MD5(oldPassword) = u.rs("password") then
			'uedit = uc_user_edit(userName,oldPassword,newPassword,"","1","","")
				u.password = MD5(newPassword)
			    call u.edituserPass(userName)
			    die "<script>alert('密码修改成功！');location.href='/';</script>"
		else
			die "<script>alert('原密码错误！');history.go(-1);</script>"
		end if
	end if
	set u = nothing
end Sub
%>