<!--#include file="inc/utf.asp"-->
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/s.asp"-->
<!--#include file="inc/users.asp"-->
<!--#include file="inc/news.asp"-->
<!--#include file="inc/carlocation.asp"-->
<!--#include file="inc/category.asp"-->
<!--#include file="inc/HuoDong.asp"-->
<!--#include file="inc/Baoming.asp"-->
<!--#include file="inc/sys.asp"-->
<!--#include file="inc/tfunction.asp"-->
<%
'Dim huoDongId,Action,userName,UserId,huoDongJifen
huoDongId = Clng(Request.QueryString("huoDongId"))
Action = Request.QueryString("Action")
userName = Replace(Request.Cookies(CookieName)("UserName"),"'","''")
password = Replace(Request.Cookies(CookieName)("password"),"'","''")
If huoDongId = "" or huoDongId = 0 Then ToUrl("?Action=Step1")
set h = new HuoDong
h.gethuoDong(huoDongId)
if not h.rs.eof then 
	huoDongJifen = h.rs("huoDongJifen")
	huoDongShengyu = h.rs("huoDongShengyu")
end if
h.rs.close
set h.rs = nothing
set h = nothing


	
	
Select Case Action
	Case "Step1"
		'checkLoginBool()为True则用户已经注册！
		if checkLoginBool(userName,password) then
			echo isBindQQBackContent(Baoming_temp("Step2"))
		else
			echo Baoming_temp("Step1LoginReg")
		end if
	Case "Step2"
		echo isBindQQBackContent(Baoming_temp("Step2"))
	Case "Step3"
		huodongZhuangtai = Replace(Request.Form("huodongZhuangtai"),"'","''")
		baoMingHuoDong = Replace(Request.Form("baoMingHuoDong"),"'","''")
		baoMingHuoDongId = cint(Replace(Request.Form("baoMingHuoDongId"),"'","''"))
		baoMingUserName = Replace(Request.Form("baoMingUserName"),"'","''")
		baoMingRealName = Replace(Request.Form("baoMingRealName"),"'","''")
		baoMingRenshu = Replace(Request.Form("baoMingRenshu"),"'","''")
		baoMingPhone = Replace(Request.Form("baoMingPhone"),"'","''")
		baomingLiuyan = Replace(Request.Form("baomingLiuyan"),"'","''")
		baoMingIdcard = Replace(Request.Form("baoMingIdcard"),"'","''")
		baoMingUseJifen = cint(Request.Form("baoMingUseJifen"))
		location = Replace(Request.Form("location"),"'","''")
		isBaoxiana = Replace(Request.Form("isBaoxian"),"'","''")
		baoxianList = Replace(Request.Form("baoxianList"),"'","''")
		baoMingJifen = Request.Form("baomingjifen")
		baoMingZhuangtai = 0
		if instr(huodongZhuangtai,"报名已满") > 0 then die "活动报名已满，不能报名，下次吧! <a href='javascript:history.go(-1);'>返回</a>"
		if instr(huodongZhuangtai,"活动结束") > 0 then die "活动已经结束，不能报名，下次吧! <a href='javascript:history.go(-1);'>返回</a>"
		if baoMingRenshu = "" then die "请填写人数! <a href='javascript:history.go(-1);'>返回</a>"
		'需要判断当前活动剩余人数是否还有！
		if baoMingUseJifen = 1 then
			if baoMingJifen < 30 then
				die "您的积分不够不能使用积分支付! <a href='javascript:history.go(-1);'>返回</a>"
			end if
			baoMingJifen = 0
		else
			baoMingJifen = baoMingRenshu * huoDongJifen
		end if
		if baomingLiuyan = "" then
			baomingLiuyan = "无"
		end if
		set bm = new Baoming
        bm.baoMingHuoDong = baoMingHuoDong
        bm.baoMingHuoDongId = baoMingHuoDongId
        bm.baoMingUserName = baoMingUserName
        bm.baoMingRealName = baoMingRealName
        bm.baomingLiuyan = baomingLiuyan
        bm.baoMingRenshu = baoMingRenshu
        bm.baoMingPhone = baoMingPhone
        bm.baoMingIdcard = baoMingIdcard
        bm.baoMingZhuangtai = baoMingZhuangtai
        bm.baoMingAddTime = now()
        bm.baoMingJifen = baoMingJifen
        bm.baoMingUseJifen = baoMingUseJifen
        bm.location = location
        bm.isBaoxian = isBaoxiana
        bm.baoxianList = baoxianList
        bm.baoMingRen = Request.cookies(CookieName)("UserName")
        shengyu = huoDongShengyu - baoMingRenshu
		if shengyu > 0 then
	        bm.addbaoMing()
	    end if
	    if shengyu = 0 then
	    	bm.addbaoMing()
	    	bm.updateHuodongZhuangtai()
	    end if
	    if shengyu < 0 then
	    	if huoDongShengyu=0 then
	    		aa = "已经没有座位了！"
	    	else
	    		aa = "只剩下"&huoDongShengyu&"个座位了！"
	    	end if
	    	die "报名人数超过现有剩余人数，现在"&aa&" <a href='javascript:history.go(-1);'>返回</a>"
	    end if
		echo isBindQQBackContent(Baoming_temp("Step3"))
	Case "LoginAction"
		userName = Request.Form("userName")
		password = Request.Form("password")
		set u=new Users
		u.siteout()
		if userName <> "" or password <> "" then
			u.getuser(userName)
			if not u.rs.eof then
				if u.rs("userName")=userName and u.rs("password")=php_md5(password) then
					'login = uc_user_login(userName,password,"0","0","","")
					Response.cookies(CookieName)("UserName") = userName
					Response.cookies(CookieName)("password") = php_md5(password)
					Response.cookies(CookieName)("isUserLogin") = "True"
					isUserLogin = true
					Tourl("./Baoming.asp?Action=Step2&huoDongId="&huoDongId)
				else
					Tourl("./Baoming.asp?Action=Step1&huoDongId="&huoDongId)
				end if
			else
				Tourl("./Baoming.asp?Action=Step1&huoDongId="&huoDongId)
			end if
		else
			die "抱歉，必须把表格填写完整才能继续注册! <a href='javascript:history.go(-1);'>返回</a>"
		end if
	Case "RegAction"
		userName = Request.Form("userName")
		RealName = Request.Form("RealName")
		password = Request.Form("password")
		Email = Request.Form("Email")
		sex = Request.Form("sex")
		phone = Request.Form("phone")
		idcard = Request.Form("idcard")
		set u=new Users
		if userName <> "" or RealName <> "" or password <> "" or sex <> "" or phone <> "" or idcard <> "" then
			'register = uc_user_register(userName,password,Email,"","")
			Response.cookies(CookieName)("UserName") = userName
			Response.cookies(CookieName)("password") = php_md5(password)
			Response.cookies(CookieName)("isUserLogin") = "True"
			'u.siteout()
			'If register > "0" Then
	    		'echo "注册成功" 
			'ElseIf register = "-1" Then 
			'    die "用户名不合法 <a href='javascript:history.go(-1);'>返回</a>" 
			'ElseIf register = "-2" Then 
			'    die "包含要允许注册的词语 <a href='javascript:history.go(-1);'>返回</a>" 
			'ElseIf register = "-3" Then 
			'    die "用户名已经存在 <a href='javascript:history.go(-1);'>返回</a>" 
			'ElseIf register = "-4" Then 
			'    die "Email 格式有误 <a href='javascript:history.go(-1);'>返回</a>" 
			'ElseIf register = "-5" Then 
			'    die "Email 不允许注册 <a href='javascript:history.go(-1);'>返回</a>" 
			'ElseIf register = "-6" Then 
			'    die "该 Email 已经被注册 <a href='javascript:history.go(-1);'>返回</a>" 
			'ElseIf register = "-7" Then 
			'    die "注册信息填写不全 <a href='javascript:history.go(-1);'>返回</a>" 
			'Else
			'    die "未定义"
			'End If
			u.userName = userName
			u.RealName = RealName
			u.password = php_md5(password)
			u.sex = sex
			u.phone = phone
			u.idcard = idcard
			u.jifen = ChushiJifen
			u.joinTime = now()
			u.adduser()

			isUserLogin = true
			Tourl("./Baoming.asp?Action=Step2&huoDongId="&huoDongId)
		end if
	Case Else
		Tourl("./Baoming.asp?Action=Step1&huoDongId="&huoDongId)
End Select
%>