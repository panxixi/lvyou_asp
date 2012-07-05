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
baomingId = Request.QueryString("baomingId")
huodongId = Request.QueryString("huodongId")
Select Case Action
	Case "JiaJianRen"
		echo userBaoming_str("templist/users/JiaJianRen.html")
	Case "JiaRen"
		JiaRen()
	Case "JianRen"
		JianRen()
	Case Else
		echo userIndex_str("templist/users/ViewHuodongs.html")
end Select

Sub JiaRen()
	Renshu = cint(Request.Form("Renshu"))
	set h = new HuoDong
	set b = new Baoming
	call h.gethuoDong(huoDongId)
	call b.getUserBaomingSingle(userName,baoMingId,huodongId)
	if not h.rs.eof then
		if h.rs("huoDongShengyu") >= Renshu then
			'die h.rs("huoDongJiFen")&"x"&Renshu + b.rs("baoMingRenshu")
			bencijifen = h.rs("huoDongJiFen") * (Renshu + b.rs("baoMingRenshu"))
			'bencijifen = bencijifen + b.rs("baoMingJifen")
			call b.updateJiaren(Renshu,baomingId)
			call h.JianShengyu(Renshu,huoDongId)
			call b.updateJifen(bencijifen,baomingId)
			'die "<script>alert('恭喜您！积分累加成功，加人成功，祝您玩的高兴！');location.href='/Mingdan.asp?huodongId="&huodongId&"';</script>"
			tourl("/Users/Huodong.asp?Action=Ybao")
		else
			'die "<script>alert('对不起！本活动已经没有剩余座位了！');history.go(-1);</script>"
			tourl("/Users/Huodong.asp?Action=Ybao")
		end if
	end if
	h.rs.close
	b.rs.close
	set b = nothing
	set h = nothing
End Sub

Sub JianRen()
	Renshu = cint(Request.Form("Renshu"))
	set h = new HuoDong
	set b = new Baoming
	call b.getUserBaomingSingle(userName,baoMingId,huodongId)
	call h.gethuoDong(huoDongId)
	if not b.rs.eof then
		if cint(b.rs("baoMingRenshu"))>=Renshu then
			bencijifen = h.rs("huoDongJiFen") * (b.rs("baoMingRenshu") - Renshu)
			'bencijifen = b.rs("baoMingJifen")-bencijifen
			call b.updateJianren(Renshu,baomingId)
			call h.jiaShengyu(Renshu,huoDongId)
			call b.updateJianJifen(bencijifen,baomingId)
			'die "<script>alert('积分按照人数比例已经相应扣除！感谢您对驴窝的支持，我们将做的更好！');location.href='/Mingdan.asp?huodongId="&huodongId&"';</script>"
			tourl("/Users/Huodong.asp?Action=Ybao")
		else
			'die "<script>alert('您减少的人数应该比你的报名人数小于或者等于才可以取消！');history.go(-1);</script>"
			tourl("/Users/Huodong.asp?Action=Ybao")
		end if
	end if
End Sub
%>