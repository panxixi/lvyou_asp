<!--#include file="../inc/utf.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/HuoDong.asp"-->
<!--#include file="../inc/baoming.asp"-->
<!--#include file="../inc/users.asp"-->
<!--#include file="../inc/sys.asp"-->
<!--#include file="../inc/category.asp"-->
<!--#include file="../inc/tfunction.asp"-->
<!--#include file="isadmin.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=<%=q_Charset%>" />
<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
}
-->
</style>
<%
	Action = Request.QueryString("Action")
	Select Case Action
		Case "ListView"
			Call ListView()
		Case "JiesuanAct"
			Call JiesuanAct()
		Case "AddAct"
			Call AddAct()
		Case "DelAct"
			Call DelAct()
		Case "EditZhuangtai"
			Call EditZhuangtai()
		Case Else
			Call SelectHuodong()
	End Select
	Public Sub JiesuanAct()
		set h=new HuoDong
		set u=new Users
		set b=new Baoming
		huodongId = Replace(Request.QueryString("huodongId"),"'","''")
		baomingId = Trim(Replace(Request.Form("baomingId"),"'","''"))
		baoMingUserName = split(Trim(Replace(Request.Form("baoMingUserName"),"'","''")),", ")
		baoMingRenshu = split(Trim(Replace(Request.Form("baoMingRenshu"),"'","''")),", ")
		baoMingJifen = split(Trim(Replace(Request.Form("baoMingJifen"),"'","''")),", ")
		baoMingUseJifen = split(Trim(Replace(Request.Form("baoMingUseJifen"),"'","''")),", ")
		baomingArr = Split(baomingId,", ")
		for i=0 to ubound(baoMingJifen)
			if baoMingUseJifen(i) = "0" then
				'echo "积分:"&baoMingJifen(i)&" 用户:"&baoMingUserName(i)&"<br>"
				call u.JiajifenByuserName(baoMingJifen(i),baoMingUserName(i))
				call b.updatejiesuanBaoming(baoMingRenshu(i),baomingArr(i))
			else
				call u.JianjifenByuserName(baoMingJifen(i),baoMingUserName(i))
				call b.updatejiesuanBaoming(baoMingRenshu(i),baomingArr(i))
			end if
		next
		'die "<br/>--over"
		call h.updateJiesuan(huodongId)
		set h = nothing
		set u = nothing
		set b = nothing
		echo "<script>alert('结算成功！');</script>"
		Tourl("Baoming.asp")
	End Sub
	
	Public Sub AddAct()
		set hd=new HuoDong
		hd.huoDongName = Replace(Request.Form("huoDongName"),"'","''")
		hd.bbsLink = Replace(Request.Form("bbsLink"),"'","''")
		hd.huoDongRenShu = Replace(Request.Form("huoDongRenShu"),"'","''")
		hd.huoDongShengyu = Replace(Request.Form("huoDongRenShu"),"'","''")
		hd.huoDongXingcheng = Replace(Request.Form("huoDongXingcheng"),"'","''")
		hd.huoDongxieyi = Replace(Request.Form("huoDongxieyi"),"'","''")
		hd.huoDongLingdui = Replace(Request.Form("huoDongLingdui"),"'","''")
		hd.huoDongZhuangtai = Replace(Request.Form("huoDongZhuangtai"),"'","''")
		hd.huoDongJifen = Replace(Request.Form("huoDongJifen"),"'","''")
		hd.huoDongChufaTime = Replace(Request.Form("huoDongChufaTime"),"'","''")
		hd.huoDongFanhuiTime = Replace(Request.Form("huoDongFanhuiTime"),"'","''")
		hd.huoDongInfo = Replace(Request.Form("huoDongInfo"),"'","''")
		hd.huoDongAddTime = now()
		aaa = Replace(Request.Form("nkeyword"),"'","''")
		aaa = Replace(aaa,"，",",")
		aaa = Replace(aaa," ",",")
		aaa = Replace(aaa,"|",",")
		hd.nkeyword=keyword_add(aaa)
		hd.addhuoDong()
		set hd = nothing
		Response.Cookies("Admin")("huoDongRenShu")=Request.Form("huoDongRenShu")
		Response.Cookies("Admin")("huoDongXingcheng")=Request.Form("huoDongXingcheng")
		Response.Cookies("Admin")("huoDongxieyi")=Request.Form("huoDongxieyi")
		Response.Cookies("Admin")("huoDongLingdui")=Request.Form("huoDongLingdui")
		Response.Cookies("Admin")("huoDongJifen")=Request.Form("huoDongJifen")
		Tourl("?Action=ListView")
	End Sub
	
	Public Sub DelAct()
		set hd=new HuoDong
		huoDongId = Cint(Request.QueryString("huoDongId"))
		hd.delhuoDong(huoDongId)
		set hd = nothing
		Tourl("?Action=ListView")
	End Sub
	
	Public Sub EditZhuangtai()
		zhuangtai = Replace(Request.QueryString("huodongZt"),"'","''")'weiman|yiman|chufa|jieshu
		huoDongId = Cint(Request.QueryString("huoDongId"))
		Set Hd = New HuoDong
		call hd.updateHuodongzhuangtai(zhuangtai,huoDongId)
		Tourl("?Action=ListView")
		set hd = nothing
	End Sub

	'#################################################################
	'活动列表的页面
	'#################################################################
	Public Sub ListView()
		HuodongId = Request.form("huoDongId")
		if huoDongId = "" then Tourl("?")
		set b = New Baoming
		rs = b.BaoMingPageList(HuodongId)
		if b.rs.eof then die "暂时没有任何数据"
		
	%>


<body>
<center><h1>积分结算</h1></center>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="39" valign="top"><table width="100%" height="37" border="0" cellpadding="0" cellspacing="1" bgcolor="#DDDDDD">
      <tr bgcolor="#FFFFFF">
        <td width="4%" height="35" align="center">编号</td>
        <td width="30%" align="center">活动名称</td>
        <td width="6%" align="center">状态</td>
        <td width="6%" align="center">已报</td>
        <td width="6%" align="center">应得积分</td>
        <td width="4%" align="center">消费积分</td>
        <td width="8%" align="center">出发时间</td>
        <td width="10%" align="center">返回时间</td>
        <td width="18%" align="center">操作</td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td valign="top">
    <form action="?Action=JiesuanAct&huoDongId=<%=HuodongId%>" method="POST" name="form1">
    <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#DDDDDD">
  <%
  do while not b.rs.eof
  %>
	  <tr bgcolor="#FFFFFF">
        <td width="4%" height="25" align="center"><%=b.rs("baoMingId")%></td>
        <td width="30%" height="25" align="center"><%="报名人昵称:"&b.rs("baoMingUserName")&" 真实姓名:"&b.rs("baoMingRealName")%></td>
        <td width="6%" height="25" align="center"><%=huoDongZtBack(b.rs("huoDongZhuangtai"))%></td>
        <td width="6%" height="25" align="center"><input type="text" name="baoMingRenshu" size="3" value="<%=replace(b.rs("baoMingRenshu")," ","")%>" /></td>
        <td width="6%" height="25" align="center"><input type="text" name="baoMingJifen" size="3" value="<%=replace(b.rs("baoMingJifen")," ","")%>" /></td>
        <td width="4%" height="25" align="center"><%=UseJifen(b.rs("baoMingUseJifen"))%></td>
        <td width="8%" height="25" align="center"><%=b.rs("huoDongChufaTime")%></td>
        <td width="10%" height="25" align="center"><%=b.rs("huoDongFanhuiTime")%></td>
        <td width="18%" height="25"><input type="hidden" name="baoMingUseJifen" value="<%=b.rs("baoMingUseJifen")%>" /><input type="hidden" name="baoMingUserName" value="<%=b.rs("baoMingUserName")%>" /><input type="hidden" name="huodongId" value="<%=b.rs("huodongId")%>" /><input type="hidden" name="baoMingId" value="<%=b.rs("baoMingId")%>" />
		<select onchange="location.href='?Action=EditZhuangtai&HuodongId=<%=b.rs("huoDongId")%>&huodongZt='+this.value+'';" name="zhuangtai"><option>选择</option><option value="0">未满</option><option value="1">已满</option><option value="2">出发</option><option value="3">结束</option></select>
		<a href="?Action=EditView&huoDongId=<%=b.rs("huoDongId")%>">修改</a> | <a href="?Action=DelAct&huoDongId=<%=b.rs("huoDongId")%>">删除</a>
		</td>
      </tr>
  <%
  b.rs.movenext()
  loop
  %>
  <tr bgcolor="#FFFFFF">
  <td><center><input type="submit" value="开始结算积分" /></center></td>
  </tr>
    </table>
    </form></td>
  </tr>
</table>
	<%
	b.rs.close
	set b.rs=nothing
	set b = nothing
	End Sub
	
	Public Sub SelectHuodong()
		Set h = new Huodong
		h.getJiesuanList()
		%>
		<center><h1>积分结算</h1></center>
		<form action="?Action=ListView" method="post" name="form1" id="form1">
			<center><select name="huoDongId">
			<%do while not h.rs.eof
			if h.rs("huoDongZhuangtai")=2 then
			%>
				<option value="<%=h.rs("huoDongId")%>"><%=h.rs("huoDongName")%>【已出发】</option>
			<%
			elseif h.rs("huoDongZhuangtai") = 3 then
			%>
				<option value="<%=h.rs("huoDongId")%>"><%=h.rs("huoDongName")%>【已结束】</option>
			<%
			elseif h.rs("huoDongZhuangtai") = 1 then
			%>
				<option value="<%=h.rs("huoDongId")%>"><%=h.rs("huoDongName")%>【满员】</option>
			<%
			elseif h.rs("huoDongZhuangtai") = 0 then
			%>
				<option value="<%=h.rs("huoDongId")%>"><%=h.rs("huoDongName")%>【未满】</option>
			<%
			end if
			h.rs.movenext()
			loop
			%>
			</select><br/>
			<input type="submit" value="好了，显示结算列表！" /></center>
		</form>
		<%
		h.rs.close
		set h = nothing
	End Sub
	
	
	Public Function getCookie(byval CookieName)
		Dim CookieValue
		CookieValue=""
		CookieValue = Replace(Request.Cookies("Admin")(CookieName),"'","''")
		if CookieValue <> "" then 
			getCookie = CookieValue
		end if
	End Function
%>