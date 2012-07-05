<!--#include file="../inc/utf.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/HuoDong.asp"-->
<!--#include file="../inc/baoming.asp"-->
<!--#include file="../inc/carlocation.asp"-->
<!--#include file="../inc/sys.asp"-->
<!--#include file="../inc/Users.asp"-->
<!--#include file="../inc/category.asp"-->
<!--#include file="../inc/tfunction.asp"-->
<!--#include file="isadmin.asp"-->
<%
	Action = Request.QueryString("Action")
	Select Case Action
		Case "ListView"
			Call ListView()
		Case "GetBaoxianList"
			Call getBaoxianList()
			die ""
		Case "HuishouView"
			Call HuishouView()
		Case "EditView"
			Call EditView()
		Case "AddView"
			Call AddView()
		Case "EditAct"
			Call EditAct()
		Case "AddAct"
			Call AddAct()
		Case "HelpSomeGuy"
			Call HelpSomeGuy()
		Case "HelpSomeGuyAct"
			Call HelpSomeGuyAct()
		Case "ChooseHuodong"
			Call ChooseHuodong()
		Case "AjaxGetUserInfo"
			Call AjaxGetUserInfo()
			die ""
		Case "DelAct"
			Call DelAct()
		Case "EditZhuangtai"
			Call EditZhuangtai()
		Case "PrintListView"
			Call PrintListView()
		Case "JiaJianRen"
			Call JiaJianRenView()
		Case "JiaRen"
			Call JiaRen()
		Case "JianRen"
			Call JianRen()
		Case "AjaxJiaRen"
			Call AjaxJiaRen()
			die ""
		Case "AjaxJianRen"
			Call AjaxJianRen()
			die ""
		Case Else
			Call ListView()
	End Select
	%>
<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
}
-->
</style>
	<%
	Public Sub AjaxGetUserInfo()
		UserName = Replace(Request.QueryString("UserName"),"'","''")
		Set u = New Users
		u.getuser(UserName)
		if not u.rs.eof then
			echo u.rs("RealName")&"|"&u.rs("Phone")&"|"&u.rs("IdCard")&"|"&u.rs("Jifen")
		else
			echo "error"
		end if
		Set u = Nothing
	End Sub
	
	Public Sub AjaxJiaRen()
		Response.Expires = -1 
		Response.ExpiresAbsolute = Now() - 1 
		Response.cachecontrol = "no-cache" 
		Dim yibao
		num = cInt(Replace(Request.QueryString("num"),"'","''"))
		huodongId = Replace(Request.QueryString("huodongId"),"'","''")
		id = Replace(Request.QueryString("id"),"'","''")
		set hd=new HuoDong
		hd.gethuoDong(huodongId)
		If Not hd.rs.eof Then
			huodongRenshu = hd.rs("huodongRenshu")
		End If
		hd.rs.close
		Set hd.rs = Nothing
		Set hd = Nothing
		newrenshu = huodongRenshu + num
		'conn.execute("update huodong set huodongrenshu=35 where huodongid=4")
		'die ""
		Conn.execute("update huodong set huodongRenshu="&newrenshu&" where huodongId="&huodongId)
		Set rss = Conn.execute("select sum([baoMingRenshu]) as zongshu from Baoming where baomingZhuangtai='0' and baoMingHuoDongId="&huodongId&"")
		If not isnumeric(rss("zongshu")) Then
			yibao = 0
		Else
			yibao = rss("zongshu")
		End if
		rss.close 
		Set rss = Nothing
		huodongShengyu = newrenshu - yibao
		if huodongShengyu>=1 then
			conn.execute("update huodong set huoDongZhuangTai=0 where huodongId="&huodongId)
		end if
		'die "update huoDong set huodongShengyu="&huodongShengyu&" where huodongId="&huodongId
		Conn.execute("update huoDong set huodongShengyu="&huodongShengyu&" where huodongId="&huodongId)
		die newrenshu&"|"&huodongShengyu&"|"&id
	End Sub

	Public Sub AjaxJianRen()
		Response.Expires = -1 
		Response.ExpiresAbsolute = Now() - 1 
		Response.cachecontrol = "no-cache" 
		Dim yibao
		num = cInt(Replace(Request.QueryString("num"),"'","''"))
		huodongId = Replace(Request.QueryString("huodongId"),"'","''")
		id = Replace(Request.QueryString("id"),"'","''")
		set hd=new HuoDong
		hd.gethuoDong(huodongId)
		If Not hd.rs.eof Then
			huodongRenshu = hd.rs("huodongRenshu")
		End If
		hd.rs.close
		Set hd.rs = Nothing
		Set hd = Nothing
		newrenshu = huodongRenshu - num
		'conn.execute("update huodong set huodongrenshu=35 where huodongid=4")
		'die ""
		Conn.execute("update huodong set huodongRenshu="&newrenshu&" where huodongId="&huodongId)
		Set rss = Conn.execute("select sum([baoMingRenshu]) as zongshu from Baoming where baomingZhuangtai='0' and baoMingHuoDongId="&huodongId&"")
		If not isnumeric(rss("zongshu")) Then
			yibao = 0
		Else
			yibao = rss("zongshu")
		End if
		rss.close 
		Set rss = Nothing
		huodongShengyu = newrenshu - yibao
		if huodongShengyu<0 then die "error"
		if huodongShengyu<=0 then
			conn.execute("update huodong set huoDongZhuangTai=1 where huodongId="&huodongId)
		end if
		'die "update huoDong set huodongShengyu="&huodongShengyu&" where huodongId="&huodongId
		Conn.execute("update huoDong set huodongShengyu="&huodongShengyu&" where huodongId="&huodongId)
		die newrenshu&"|"&huodongShengyu&"|"&id
	End Sub

	Public Sub HelpSomeGuyAct
		huodongZhuangtai = Replace(Request.Form("huodongZhuangtai"),"'","''")
		baoMingHuoDong = Replace(Request.Form("baoMingHuoDong"),"'","''")
		baoMingHuoDongId = cint(Replace(Request.Form("baoMingHuoDongId"),"'","''"))
		baoMingUserName = Replace(Request.Form("baoMingUserName"),"'","''")
		baoMingRealName = Replace(Request.Form("baoMingRealName"),"'","''")
		baoMingRenshu = Replace(Request.Form("baoMingRenshu"),"'","''")
		baoMingPhone = Replace(Request.Form("baoMingPhone"),"'","''")
		baomingLiuyan = Replace(Request.Form("baomingLiuyan"),"'","''")
		location = Replace(Request.Form("location"),"'","''")
		isBaoxiana = Replace(Request.Form("isBaoxian"),"'","''")
		baoxianList = Replace(Request.Form("baoxianList"),"'","''")
		baoMingIdcard = Replace(Request.Form("baoMingIdcard"),"'","''")
		baoMingUseJifen = cint(Request.Form("baoMingUseJifen"))
		baoMingJifen = Request.Form("baomingjifen")
		set h = new HuoDong
		h.gethuoDong(baoMingHuoDongId)
		if not h.rs.eof then 
			huoDongJifen = h.rs("huoDongJifen")
			huoDongShengyu = h.rs("huoDongShengyu")
		end if
		h.rs.close
		set h.rs = nothing
		set h = nothing
		baoMingZhuangtai = 0
		if huodongZhuangtai = "活动结束" then die "<script>alert('活动已经结束，不能报名，下次吧!');history.go(-1);</script>"
		if baoMingRenshu = "" then die "<script>alert('请填写人数!');history.go(-1);</script>"
		'需要判断当前活动剩余人数是否还有！
		if baoMingUseJifen = 1 then
			if baoMingJifen < 30 then
				die "<script>alert('您的积分不够不能使用积分支付!');history.go(-1);</script>"
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
        bm.isBaoxian = isBaoxiana
        bm.baoxianList = baoxianList
        bm.baoMingRealName = baoMingRealName
        bm.baomingLiuyan = baomingLiuyan
        bm.baoMingRenshu = baoMingRenshu
        bm.baoMingPhone = baoMingPhone
        bm.baoMingIdcard = baoMingIdcard
        bm.baoMingZhuangtai = baoMingZhuangtai
        bm.baoMingAddTime = now()
        bm.baoMingJifen = baoMingJifen
        bm.location = location
        bm.baoMingUseJifen = baoMingUseJifen
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
	    	die "<script>alert('报名人数超过现有剩余人数，现在"&aa&"');history.go(-1);</script>"
	    end if
	    set bm = Nothing
	    echo "<script>alert('报名成功！');location.href='Baoming.asp';</script>"
	End Sub
	
	Public Sub ChooseHuodong()
		Set h = New Huodong
		h.getHuoDongs 50,""
	%>
	<center><h1>帮助会员报名</h1></center>
	请在下面活动列表中点击一个需要报名的活动<br/><UL Style="Display:inline;float:left;">
	<%
		Do While Not h.rs.eof
	%>
	<Li><a href="?Action=HelpSomeGuy&huoDongId=<%=h.rs("huoDongId")%>&huoDongName=<%=Server.UrlEnCode(h.rs("huoDongName"))%>"><%=h.rs("huoDongName")%></a></Li>
	<%
		h.rs.Movenext()
		Loop
	%>
	</UL>
	<%
		h.rs.close
		set h.rs = nothing
		set h = nothing
	End Sub
	
	Public Sub HelpSomeGuy()
		huoDongName = Request.QueryString(URLDecode("huoDongName"))
		huoDongId = Request.QueryString("huoDongId")
	%>
		<script src="/js/Ajax.js" type="text/javascript"></script>
		<script type="text/javascript">
			function ajaxGetUser(obj){
				userName = obj.value;
				var ha = new QiAjax();
   				ha.ajax_get("?Action=AjaxGetUserInfo&UserName="+userName, call_back);
			}
		   function call_back(text,xml)
		   {
		     if(text!="error"){
		     	//alert(text);
		     	$("#warning").html("<font color=\"#FF0000\">恭喜该用户存在！系统自动填写表单！</font>")
		     	userArr = text.split("|");
			     $("#baoMingRealName").val(userArr[0]);
			     $("#baoMingPhone").val(userArr[1]);
			     $("#baoMingIdcard").val(userArr[2]);
			     $("#jifen").html(userArr[3]);
			     $("#baomingjifen").val(userArr[3]);
		     }else{
				
			     $("#warning").html("<font color=\"#FF0000\">该用户未注册，是否准备帮他<a href='?Action=RegUser'>注册一个</a>?</font>");
		     }
		     
		   }
		</script>
		<script type="text/javascript" src="http://xp.xinqu.cc/js/jquery-1.4.2.min.js"></script> 
  <!--右边start-->
<script>
checkform = function(obj){
	zuhe();
	return true;
}
</script>
<script language="javascript">
var num=0;
$(document).ready(function(){  
     $("#isBaoxian").click(function(){  
         if($(this).attr("checked")){  
             $("#baoxian").show();
         }else{  
             $("#baoxian").hide(); 
         }  
     })  
 });  


function del(o){
 document.getElementById("d").removeChild(document.getElementById("div_"+o));
 num = num - 1;
}
function addinput(val){
	num = val;
	document.getElementById("d").innerHTML= '';
	for(i = 0; i < num; i++){
		document.getElementById("d").innerHTML+='<div id="div_'+i+'">姓名:<input id="name_'+i+'" name="name_'+i+'" type="text" value=""  /> 身份证:<input type="text" id="idcard_'+i+'" name="idcard_'+i+'" value="610" /><input type="button" value="删除"  onclick="del('+i+')"/></div>';
	}
}
function zuhe(){
	var data='';
	//alert(num);
	for(i = 0; i< num; i++){
		var name = document.getElementById("name_"+i).value;
		var idcard = document.getElementById("idcard_"+i).value;
		data += name+"$"+idcard+"|"
	}
	$("#baoxianList").val(data);
}

</script>
		<form action="?Action=HelpSomeGuyAct&huoDongId=<%=huoDongId%>" method="post" name="form1" id="form1" onsubmit="return checkform()">
		<input type="submit" name="xbutton" id="xxbutton" value="下 一 步" /><br/>
		<input type="hidden" name="baoMingHuoDong" value="<%=huoDongName%>" />
        <input type="hidden" name="baoMingHuoDongId" value="<%=huoDongId%>" />
        <p>报名用户:
            <input type="text" name="baoMingUserName" value="" onBlur="ajaxGetUser(this);" /><span id="warning"><font color="#FF0000">只需输入用户名和报名人数即可</font></span>
          <br/>
          报名人数:
          <input type="text" name="baoMingRenshu" id="baoMingRenshu" value="1" />
          <br/>
          真实姓名:
          <input type="text" name="baoMingRealName" id="baoMingRealName" value="" />
          <br/>
          联系电话:
          <input type="text" name="baoMingPhone" id="baoMingPhone" value="" /><br/><font color="#FF0000">(温馨提示：用于接收短信请尽量填写手机号码)</font>
          <br/>
          身份证号:
          <input type="text" name="baoMingIdcard" id="baoMingIdcard" value="" /><br/>
          乘车地点：
          <%=GetCarLocationView()%><br/>
             <!--保险-->
                             是否需要购买保险
<input type="checkbox" name="isBaoxian" id="isBaoxian" value="1"><br/>
<div id="baoxian" style="display:none">
	输入要购买保险人数：<input onkeyup="addinput(this.value);this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')" id="baoxianRenshu" name="baoxianRenshu">
	 
	<div id="d">
	</div>
	<input type="hidden" name="baoxianList" id="baoxianList">
</div>
							<!--保险-->
          当前积分:<span id="jifen"></span><br/>
          <input type="hidden" name="baoMingUseJifen" value="0" />
          用户电话:<span id="phone"></span><br/>
          留言:<textarea type="text" name="baomingLiuyan" id="baomingLiuyan" cols="30" rows="8">无</textarea>
          <br/><input type="hidden" name="baomingjifen" id="baomingjifen" value="">
          <input type="hidden" name="huodongZhuangtai" id="huodongZhuangtai" value="">
    </form>
                              
                     
	<%
	End Sub

	Public Sub JiaRen()
		Renshu = cint(Request.Form("Renshu"))
		userName =Request("userName")
		baomingId = Request("baomingId")
		huodongId = Request("huodongId")
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
				die "<script>alert('恭喜您！积分累加成功，加人成功，管理员辛苦了！');history.go(-1);</script>"
			else
				die "<script>alert('对不起！本活动已经没有剩余座位了！');history.go(-1);</script>"
			end if
		end if
		h.rs.close
		b.rs.close
		set b = nothing
		set h = nothing
	End Sub

	Public Sub JianRen()
		Renshu = cint(Request.Form("Renshu"))
		userName =Request("userName")
		baomingId = Request("baomingId")
		huodongId = Request("huodongId")
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
				die "<script>alert('积分按照人数比例已经相应扣除！管理员辛苦了！');history.go(-1);</script>"
			else
				die "<script>alert('您减少的人数应该比你的报名人数小于或者等于才可以取消！');history.go(-1);</script>"
			end if
		end if
	End Sub
	
	Public Sub JiaJianRenView()
		userName =Request.QueryString("userName")
		baomingId = Request.QueryString("baomingId")
		huodongId = Request.QueryString("huodongId")
	%>
	<table width="100%" height="493" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td height="184" align="center" valign="top"><form id="form1" name="form1" method="post" action="?Action=JiaRen&baomingId=<%=baomingId%>&huodongId=<%=huodongId%>">
                  <table width="100%" height="75" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td align="center">添加活动人数</td>
                    </tr>
                  </table>
                  <table width="100%" height="93" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="36%" align="right">增加数量：</td>
                      <td width="64%" align="left">
					  <select name="Renshu">
					  <option value="1">增加1人</option>
					  <option value="2">增加2人</option>
					  <option value="3">增加3人</option>
					  <option value="4">增加4人</option>
					  <option value="5">增加5人</option>
					  <option value="6">增加6人</option>
					  <option value="7">增加7人</option>
					  <option value="8">增加8人</option>
					  <option value="9">增加9人</option>
					  <option value="10">增加10人</option>
					  </select></td>
                    </tr>
                    <tr><td width="36%" align="right">
                    用户名
                    </td>
                    <td width="64%" align="left"><input type="text" name="userName" value="<%=userName%>"></td>
                    </tr>
                    <tr>
                      <td colspan="2" align="center"><input type="submit" name="Submit" value="好了，我要加人！" /></td>
                    </tr>
                  </table>
                  </form></td>
              </tr>
              <tr>
                <td align="center" valign="top"><form id="form1" name="form1" method="post" action="?Action=JianRen&baomingId=<%=baomingId%>&huodongId=<%=huodongId%>">
                  <table width="100%" height="75" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td align="center">退订活动人数</td>
                    </tr>
                  </table>
                  <table width="100%" height="93" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="36%" align="right">减少人数：</td>
                      <td width="64%" align="left"><select name="renshu">
                        <option value="1">减少1人</option>
                        <option value="2">减少2人</option>
                        <option value="3">减少3人</option>
                        <option value="4">减少4人</option>
                        <option value="5">减少5人</option>
                        <option value="6">减少6人</option>
                        <option value="7">减少7人</option>
                        <option value="8">减少8人</option>
                        <option value="9">减少9人</option>
                        <option value="10">减少10人</option>
                      </select></td>
                    </tr>
                    <tr><td width="36%" align="right">
                    用户名
                    </td>
                    <td width="64%" align="left"><input type="text" name="userName" value="<%=userName%>"></td>
                    </tr>
                    <tr>
                    <tr>
                      <td colspan="2" align="center"><input type="submit" name="Submit" value="好了，我要减人！" /></td>
                    </tr>
                  </table>
                                </form>         <br/><br/><br/>  <input type="button" value="返回列表" onClick="location.href='baoming.asp';" />     </td>
                </tr>
              
            </table>
	<%
	End Sub
	
	Public Sub EditAct()
		set hd=new Baoming
		hd.baoMingId = Replace(Request.QueryString("baoMingId"),"'","''")
		hd.baoMingUserName = Replace(Request.Form("baoMingUserName"),"'","''")
		hd.baoMingRealName = Replace(Request.Form("baoMingRealName"),"'","''")
		hd.baoMingRenshu = Replace(Request("baoMingRenshu"),"'","''")
		hd.baoMingPhone = Replace(Request.Form("baoMingPhone"),"'","''")
		hd.baoMingIdcard = Replace(Request.Form("baoMingIdcard"),"'","''")
		hd.baoMingJifen = Replace(Request.Form("baoMingJifen"),"'","''")
		hd.baoMingAddTime = Replace(Request.Form("baoMingAddTime"),"'","''")
		hd.baomingUseJifen = Replace(Request.Form("baomingUseJifen"),"'","''")
		hd.baoMingLiuyan = Replace(Request.Form("baoMingLiuyan"),"'","''")
		hd.updatebaoMing()
		set hd = nothing
		Tourl("baoming.asp")
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
		Tourl("baoming.asp")
	End Sub
	
	Public Sub DelAct()
		set bm=new Baoming
		set h = new HuoDong
		
		baomingId = CLng(Request.QueryString("baomingId"))
		huoDongId = CLng(Request.QueryString("huoDongId"))
		renshu = CLng(Request.QueryString("renshu"))
		bm.updatebaoMingzt 1,baomingId
		call h.jiaShengyu(Renshu,huoDongId)
		set bm = nothing
		Tourl("baoming.asp")
	End Sub
	
	Public Sub EditZhuangtai()
		zhuangtai = Replace(Request.QueryString("huodongZt"),"'","''")'weiman|yiman|chufa|jieshu
		huoDongId = Cint(Request.QueryString("huoDongId"))
		Set Hd = New HuoDong
		call hd.updateHuodongzhuangtai(zhuangtai,huoDongId)
		Tourl("baoming.asp")
		set hd = nothing
	End Sub


	Public Sub HuishouView()
		set b = New Baoming
		rs = b.HuishouBaoMing()
		if b.rs.eof then die "暂时没有任何数据"
		b.rs.PageSize = 20
		Page = CLng(Request("Page"))
	    If Page < 1 Then Page = 1
	    If Page > b.rs.PageCount Then Page = b.rs.PageCount
	    b.rs.AbsolutePage = Page
%>
<body>
<center><h1>报名回收站</h1></center>
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
    <td valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#DDDDDD">
  <%
  for i=1 to b.rs.PageSize
  if b.rs.EOF then Exit For
  %>
	  <tr bgcolor="#FFFFFF">
        <td width="4%" height="25" align="center"><%=b.rs("baoMingId")%></td>
        <td width="30%" height="25" align="center"><a href="?Action=PrintListView&huoDongId=<%=b.rs("baoMingHuoDongId")%>"><%=b.rs("baoMingHuoDong")%></a> 报名人:<%=b.rs("baoMingUserName")%></td>
        <td width="6%" height="25" align="center"><%=huoDongZtBack(b.rs("huoDongZhuangtai"))%></td>
        <td width="6%" height="25" align="center"><%=b.rs("baoMingRenshu")%></td>
        <td width="6%" height="25" align="center"><%=b.rs("baoMingJifen")%></td>
        <td width="4%" height="25" align="center"><%=UseJifen(b.rs("baoMingUseJifen"))%></td>
        <td width="8%" height="25" align="center"><%=b.rs("huoDongChufaTime")%></td>
        <td width="10%" height="25" align="center"><%=b.rs("huoDongFanhuiTime")%></td>
        <td width="18%" height="25">
		 <a href="?Action=DelBaoming&baoMingId=<%=b.rs("baomingId")%>">彻底删除</a>
		</td>
      </tr>
  <%
  b.rs.movenext()
  next
  %>
    </table></td>
  </tr>
  <tr>
    <td height="30" align="center" valign="middle" bgcolor="#FFFFFF">
    <%
		if page=1 then
			echo "首页 上一页 第"&page&"页 <a href='baoming.asp?page="&page+1&"'>下一页</a>"
			echo "<a href='baoming.asp?page="&b.rs.pagecount&"'>尾页</a>"
		elseif page=b.rs.pagecount then
			echo "<a href='baoming.asp?page=1'>首页</a> <a href='baoming.asp?page="&page-1&"'>上一页</a> 第"&page&"页 下一页 尾页"
		else
			echo "<a href='baoming.asp?page=1'>首页</a>&nbsp;<a href='baoming.asp?page="&page-1&"'>上一页</a> 第"&page&"页 <a href='baoming.asp?page="&page+1&"'>下一页</a> <a href='baoming.asp?page="&b.rs.pagecount&"'>尾页</a>"
		end if
			%>
    </td>
  </tr>
</table>
	<%
	b.rs.close
	set b.rs=nothing
	set b = nothing
	End Sub

Public Sub getBaoxianList()
	baomingId = request.queryString("baomingId")
	set bm = New Baoming
	bm.getbaoMing(baomingId)
	baoxianList = bm.rs("baoxianList")
	ArrayList = Split(baoxianList,"|")
	for i = 0 to ubound(ArrayList) -1
		arrayList2 = Split(ArrayList(i),"$")
		if instr(ArrayList(i),"$") >=0 then
			echo "姓名:"&arrayList2(0)&" 身份证:"&arrayList2(1)
			echo "<br/>"
		end if
	next
	bm.rs.close
	set bm.rs = nothing
	set bm = nothing
End Sub

Public Function getBaoxianListFun(baomingId)
	set bm = New Baoming
	bm.getbaoMing(baomingId)
	baoxianList = bm.rs("baoxianList")
	ArrayList = Split(baoxianList,"|")
	dim str
	for i = 0 to ubound(ArrayList) -1
		arrayList2 = Split(ArrayList(i),"$")
		if instr(ArrayList(i),"$") >=0 then
			str = str& "姓名:"&arrayList2(0)&" 身份证:"&arrayList2(1)
			str = str& "<br/>"
		end if
	next
	bm.rs.close
	set bm.rs = nothing
	set bm = nothing
	getBaoxianListFun = str
End Function
	'#################################################################
	'活动列表的页面
	'#################################################################
	Public Sub ListView()
		showid = request.querystring("showid")
		searchKey = Request.querystring("searchkey")
		set b = New Baoming
		if IsNumeric(showid) and showid <> "" then
			rs = b.BaoMingPageList(showid)
			showidstr = "&showid="&showid
		elseif searchKey<>"" then
			rs = b.SearchBaoming(searchKey)
			showidstr = "&searchKey="&server.URLENCODE(searchKey)
		else
			rs = b.allBaoMing()
		end if
		if b.rs.eof then die "暂时没有任何数据"
		b.rs.PageSize = 20
		Page = CLng(Request("Page"))
	    If Page < 1 Then Page = 1
	    If Page > b.rs.PageCount Then Page = b.rs.PageCount
	    b.rs.AbsolutePage = Page
	%>

<style>
.tip{
	text-decoration:none;
	position:relative;
	z-index: 500;
}
.tip span {display:none;}
.tip:hover {cursor:hand;}
.tip:hover .popbox {
	display:block; 
	position:absolute; 
	top:-15px;
	left:180px;
	width:400px; 
	background-color:#424242; 
	color:#fff; 	
	padding:5px;
	z-index: 9999;
}
.tip:hover .popboxBaoxian {
	display:block; 
	position:absolute; 
	top:-15px;
	left:23px;
	width:300px; 
	background-color:#424242; 
	color:#fff; 	
	padding:5px;
	z-index: 9999;
}</style>
<body>
<center><h1>报名管理</h1></center><br/>
<p align="right">
	<select name="huodong" onchange="location.href='?showId='+this.value;">
		<option value="">请选择要单独查看的活动</option>
		<%
		Set hd = New Huodong
		hd.getSelecthuodong()
		do while not hd.rs.eof
			%>
			<option value="<%=hd.rs("huodongId")%>"><%=hd.rs("huodongName")%></option>
			<%
			hd.rs.movenext
		loop
		hd.rs.close
		set hd.rs = nothing
		%>
	</select>
</p>
<p align="left">
<form action="" name="search" method="GET">
	<input type="text" name="searchKey" value="姓名|电话|身份证|网名" style="width:200px;height:24px;" onfocus="if(this.value=='姓名|电话|身份证|网名'){this.value='';}" onblur="if(this.value==''){this.value=='姓名|电话|身份证|网名';}"><input type="submit" value="查找" />
</form>
</p>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="39" valign="top"><table width="100%" height="37" border="0" cellpadding="0" cellspacing="1" bgcolor="#DDDDDD">
      <tr bgcolor="#FFFFFF">
        <td width="4%" height="35" align="center">编号</td>
        <td width="30%" align="center">活动名称</td>
        <td width="6%" align="center">状态</td>
        <td width="6%" align="center">用户名</td>
        <td width="6%" align="center">已报</td>
        <td width="6%" align="center">应得积分</td>
        <td width="4%" align="center">买保险</td>
        <td width="8%" align="center">出发时间</td>
        <td width="10%" align="center">返回时间</td>
        <td width="18%" align="center">操作</td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#DDDDDD">
  <%
  for i=1 to b.rs.PageSize
  if b.rs.EOF then Exit For
  %>
	  <tr bgcolor="#FFFFFF">
        <td width="4%" height="25" align="center"><%=b.rs("baoMingId")%></td>
        <td width="30%" height="25" align="center"><a href="?Action=PrintListView&huoDongId=<%=b.rs("baoMingHuoDongId")%>" target="_blank" class="tip"><%=cutStr(b.rs("baoMingHuoDong"), 28)%><span class="popbox"><%=b.rs("baoMingHuoDong")%></span></a></td>
        <td width="6%" height="25" align="center"><%=huoDongZtBacktext(b.rs("huoDongZhuangtai"))%></td>
        <td width="6%" align="center"><%=b.rs("baomingUserName")%></td>
        <td width="6%" height="25" align="center"><%=b.rs("baoMingRenshu")%></td>
        <td width="6%" height="25" align="center"><%=b.rs("baoMingJifen")%></td>
        <td width="4%" height="25" align="center"><%
        bx = isBaoxian(b.rs("isBaoxian"))
        if bx = "YES" then
        	echo "<a href=""?Action=GetBaoxianList&baomingId="&b.rs("baoMingId")&""" target=""_blank"" class=""tip""><font color=""#FF0000"">YES</font><span class=""popboxBaoxian"">"&getBaoxianListFun(b.rs("baoMingId"))&"</span></a>"
        else
        	echo "NO"
        end if
        %></td>
        <td width="8%" height="25" align="center"><%=b.rs("huoDongChufaTime")%></td>
        <td width="10%" height="25" align="center"><%=b.rs("huoDongFanhuiTime")%></td>
        <td width="18%" height="25">
		<a href="?Action=EditView&baomingId=<%=b.rs("baomingId")%>">修改</a> | <%if b.rs("huoDongZhuangtai") = 2 or b.rs("huoDongZhuangtai") = 3 then%><font color="#999999">退活动</font> | <font color="#999999">加减人</font><%else%><a href="?Action=DelAct&baomingId=<%=b.rs("baomingId")%>&huoDongId=<%=b.rs("baoMingHuoDongId")%>&renshu=<%=b.rs("baoMingRenshu")%>">退活动</a> | <a href="?Action=JiaJianRen&userName=<%=b.rs("baoMingUserName")%>&baomingId=<%=b.rs("baomingId")%>&huodongId=<%=b.rs("huodongId")%>">加减人</a><%end if%></td>
      </tr>
  <%
  b.rs.movenext()
  next
  %>
    </table></td>
  </tr>
  <tr>
    <td height="30" align="center" valign="middle" bgcolor="#FFFFFF">
    <%
		if page=1 then
			echo "首页 上一页 第"&page&"页 <a href='baoming.asp?page="&page+1&showidstr&"'>下一页</a>"
			echo "<a href='baoming.asp?page="&b.rs.pagecount&showidstr&"'>尾页</a>"
		elseif page=b.rs.pagecount then
			echo "<a href='baoming.asp?page=1'>首页</a> <a href='baoming.asp?page="&page-1&showidstr&"'>上一页</a> 第"&page&"页 下一页 尾页"
		else
			echo "<a href='baoming.asp?page=1'>首页</a>&nbsp;<a href='baoming.asp?page="&page-1&showidstr&"'>上一页</a> 第"&page&"页 <a href='baoming.asp?page="&page+1&showidstr&"'>下一页</a> <a href='baoming.asp?page="&b.rs.pagecount&showidstr&"'>尾页</a>"
		end if
			%>
    </td>
  </tr>
</table>
	<%
	b.rs.close
	set b.rs=nothing
	set b = nothing
	End Sub
	
	'#################################################################
	'帮会员报名
	'#################################################################
	Public Sub AddView()
	%>
			
		<form action="?Action=AddAct" method="post" name="form1" id="form1">
		  <table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999"><tr><td colspan="2" bgcolor="#FFFFFF" ><table width="100%" border="0"><tr><td align="center"><table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999">
		    <tr>
		      <td width="150" bgcolor="#FFFFFF" >活动名称：</td>
		      <td bgcolor="#FFFFFF" align="left" ><input name="huoDongName" type="text" value="" /></td>
		    </tr>
		    <tr>
		      <td width="150" bgcolor="#FFFFFF" >论坛链接：</td>
		      <td bgcolor="#FFFFFF" align="left" ><input name="bbsLink" type="text" value="http://bbs.029lvwo.com" /></td>
		    </tr>
		    <tr>
		      <td width="150" bgcolor="#FFFFFF" >活动人数：</td>
		      <td  align="left" bgcolor="#FFFFFF"><input name="huoDongRenShu" type="text" value="<%=getCookie("huoDongRenShu")%>" />      </td>
		    </tr>
		    <tr>
		      <td bgcolor="#FFFFFF" >活动领队：</td>
		      <td  align="left" bgcolor="#FFFFFF"><input name="huoDongLingdui" type="text" value="<%=getCookie("huoDongLingdui")%>" /> <a href="javascript:void(0);" onClick="document.form1.huoDongLingdui.value='橡皮';">橡皮</a></td>
		    </tr>
		    <tr>
		      <td width="150" bgcolor="#FFFFFF" >活动状态：</td>
		      <td bgcolor="#FFFFFF"  align="left"><input name="huoDongZhuangtai" type="text" value="0" />  <a href="javascript:void(0);" onClick="document.form1.huoDongZhuangtai.value='0';">活动未满</a> | <a href="javascript:void(0);" onClick="document.form1.huoDongZhuangtai.value='1';">活动已满</a> | <a href="javascript:void(0);" onClick="document.form1.huoDongZhuangtai.value='2';">活动出发</a> | <a href="javascript:void(0);" onClick="document.form1.huoDongZhuangtai.value='3';">活动结束</a>   </td>
		    </tr>
		    <tr>
		      <td width="150" bgcolor="#FFFFFF" >活动积分：</td>
		      <td bgcolor="#FFFFFF"  align="left"><input name="huoDongJifen" type="text" value="<%=getCookie("huoDongJifen")%>" />      </td>
		    </tr>
		    <tr>
		      <td width="150" bgcolor="#FFFFFF" >出发时间：</td>
		      <td bgcolor="#FFFFFF"  align="left"><input id="d5221" name="huoDongChufaTime" class="Wdate" type="text" onFocus="var d5222=$dp.$('d5222');WdatePicker({onpicked:function(){d5222.focus();},maxDate:'#F{$dp.$D(\'d5222\')}'})"/></td>
		    </tr>
		    <tr>
		      <td width="150" bgcolor="#FFFFFF" >返回时间：</td>
		      <td bgcolor="#FFFFFF"  align="left"><input id="d5222" name="huoDongFanhuiTime" class="Wdate" type="text" onFocus="WdatePicker({minDate:'#F{$dp.$D(\'d5221\')}'})"/>      </td>
		    </tr>
		      <tr>
			    <td width="150" bgcolor="#FFFFFF">分类简介：</td>
			    <td bgcolor="#FFFFFF"><textarea name="huoDongInfo" cols="60" rows="15" id="huoDongInfo"></textarea></td>
			  </tr>
		    <tr>
		      <td width="150" bgcolor="#FFFFFF" >活动行程安排：</td>
		      <td bgcolor="#FFFFFF"  align="left">
				      <%  
		  Set oFCKeditor = New FCKeditor 
		  oFCKeditor.BasePath = "../fckeditor/"  
		  oFCKeditor.ToolbarSet = "Default" 
		  oFCKeditor.Width = "650" 
		  oFCKeditor.Height = "400" 
		  oFCKeditor.Value = getCookie("huoDongXingcheng")
		  oFCKeditor.Create "huoDongXingcheng" 
		 %></td>
		    </tr>
		    <tr>
		      <td width="150" bgcolor="#FFFFFF" >本次活动协议：</td>
		      <td bgcolor="#FFFFFF"  align="left"><%  
		  Set oFCKeditor = New FCKeditor 
		  oFCKeditor.BasePath = "../fckeditor/"  
		  oFCKeditor.ToolbarSet = "Default" 
		  oFCKeditor.Width = "650" 
		  oFCKeditor.Height = "400" 
		  oFCKeditor.Value = getCookie("huoDongxieyi")
		  oFCKeditor.Create "huoDongxieyi" 
		 %></td>
		    </tr>
	<tr>
    <td width="10%" bgcolor="#FFFFFF">关键字：</td>
    <td bgcolor="#FFFFFF">
      <input name="nkeyword" type="text" id="nkeyword" size="50" /> <%call gettag()%>
    </td>
  	</tr>
		    <tr>
		      <td colspan="2" bgcolor="#FFFFFF" ><table width="100%" border="0">
		          <tr>
		            <td align="center"><input type="submit" value="保存设置" />            </td>
		          </tr>
		      </table></td>
		    </tr>
		  </table></td>
		  </tr>
		</table>      </td>
		    </tr>
		</table>
		</form>
	<%
	End Sub
	
	
	
	'#################################################################
	'修改活动的表单页面
	'#################################################################
	Public Sub EditView()
		baomingId = Cint(Request.QueryString("baomingId"))
		set hd=new Baoming
		hd.getbaoMing(baomingId)
	%>
			<!--修改活动-->
		<form action="?Action=EditAct&baoMingId=<%=baomingId%>" method="post" name="form1" id="form1">
		  <table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999"><tr><td colspan="2" bgcolor="#FFFFFF" ><table width="100%" border="0"><tr><td align="center"><table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999">
		    <tr>
		      <td width="150" bgcolor="#FFFFFF" >活动名称：</td>
		      <td bgcolor="#FFFFFF" align="left" ><input name="baoMinghuodong" type="text" value="<%=hd.rs("baoMinghuodong")%>" readonly disabled /></td>
		    </tr>
		    <tr>
		      <td width="150" bgcolor="#FFFFFF" >报名用户：</td>
		      <td bgcolor="#FFFFFF" align="left" ><input name="baoMingUserName" type="text" value="<%=hd.rs("baoMingUserName")%>" /></td>
		    </tr>
		    <tr>
		      <td width="150" bgcolor="#FFFFFF" >真实姓名：</td>
		      <td  align="left" bgcolor="#FFFFFF"><input name="baoMingRealName" type="text" value="<%=hd.rs("baoMingRealName")%>" />      </td>
		    </tr>
		    <tr>
		      <td bgcolor="#FFFFFF" >报名人数：</td>
		      <td  align="left" bgcolor="#FFFFFF"><input name="baoMingRenshu" type="text" value="<%=hd.rs("baoMingRenshu")%>" readonly /> (人数不能直接修改，在外面加减人就可以)</td>
		    </tr>
		    <tr>
		      <td width="150" bgcolor="#FFFFFF" >报名人电话：</td>
		      <td bgcolor="#FFFFFF"  align="left"><input name="baoMingPhone" type="text" value="<%=hd.rs("baoMingPhone")%>" /> </td>
		    </tr>
		    <tr>
		      <td width="150" bgcolor="#FFFFFF" >报名身份证：</td>
		      <td bgcolor="#FFFFFF"  align="left"><input name="baoMingIdcard" type="text" value="<%=hd.rs("baoMingIdcard")%>" /> </td>
		    </tr>
		    <tr>
		      <td width="150" bgcolor="#FFFFFF" >本次报名积分：</td>
		      <td bgcolor="#FFFFFF"  align="left"><input name="baoMingJifen" type="text" value="<%=hd.rs("baoMingJifen")%>" />      </td>
		    </tr>
		    <tr>
		      <td width="150" bgcolor="#FFFFFF" >报名添加时间：</td>
		      <td bgcolor="#FFFFFF"  align="left"><input id="d5221" name="baoMingAddTime" class="Wdate" type="text" value="<%=hd.rs("baoMingAddTime")%>" onFocus="var d5222=$dp.$('d5222');WdatePicker({onpicked:function(){d5222.focus();},maxDate:'#F{$dp.$D(\'d5222\')}'})"/>  </td>
		    </tr>
		    <tr>
		      <td width="150" bgcolor="#FFFFFF" >报名使用积分：</td>
		      <td bgcolor="#FFFFFF"  align="left"> <input type="radio" name="baomingUseJifen" value="1" <%if hd.rs("baomingUseJifen")=1 then echo "checked"%>> 是  <input type="radio" name="baomingUseJifen" value="0" <%if hd.rs("baomingUseJifen")=0 then echo "checked"%>>否    </td>
		    </tr>
		      <tr>
			    <td width="150" bgcolor="#FFFFFF">报名留言：</td>
			    <td bgcolor="#FFFFFF"><textarea name="baoMingLiuyan" cols="60" rows="15" id="baoMingLiuyan" readonly><%=hd.rs("baoMingLiuyan")%></textarea></td>
			  </tr>
		   
		    <tr>
		      <td colspan="2" bgcolor="#FFFFFF" ><table width="100%" border="0">
		          <tr>
		            <td align="center"><input type="submit" value="保存设置" />            </td>
		          </tr>
		      </table></td>
		    </tr>
		  </table></td>
		  </tr>
		</table>      </td>
		    </tr>
		</table>
		</form>
	<%
	hd.rs.close
	set hd.rs = nothing
	set hd = nothing
	End Sub
	
	Public Sub PrintListView()
	HuodongId = Request.QueryString("huoDongId")
	set b = New Baoming
	rs = b.BaoMingPageList(HuodongId)
    sql = "select sum(baomingRenshu) as zongRenshu from Baoming where baoMingHuoDongId="&HuodongId&" and baomingZhuangtai='0' "
  	zongRenshu = conn.execute(sql)(0)
	%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>无标题文档</title>
  <script language="javascript">  
      <!--   
      function printWithAlert() {   
        document.all.WebBrowser.ExecWB(6,1);   
      }   
  
      function printWithoutAlert() {   
        document.all.WebBrowser.ExecWB(6,6);   
      }   
  
      function printSetup() {   
        document.all.WebBrowser.ExecWB(8,1);   
      }   
  
      function printPrieview() {   
        document.all.WebBrowser.ExecWB(7,1);   
      }   
  
      function printImmediately() {   
        document.all.WebBrowser.ExecWB(6,6);   
        window.close();   
      }   
      -->  
  </script>  
  <OBJECT  id=WebBrowser  classid=CLSID:8856F961-340A-11D0-A96B-00C04FD705A2 style="display:none">  </OBJECT>  
<style rel="stylesheet" type="text/css" media="all" />
<!--
body {
margin: 9px;
padding: 0;
color: black;
text-decoration: none;
font-size: 10px;
font-family: "Courier New";
}
table#dd {
background-color: #6CADD9;
}


table#dd thead th {
background-color: #6CADD9;
color: #FFFFFF;

}
#dd td {
padding: 2px;
width: 100px;
}
table#dd tbody.tb1 td {

background-color: #FFFFFF;
}
table#dd tbody.tb2 td {
background-color: #F7F7F7;
}
table#dd tbody td:hover {
background-color: #BFEDF9;
}
table#dd colgroup col.name {
background-color: #E6E6E6;
width: 100pt;
font-weight: normal;
}
@media print {
.noprint { display: none; }
}-->
</style>

  <style media=print>  
  .Noprint{display:none;}   
  .PageNext{page-break-after: always;}   
  </style>  
  <style media=print>  
 table {   
    background-color: #FFFFFF;   
    border-top-width: 1px;   
    border-right-width: 1px;   
    border-bottom-width: 1px;   
    border-left-width: 1px;   
    border-top-style: none;   
    border-right-style: solid;   
    border-bottom-style: solid;   
    border-left-style: none;   
    border-top-color: #000000;   
    border-right-color: #000000;   
    border-bottom-color: #000000;   
    border-left-color: #000000;   
    height:20pt;   
   }   

      
   th{   
   background-color: #FFFFFF;   
    border-top-width: 1px;   
    border-right-width: 1px;   
    border-bottom-width: 1px;   
    border-left-width: 1px;   
    border-top-style: solid;   
    border-right-style: none;   
    border-bottom-style: none;   
    border-left-style: solid;   
    border-top-color: #000000;   
    border-right-color: #000000;   
    border-bottom-color: #000000;   
    border-left-color: #000000;   
   }   
   td{
   background-color: #FFFFFF;   
    border-top-width: 1px;   
    border-right-width: 1px;   
    border-bottom-width: 1px;   
    border-left-width: 1px;   
    border-top-style: solid;   
    border-right-style: none;   
    border-bottom-style: none;   
    border-left-style: solid;   
    border-top-color: #000000;   
    border-right-color: #000000;   
    border-bottom-color: #000000;   
    border-left-color: #000000;   
   }
  </style>  
</head>
<body>

  
  <table border="0" align="center" cellspacing="1" id="dd">
    <caption>
    <h3>活动名称:<%=b.rs("baoMingHuoDong")%>人数合计:<%=zongRenshu%> 旅行者报名系统授权：西安驴窝</h3>
    </caption>
    <colgroup>
	    <col />
	    <col />
	    <col class="name" />
	    <col />
	    <col class="name" />
	    <col />
	    <col class="name" />
    </colgroup>
    <thead>
      <tr>
        <th>用户名</th>
        <th>真实姓名</th>
        <th width="130">集合</th>
        <th>报名人数</th>
        <th>是否买保险</th>
        <th>联系电话</th>
        <th>留言</th>
        <th>签到</th>
      </tr>
    </thead>
<%
dim phonelist
Do while not b.rs.eof
phonelist = phonelist & b.rs("baoMingPhone") & chr(10)
%>
    <tbody class="tb1">
      <tr>
        <td><%=b.rs("baoMingUserName")%></td>
        <td><%=b.rs("baoMingRealName")%>&nbsp;</td>
        <td width="130"><%=b.rs("location")%>&nbsp;</td>
        <td><%=b.rs("baoMingRenshu")%></td>
        <td><%=isBaoxianCn(b.rs("isBaoxian"))%>&nbsp;</td>
        <td><%=b.rs("baoMingPhone")%>&nbsp;</td>
        <td><%=b.rs("baomingLiuyan")%></td>
        <td>&nbsp;</td>
      </tr>
    </tbody>
<%
b.rs.movenext
Loop
%>
</table>
<center>
<input type="button" onclick="document.getElementById('phoneList').style.display='';" value="导出手机列表" class="Noprint">
    <input type=button value="打印" onclick="printWithAlert()" class="Noprint">  
    <input type=button value="直接打印" onclick="printWithoutAlert()" class="Noprint">  
    <input type=button value="打印设置" onclick="printSetup()" class="Noprint">  
    <input type=button value="预览" onclick="printPrieview()" class="Noprint">  
</center>

<div id="phoneList" style="display:none;" align="center"><textarea cols="35" rows="20"><%=phonelist%></textarea><br/><input type="button" onclick="document.getElementById('phoneList').style.display='none';" value="关闭"></div>
</body>
</html>

	<%
	b.rs.close
	set b = nothing
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