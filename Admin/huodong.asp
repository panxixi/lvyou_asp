<!--#include file="../inc/utf.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/HuoDong.asp"-->
<!--#include file="../inc/sys.asp"-->
<!--#include file="../inc/category.asp"-->
<!--#include file="../inc/tfunction.asp"-->
<!--#include file="isadmin.asp"-->
<!--#include file="HuoDongAutoTag.asp" -->
<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
}
-->
</style>
<script language="JavaScript" type="text/javascript" src="my97datepicker/WdatePicker.js"></script>
<%
	Action = Request.QueryString("Action")
	Select Case Action
		Case "ListView"
			Call ListView()
		Case "EditView"
			Call EditView()
		Case "AddView"
			Call AddView()
		Case "EditAct"
			Call EditAct()
		Case "AddAct"
			Call AddAct()
		Case "DelAct"
			Call DelAct()
		Case "EditZhuangtai"
			Call EditZhuangtai()
		Case "AjaxJiaRen"
			Call AjaxJiaRen()
		Case Else
			Call ListView()
	End Select

	Public Sub EditAct()
		set hd=new HuoDong
		hd.huoDongId = cInt(Replace(Request.QueryString("huoDongId"),"'","''"))
		bbslink = Replace(Request.Form("bbsLink"),"'","''")
		If bbslink = "" Then bbslink="#"
		hd.bbsLink = bbslink
		hd.huoDongName = Replace(Request.Form("huoDongName"),"'","''")
		hd.huoDongRenShu = Replace(Request.Form("huoDongRenShu"),"'","''")
		hd.huoDongXingcheng = Replace(Request.Form("huoDongXingcheng"),"'","''")
		hd.huoDongxieyi = Replace(Request.Form("huoDongxieyi"),"'","''")
		hd.huoDongLingdui = Replace(Request.Form("huoDongLingdui"),"'","''")
		hd.huoDongJifen = Replace(Request.Form("huoDongJifen"),"'","''")
		hd.huoDongZhuangtai = Replace(Request.Form("huoDongZhuangtai"),"'","''")
		hd.huoDongChufaTime = Replace(Request.Form("huoDongChufaTime"),"'","''")
		hd.huoDongFanhuiTime = Replace(Request.Form("huoDongFanhuiTime"),"'","''")
		hd.huoDongInfo = Replace(Request.Form("huoDongInfo"),"'","''")
		aaa = Replace(Request.Form("nkeyword"),"'","''")
		aaa = Replace(aaa,"，",",")
		aaa = Replace(aaa," ",",")
		aaa = Replace(aaa,"|",",")
		hd.nkeyword=keyword_add(aaa)
		hd.updateHuodong()
		set hd = nothing
		Tourl("?Action=ListView")
	End Sub
	
	Public Sub AddAct()
		set hd=new HuoDong
		hd.huoDongName = Replace(Request.Form("huoDongName"),"'","''")
		hd.huoDongRenShu = Replace(Request.Form("huoDongRenShu"),"'","''")
		bbslink = Replace(Request.Form("bbsLink"),"'","''")
		If bbslink = "" Then bbslink="#"
		hd.bbsLink = bbslink
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
		set hd = New HuoDong
		rs = hd.getAllHuoDong()
		if hd.rs.eof then die "暂时没有任何数据"
		hd.rs.PageSize = 15
		Page = CLng(Request("Page"))
	    If Page < 1 Then Page = 1
	    If Page > hd.rs.PageCount Then Page = hd.rs.PageCount
	    hd.rs.AbsolutePage = Page
	%>

<script src="/js/Ajax.js" type="text/javascript"></script>
		<script type="text/javascript">
			function ajaxJiaRen(num,huodongId,id){
				var ha = new QiAjax();
   				ha.ajax_get("baoming.asp?Action=AjaxJiaRen&num="+num+"&huodongId="+huodongId+"&id="+id, call_back);
			}
			function ajaxJianRen(num,huodongId,id){
				var ha = new QiAjax();
   				ha.ajax_get("baoming.asp?Action=AjaxJianRen&num="+num+"&huodongId="+huodongId+"&id="+id, call_backjian);
			}
		   function call_back(text,xml)
		   {
		     if(text!="error"){
		     	//alert(text);
		     	userArr = text.split("|");
				id = userArr[2];
				 $("#renshu_"+id).html(userArr[0]);
			     $("#shengyurenshu_"+id).html(userArr[1]);
		     }else{
				 alert("出错！活动已经结束！不能增减人数！");
			     //$("warning").innerHTML = "<font color=\"#FF0000\">该用户未注册，是否准备帮他<a href='?Action=RegUser'>注册一个</a>?</font>";
		     }
		     
		   }
		   function call_backjian(text,xml){
		   	if(text!="error"){
		     	//alert(text);
		     	userArr = text.split("|");
				id = userArr[2];
				 $("#renshu_"+id).html(userArr[0]);
			     $("#shengyurenshu_"+id).html(userArr[1]);
		     }else{
				 alert("出错！不能再点啦！");
			     //$("warning").innerHTML = "<font color=\"#FF0000\">该用户未注册，是否准备帮他<a href='?Action=RegUser'>注册一个</a>?</font>";
		     }
		   }
		</script>
<body>
<script type="text/javascript" src="js/jquery-1.4.2.min.js"></script>
<link href="images/style.css" rel="stylesheet" type="text/css" id="compStyle"/>
<script type="text/javascript" src="js/simpleMenu_split.js"></script>

<center><h1>活动管理</h1></center>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="39" valign="top"><table width="100%" height="37" border="0" cellpadding="0" cellspacing="1" bgcolor="#DDDDDD">
      <tr bgcolor="#FFFFFF">
        <td width="4%" height="35" align="center">编号</td>
        <td width="30%" align="center">活动名称</td>
        <td width="6%" align="center">活动人数</td>
        <td width="6%" align="center">剩余人数</td>
        <td width="6%" align="center">领队</td>
        <td width="4%" align="center">积分</td>
        <td width="8%" align="center">状态</td>
        <td width="8%" align="center">出发时间</td>
        <td width="10%" align="center">返回时间</td>
        <td width="18%" align="center">操作</td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#DDDDDD">
  <%
  for i=1 to hd.rs.PageSize
  if hd.rs.EOF then Exit For
  %>
	  <tr bgcolor="#FFFFFF">
        <td width="4%" height="25" align="center"><%=hd.rs("huoDongId")%></td>
        <td width="30%" height="25" align="center"><%=hd.rs("huoDongName")%></td>
        <td width="6%" height="25" align="center"><span id="renshu_<%=hd.rs("huoDongId")%>"><%=hd.rs("huoDongRenShu")%></span>
        
        
        <div class="simpleMenu hand" style="width:60px;">
			<div class="simpleMenu_link arrow border white_bg">
				<a href="javascript:;">加减人</a>
			</div>


			<div class="simpleMenu_con" style="width:50px;">
				 <a href="javascript:ajaxJiaRen(1,<%=hd.rs("huoDongId")%>,<%=hd.rs("huoDongId")%>);" onclick=""><span class="icon_add">+1</span></a> <a href="javascript:ajaxJiaRen(5,<%=hd.rs("huoDongId")%>,<%=hd.rs("huoDongId")%>);" onclick=""><span class="icon_add">+5</span></a> <a href="javascript:ajaxJiaRen(10,<%=hd.rs("huoDongId")%>,<%=hd.rs("huoDongId")%>);" onclick=""><span class="icon_add">+10</span></a>
        <a href="javascript:ajaxJianRen(1,<%=hd.rs("huoDongId")%>,<%=hd.rs("huoDongId")%>);" onclick=""><span class="icon_delete">-1</span></a> <a href="javascript:ajaxJianRen(5,<%=hd.rs("huoDongId")%>,<%=hd.rs("huoDongId")%>);" onclick=""><span class="icon_delete">-5</span></a> <a href="javascript:ajaxJianRen(10,<%=hd.rs("huoDongId")%>,<%=hd.rs("huoDongId")%>);" onclick=""><span class="icon_delete">-10</span></a>
			</div>
		</div>

       </td>
        <td width="6%" height="25" align="center"><span id="shengyurenshu_<%=hd.rs("huoDongId")%>"><%=hd.rs("huoDongShengyu")%><span></td>
        <td width="6%" height="25" align="center"><%=hd.rs("huoDongLingdui")%></td>
        <td width="4%" height="25" align="center"><%=hd.rs("huoDongJifen")%></td>
        <td width="8%" height="25" align="center"><%=huoDongZtBackText(hd.rs("huoDongZhuangtai"))%></td>
        <td width="8%" height="25" align="center"><%=hd.rs("huoDongChufaTime")%></td>
        <td width="10%" height="25" align="center"><%=hd.rs("huoDongFanhuiTime")%></td>
        <td width="18%" height="25">
		<select onchange="location.href='?Action=EditZhuangtai&HuodongId=<%=hd.rs("huoDongId")%>&huodongZt='+this.value+'';" name="zhuangtai"><option>选择</option><option value="0">未满</option><option value="1">已满</option><option value="2">出发</option><option value="3">结束</option></select>
		<a href="?Action=EditView&huoDongId=<%=hd.rs("huoDongId")%>">修改</a> <!--| <a href="?Action=DelAct&huoDongId=<%=hd.rs("huoDongId")%>">删除--></a>
		</td>
      </tr>
  <%
  hd.rs.movenext()
  next
  %>
    </table></td>
  </tr>
  <tr>
    <td height="30" align="center" valign="middle" bgcolor="#FFFFFF">
    <%
		if page=1 then
			echo "首页 上一页 第"&page&"页 <a href='huodong.asp?page="&page+1&"'>下一页</a>"
			zuihouyiye = hd.rs.pagecount
			echo "<a href='huodong.asp?page="&zuihouyiye&"'>尾页</a>"
		elseif page=hd.rs.pagecount then
			echo "<a href='huodong.asp?page=1'>首页</a> <a href='huodong.asp?page="&page-1&"'>上一页</a> 第"&page&"页 下一页 尾页"
		else
		zuihouyiye = hd.rs.pagecount
			echo "<a href='huodong.asp?page=1'>首页</a>&nbsp;<a href='huodong.asp?page="&page-1&"'>上一页</a> 第"&page&"页 <a href='huodong.asp?page="&page+1&"'>下一页</a> <a href='huodong.asp?page="&zuihouyiye&"'>尾页</a>"
		end if
			%>
    </td>
  </tr>
</table>
	<%
	hd.rs.close
	set hd.rs=nothing
	set hd = nothing
	End Sub
	
	'#################################################################
	'添加活动的表单提交页面
	'#################################################################
	Public Sub AddView()
	%>
			
			<script type="text/javascript" charset="utf-8" src="kindeditor/kindeditor.js"></script>
			<script type="text/javascript">
			       KE.show({
			           id : 'content_huoDongXingcheng',

					   cssPath : 'kindeditor/index.css',skinType: 'tinymce',
			        items : [
			            'source', 'preview',  'print', 'undo', 'redo', 'cut', 'copy', 'paste',
			            'plainpaste', 'wordpaste', 'justifyleft', 'justifycenter', 'justifyright',
			            'justifyfull', 'insertorderedlist', 'insertunorderedlist', 'indent', 'outdent', 'subscript',
			            'superscript', 'date', 'time', 'specialchar', 'emoticons', 'link', 'unlink', '-',
			            'title', 'fontname', 'fontsize', 'textcolor', 'bgcolor', 'bold',
			            'italic', 'underline', 'strikethrough', 'removeformat', 'selectall', 'image',
			            'flash', 'media', 'layer', 'table', 'hr', 'about'
			        ]		   
			       });
			       KE.show({
			           id : 'content_huoDongxieyi',

					   cssPath : 'kindeditor/index.css',skinType: 'tinymce',
			        items : [
			            'source', 'preview',  'print', 'undo', 'redo', 'cut', 'copy', 'paste',
			            'plainpaste', 'wordpaste', 'justifyleft', 'justifycenter', 'justifyright',
			            'justifyfull', 'insertorderedlist', 'insertunorderedlist', 'indent', 'outdent', 'subscript',
			            'superscript', 'date', 'time', 'specialchar', 'emoticons', 'link', 'unlink', '-',
			            'title', 'fontname', 'fontsize', 'textcolor', 'bgcolor', 'bold',
			            'italic', 'underline', 'strikethrough', 'removeformat', 'selectall', 'image',
			            'flash', 'media', 'layer', 'table', 'hr', 'about'
			        ]		   
			       });
			   </script>
			   <center><h1>添加活动</h1></center>
		<form action="?Action=AddAct" method="post" name="form1" id="form1">
		  <table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999"><tr><td colspan="2" bgcolor="#FFFFFF" ><table width="100%" border="0"><tr><td align="center"><table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999">
		    <tr>
		      <td width="150" bgcolor="#FFFFFF" >活动名称：</td>
		      <td bgcolor="#FFFFFF" align="left" ><input name="huoDongName" type="text" value="" /></td>
		    </tr>
		    <tr>
		      <td width="150" bgcolor="#FFFFFF" >活动人数：</td>
		      <td  align="left" bgcolor="#FFFFFF"><input name="huoDongRenShu" type="text" value="<%=getCookie("huoDongRenShu")%>" />      </td>
		    </tr>
		    <tr>
		      <td width="150" bgcolor="#FFFFFF" >论坛链接：</td>
		      <td  align="left" bgcolor="#FFFFFF"><input name="bbsLink" type="text" value="" />      </td>
		    </tr>
		    <tr>
		      <td bgcolor="#FFFFFF" >活动领队：</td>
		      <td  align="left" bgcolor="#FFFFFF"><input name="huoDongLingdui" type="text" value="<%=getCookie("huoDongLingdui")%>" /> <a href="javascript:void(0);" onclick="document.form1.huoDongLingdui.value='橡皮';">橡皮</a></td>
		    </tr>
		    <tr>
		      <td width="150" bgcolor="#FFFFFF" >活动状态：</td>
		      <td bgcolor="#FFFFFF"  align="left"><input name="huoDongZhuangtai" type="text" value="0" />  <a href="javascript:void(0);" onclick="document.form1.huoDongZhuangtai.value='0';">活动未满</a> | <a href="javascript:void(0);" onclick="document.form1.huoDongZhuangtai.value='1';">活动已满</a> | <a href="javascript:void(0);" onclick="document.form1.huoDongZhuangtai.value='2';">活动出发</a> | <a href="javascript:void(0);" onclick="document.form1.huoDongZhuangtai.value='3';">活动结束</a>   </td>
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
				      <textarea id="content_huoDongXingcheng" name="huoDongXingcheng" style="width:650px;height:400px;visibility:hidden;"></textarea></td>
		    </tr>
		    <tr>
		      <td width="150" bgcolor="#FFFFFF" >本次活动协议：</td>
		      <td bgcolor="#FFFFFF"  align="left"><textarea id="content_huoDongxieyi" name="huoDongxieyi" style="width:650px;height:400px;visibility:hidden;"></textarea></td>
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
		huoDongId = Cint(Request.QueryString("huoDongId"))
		set hd=new HuoDong
		hd.gethuoDong(huoDongId)
	%>
			<script type="text/javascript" charset="utf-8" src="kindeditor/kindeditor.js"></script>
			<script type="text/javascript">
			       KE.show({
			           id : 'content_huoDongXingcheng',

					   cssPath : 'kindeditor/index.css',skinType: 'tinymce',
			        items : [
			            'source', 'preview',  'print', 'undo', 'redo', 'cut', 'copy', 'paste',
			            'plainpaste', 'wordpaste', 'justifyleft', 'justifycenter', 'justifyright',
			            'justifyfull', 'insertorderedlist', 'insertunorderedlist', 'indent', 'outdent', 'subscript',
			            'superscript', 'date', 'time', 'specialchar', 'emoticons', 'link', 'unlink', '-',
			            'title', 'fontname', 'fontsize', 'textcolor', 'bgcolor', 'bold',
			            'italic', 'underline', 'strikethrough', 'removeformat', 'selectall', 'image',
			            'flash', 'media', 'layer', 'table', 'hr', 'about'
			        ]		   
			       });
			       KE.show({
			           id : 'content_huoDongxieyi',

					   cssPath : 'kindeditor/index.css',skinType: 'tinymce',
			        items : [
			            'source', 'preview',  'print', 'undo', 'redo', 'cut', 'copy', 'paste',
			            'plainpaste', 'wordpaste', 'justifyleft', 'justifycenter', 'justifyright',
			            'justifyfull', 'insertorderedlist', 'insertunorderedlist', 'indent', 'outdent', 'subscript',
			            'superscript', 'date', 'time', 'specialchar', 'emoticons', 'link', 'unlink', '-',
			            'title', 'fontname', 'fontsize', 'textcolor', 'bgcolor', 'bold',
			            'italic', 'underline', 'strikethrough', 'removeformat', 'selectall', 'image',
			            'flash', 'media', 'layer', 'table', 'hr', 'about'
			        ]		   
			       });
			   </script>
			<!--修改活动-->
			<center><h1>修改活动</h1></center>
		<form action="?Action=EditAct&huoDongId=<%=huoDongId%>" method="post" name="form1" id="form1">
		  <table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999"><tr><td colspan="2" bgcolor="#FFFFFF" ><table width="100%" border="0"><tr><td align="center"><table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999">
		    <tr>
		      <td width="150" bgcolor="#FFFFFF" >活动名称：</td>
		      <td bgcolor="#FFFFFF" align="left" ><input name="huoDongName" type="text" value="<%=hd.rs("huoDongName")%>" /></td>
		    </tr>
		    <tr>
		      <td width="150" bgcolor="#FFFFFF" >活动人数：</td>
		      <td  align="left" bgcolor="#FFFFFF"><input name="huoDongRenShu" type="text" value="<%=hd.rs("huoDongRenShu")%>" readonly />   (请在列表修改)   </td>
		    </tr>
		    <tr>
		      <td width="150" bgcolor="#FFFFFF" >论坛链接：</td>
		      <td  align="left" bgcolor="#FFFFFF"><input name="bbsLink" type="text" value="<%=hd.rs("bbsLink")%>" />      </td>
		    </tr>
		    <tr>
		      <td bgcolor="#FFFFFF" >活动领队：</td>
		      <td  align="left" bgcolor="#FFFFFF"><input name="huoDongLingdui" type="text" value="<%=hd.rs("huoDongLingdui")%>" /> <a href="javascript:void(0);" onclick="document.form1.huoDongLingdui.value='橡皮';">橡皮</a></td>
		    </tr>
		    <tr>
		      <td width="150" bgcolor="#FFFFFF" >活动状态：</td>
		      <td bgcolor="#FFFFFF"  align="left"><input name="huoDongZhuangtai" type="text" value="<%=hd.rs("huoDongZhuangtai")%>" />  <a href="javascript:void(0);" onclick="document.form1.huoDongZhuangtai.value='0';">活动未满</a> | <a href="javascript:void(0);" onclick="document.form1.huoDongZhuangtai.value='1';">活动已满</a> | <a href="javascript:void(0);" onclick="document.form1.huoDongZhuangtai.value='2';">活动出发</a> | <a href="javascript:void(0);" onclick="document.form1.huoDongZhuangtai.value='3';">活动结束</a>   </td>
		    </tr>
		    <tr>
		      <td width="150" bgcolor="#FFFFFF" >活动积分：</td>
		      <td bgcolor="#FFFFFF"  align="left"><input name="huoDongJifen" type="text" value="<%=hd.rs("huoDongJifen")%>" />      </td>
		    </tr>
		    <tr>
		      <td width="150" bgcolor="#FFFFFF" >出发时间：</td>
		      <td bgcolor="#FFFFFF"  align="left"><input id="d5221" name="huoDongChufaTime" class="Wdate" type="text" value="<%=hd.rs("huoDongChufaTime")%>" onFocus="var d5222=$dp.$('d5222');WdatePicker({onpicked:function(){d5222.focus();},maxDate:'#F{$dp.$D(\'d5222\')}'})"/>  </td>
		    </tr>
		    <tr>
		      <td width="150" bgcolor="#FFFFFF" >返回时间：</td>
		      <td bgcolor="#FFFFFF"  align="left"><input id="d5222" name="huoDongFanhuiTime" class="Wdate" value="<%=hd.rs("huoDongFanhuiTime")%>" type="text" onFocus="WdatePicker({minDate:'#F{$dp.$D(\'d5221\')}'})"/>      </td>
		    </tr>
		      <tr>
			    <td width="150" bgcolor="#FFFFFF">分类简介：</td>
			    <td bgcolor="#FFFFFF"><textarea name="huoDongInfo" cols="60" rows="15" id="huoDongInfo"><%=hd.rs("huoDongInfo")%></textarea></td>
			  </tr>
		    <tr>
		      <td width="150" bgcolor="#FFFFFF" >活动行程安排：</td>
		      <td bgcolor="#FFFFFF"  align="left">
				      <textarea id="content_huoDongXingcheng" name="huoDongXingcheng" style="width:650px;height:400px;visibility:hidden;"><%=cstr(hd.rs("huoDongXingcheng"))%></textarea></td>
		    </tr>
		    <tr>
		      <td width="150" bgcolor="#FFFFFF" >本次活动协议：</td>
		      <td bgcolor="#FFFFFF"  align="left"><textarea id="content_huoDongxieyi" name="huoDongxieyi" style="width:650px;height:400px;visibility:hidden;"><%=cstr(hd.rs("huoDongxieyi"))%></textarea></td>
		    </tr>
	<tr>
    <td width="10%" bgcolor="#FFFFFF">关键字：</td>
    <td bgcolor="#FFFFFF">
    <%
	keyword=split(hd.rs("nkeyword"),",")
	for i=0 to ubound(keyword)
	set rs=server.CreateObject("adodb.recordset")
    set rs.activeconnection=conn
    rs.cursortype=3
    sql="select * from tags where id="&dkh(keyword(i))
    rs.open sql
    kk=kk+","&rs("tag_name")
	next
	kk2=cut(2,kk)
	%>
      <input name="nkeyword" type="text" id="nkeyword" size="50" value="<%=kk2%>" /> <%call gettag()%>
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
	hd.rs.close
	set hd.rs = nothing
	set hd = nothing
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