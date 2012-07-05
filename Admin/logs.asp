<!--#include file="../inc/utf.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/news.asp"-->
<!--#include file="../inc/sys.asp"-->
<!--#include file="../inc/category.asp"-->
<!--#include file="../inc/tfunction.asp"-->
<!--#include file="isadmin.asp"-->
<!--#include file="autotag.asp" -->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<titlt></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="../jquery/jquery.treeview.css" />
	
<script src="../jquery/jquery.js" type="text/javascript"></script>
<script src="../jquery/jquery.cookie.js" type="text/javascript"></script>
<script src="../jquery/jquery.treeview.js" type="text/javascript"></script>
	
<script type="text/javascript">
	$(function() {
	$("#red").treeview();
	});
</script>
<style type="text/css">
<!--
body {
	background-color: #E6EFF7;
	margin-top: 0px;
	margin-bottom: 0px;
}
body,td,th {
	font-size: 12px;
}
a:link {
	color: #1F3A87;
	text-decoration: none;
}
a:visited {
	text-decoration: none;
	color: #1F3A87;
}
a:hover {
	text-decoration: none;
	color: #BC2931;
}
a:active {
	text-decoration: none;
	color: #1F3A87;
}
-->
</style>
<script language="JavaScript">
//检查选择的新闻类别，并执行删除操作
function SelectChk()
{
  var s = false; //用来记录是否存在被选中的复选框
  var Cateid, n=0; 
  var strid, strurl;
  var nn = self.document.all.item("Cate"); //返回复选框Cate的数量
  for (j=0; j<nn.length; j++) {
    if (self.document.all.item("Cate",j).checked) {
      n = n + 1;
      s = true;
      Cateid = self.document.all.item("Cate",j).id+"";  //转换为字符串
      //生成要删除新闻类别编号的列表
      if(n==1) {
        strid = Cateid;
      }
      else {
        strid = strid + "," + Cateid;
      }
    }
  }
  strurl = "right.asp?fl=4&cid=" + strid;
  if(!s) {
    alert("请选择要删除的新闻类别!");
    return false;
  }	
  if (confirm("你确定要删除这些新闻类别吗？")) {
    form1.action = strurl;
    form1.submit();
  }
}
function SelectNews()
{
  var s = false; //用来记录是否存在被选中的复选框
  var Cateid, n=0; 
  var strid, strurl;
  var nn = self.document.all.item("Cate"); //返回复选框Cate的数量
  for (j=0; j<nn.length; j++) {
    if (self.document.all.item("Cate",j).checked) {
      n = n + 1;
      s = true;
      Cateid = self.document.all.item("Cate",j).id+"";  //转换为字符串
      //生成要删除新闻类别编号的列表
      if(n==1) {
        strid = Cateid;
      }
      else {
        strid = strid + "," + Cateid;
      }
    }
  }
  strurl = "right.asp?fl=7&newsid=" + strid;
  if(!s) {
    alert("请选择要删除的新闻!");
    return false;
  }	
  if (confirm("你确定要删除这些新闻吗？")) {
    form1.action = strurl;
    form1.submit();
  }
}
function sltall()
{
	var nn = self.document.all.item("Cate");
	for(j=0;j<nn.length;j++)
	{
		self.document.all.item("Cate",j).checked = true;
	}
}
function sltnull()
{
	var nn = self.document.all.item("Cate");
	for(j=0;j<nn.length;j++)
	{
		self.document.all.item("Cate",j).checked = false;
	}
}

// simonsu 10/1/2009
// showhide 用法showhide(当前,外链id,内链id)
function showhide(linkture,w_link,r_link){
	_obj_checkbox=	document.getElementById(linkture);	
	_obj_w_link=	document.getElementById(w_link);	
	_obj_r_link=	document.getElementById(r_link);	

	//alert(_obj_checkbox.checked);
	if(_obj_checkbox.checked==true){
	_obj_w_link.style.display = '';	
	_obj_r_link.style.display = 'none';
		//alert("显外链");
	}
	else{
	_obj_w_link.style.display = 'none';	
	_obj_r_link.style.display = '';	
	//alert("显内链");
	} 

}


</script>
	
</head>

<body>

<%
sys=request.QueryString("ss")
if sys="" then
sys="x"
end if
select case sys
case "x"
case 0
%>
<table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999"><tr><td bgcolor="#FFFFFF"><table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999">
  <tr>
    <td bgcolor="#FFFFFF">网站路径：
      <%
	  	'检查组件是否被支持
	Function IsObjInstalled(strClassString)
		On Error Resume Next
		Dim xTestObj
		Set xTestObj = Server.CreateObject(strClassString)
		If Err Then
			IsObjInstalled = False
		else
			IsObjInstalled = True
		end if
		Set xTestObj = Nothing
	End Function
	uu=request.ServerVariables("URL")
	vv=len(uu)-15
	response.Write """"&mid(uu,1,vv)&""""
	%>
        <font color="#999999">&nbsp;&nbsp;&nbsp;&nbsp;如果使用中出现路径错误，请修改inc文件夹下的const.asp文件，把install的值改成这个。</font></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">生成静态：
        <%
	if IsObjInstalled("Scripting.FileSystemObject")=false then
	response.Write "<font color=#ff0000>不支持生成静态</font>"
	else
	response.Write "可以生成静态"
	end if
	%>
      <font color="#999999">&nbsp;&nbsp;&nbsp;&nbsp;如果不支持静态请以ASP方式运行本系统。</font></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">支持采集：
        <%
	if IsObjInstalled("Microsoft.XMLHTTP")=false then
	response.Write "<font color=#ff0000>不支持采集功能</font>"
	else
	response.Write "支持采集"
	end if
	%>
      <font color="#999999">&nbsp;&nbsp;&nbsp;&nbsp;下一版本中将会增加采集功能……</font></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">支持数据库：
        <%
	if IsObjInstalled("adodb.recordset")=false then
	response.Write "<font color=#ff0000>不支持数据库连接</font>"
	else
	response.Write "支持数据库连接"
	end if
	%>
      <font color="#999999">&nbsp;&nbsp;&nbsp;&nbsp;如果连这个都不支持就别用这个系统了……</font></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">服务器名称：<%=request.ServerVariables("SERVER_NAME")%></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">操作系统：<%=request.ServerVariables("os")%></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">IIS版本：<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">本文件实际路径：<%=server.mappath(Request.ServerVariables("SCRIPT_NAME"))%></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">QCMS版本：<b><%=Version%></b></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">开发人员：<b>Qesy</b></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><div align="center"><a href="http://qesy.5d6d.com" target="_blank">QCMS网站管理系统</a> 使用探针</div></td>
  </tr>
</table>  <b></b></td>
    
  </tr>
</table>

<%
case 1
set cn=new config
rs=cn.getconfig()
%>
<form action="right.asp?fl=1" method="post" name="form1" id="form1">
  <table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999"><tr><td colspan="2" bgcolor="#FFFFFF" ><table width="100%" border="0"><tr><td align="center"><table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999">
    <tr>
      <td width="150" bgcolor="#FFFFFF" >网站名称：</td>
      <td bgcolor="#FFFFFF" align="left" ><input name="webname" type="text" id="webname" value="<%=cn.rs("webname")%>" /></td>
    </tr>
    <tr>
      <td width="150" bgcolor="#FFFFFF" >网站首页：</td>
      <td  align="left" bgcolor="#FFFFFF"><input name="index" type="text" id="index" value="<%=cn.rs("index")%>" />      </td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF" >首页模版：</td>
      <td  align="left" bgcolor="#FFFFFF"><input name="itemp" type="text" id="itemp" value="<%=cn.rs("itemp")%>" /></td>
    </tr>
    <tr>
      <td width="150" bgcolor="#FFFFFF" >站长邮箱：</td>
      <td bgcolor="#FFFFFF"  align="left"><input name="email" type="text" id="email" value="<%=cn.rs("email")%>" />      </td>
    </tr>
    <tr>
      <td width="150" bgcolor="#FFFFFF" >关键字：</td>
      <td bgcolor="#FFFFFF"  align="left"><input name="keywords" type="text" id="keywords" value="<%=cn.rs("keywords")%>" />      </td>
    </tr>
    <tr>
      <td width="150" bgcolor="#FFFFFF" >网站简介：</td>
      <td bgcolor="#FFFFFF"  align="left"><input name="info" type="text" id="info" value="<%=cn.rs("info")%>" />      </td>
    </tr>
    <tr>
      <td width="150" bgcolor="#FFFFFF" >网站版权：</td>
      <td bgcolor="#FFFFFF"  align="left"><textarea name="copyright" cols="30" rows="4" id="copyright"><%=cn.rs("copyright")%></textarea></td>
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
case 2
%>
<form action="right.asp?fl=2" method="post">
<table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999">
  <tr>
    <td width="150" bgcolor="#FFFFFF">用户帐号：</td>
    <td bgcolor="#FFFFFF">
      <input type="text" name="admin_name" id="admin_name" />
    </td>
  </tr>
  <tr>
    <td width="150" bgcolor="#FFFFFF">用户密码：</td>
    <td bgcolor="#FFFFFF">
      <input type="password" name="admin_password" id="admin_password" />
    </td>
  </tr>
  <tr>
      <td colspan="2" bgcolor="#FFFFFF" ><table width="100%" border="0">
          <tr>
            <td align="center"><input type="submit" value="保存设置" />            </td>
          </tr>
      </table></td>
    </tr>
</table>
</form>
<%
case 3
%>
<form action="right.asp?fl=3" method="post">
<table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999">
  <tr>
    <td width="150" bgcolor="#FFFFFF">用户帐号：</td>
    <td bgcolor="#FFFFFF">
      <font color="#FF0000"><%=session("admin_name")%></font>
      <input name="id" type="hidden" id="id" value="<%=session("id")%>">    </td>
  </tr>
  <tr>
    <td width="150" bgcolor="#FFFFFF">用户密码：</td>
    <td bgcolor="#FFFFFF">
      <input type="password" name="admin_password" id="admin_password" />
    </td>
  </tr>
  <tr>
      <td colspan="2" bgcolor="#FFFFFF" ><table width="100%" border="0">
          <tr>
            <td align="center"><input type="submit" value="保存设置" />            </td>
          </tr>
      </table></td>
    </tr>
</table>
</form>
<%
case 4
set ct=new category
rs=ct.getcategorylist()
%>
<form method="post" name="form1" id="form1">
<table width="100%" border="0" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999">
	<tr>
		<td bgcolor="#FFFFFF"><b>&nbsp;&nbsp;&nbsp;&nbsp;[cid]&nbsp;名称&nbsp;(排序)</b>--------<b>操作&nbsp;&nbsp;选择</b></td>
	</tr>
	<tr>
		<td bgcolor="#FFFFFF"><ul id="red" class="treeview-red"><%MainFl()%></ul></td>
	</tr>
	<tr>
		<td colspan="1" bgcolor="#FFFFFF" >
      <table width="100%" border="0">
          <tr>
            <td align="center"><input type="button" value="全选" onClick="sltall()" />
            <input type="button" value="清空" onClick="sltnull()" />
            <input name="tijiao" type="submit" value="删除" onClick="SelectChk()" /></td>
          </tr>
      </table>
      </td>
    </tr>
</table>
</form>
<%
case 6
cid=request.QueryString("cid")
types=request.QueryString("types")
%>
<script language="JavaScript">
//绑定页面加载完成事件调用函数
window.onload=page_onload;
function page_onload(){
	//showhide();			//在打开页时，就是判断了
	showhide('linkture','w_link','r_link');
}
</script>
<form action="right.asp?fl=6" method="post" name="form2" id="form2">
<table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999">
  <tr>
    <td width="10%" bgcolor="#FFFFFF">分类名称：</td>
    <td bgcolor="#FFFFFF">
      <input name="cname" type="text" id="cname" />
      <select name="types" size="1" id="types">
        <option value="0" selected="selected">新闻分类</option>
        <option value="1">产品展示</option>
      </select>
      <input name="cid" type="hidden" id="cid" value="<%=cid%>" />
      <input name="ons" type="checkbox" id="ons" value="1" />
      显示<input name="linkture" type="checkbox" id="linkture" value="1" onClick="showhide('linkture','w_link','r_link');"/>
      外连
      <input name="px" type="text" id="px" size="2" />
      排序</td>
  </tr>
  
  <tr id="w_link">
    <td width="10%" bgcolor="#FFFFFF">外连地址：</td>
    <td bgcolor="#FFFFFF"><input name="link" type="text" id="link" size="30" /></td>
  </tr>
  <table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999" id="r_link">
  <tr>
    <td bgcolor="#FFFFFF">分类模版：</td>
    <td bgcolor="#FFFFFF">
      <input name="ctemp" type="text" id="ctemp" value="templist\newslist.html" size="50" />
    </td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">内容模版：</td>
    <td bgcolor="#FFFFFF">
      <input name="ntemp" type="text" id="ntemp" value="templist\view.html" size="50" />
    </td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF">分类关键字：</td>
    <td bgcolor="#FFFFFF"><input name="keyword" type="text" id="keyword" size="30" /></td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF">分类简介：</td>
    <td bgcolor="#FFFFFF"><textarea name="info" cols="60" rows="10" id="info"></textarea></td>
  </tr></table>
<tr>
      <td colspan="2" bgcolor="#FFFFFF" ><table width="100%" border="0">
          <tr>
            <td align="center"><input type="submit" value="保存设置" />            </td>
          </tr>
      </table></td>
    </tr>
</table>
</form>
<%
case 7
cid=request.QueryString("cid")
set ns=new news
rs=ns.pagelist(cid)
if ns.rs.eof then
response.Write "暂时没有任何记录！"
else
    ns.rs.PageSize = 20
	Page = CLng(Request("Page"))
    If Page < 1 Then Page = 1
    If Page > ns.rs.PageCount Then Page = ns.rs.PageCount
    ns.rs.AbsolutePage = Page
%>
<form method="post" name="form1" id="form1">
<table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999">
  <%
  for i=1 to ns.rs.PageSize
  if ns.rs.EOF then Exit For
  %>
  <tr>
    <td width="10%" bgcolor="#FFFFFF">&nbsp;&nbsp;<a href="right.asp?ss=7&cid=<%=ns.rs("cid")%>"><%=ns.rs("cname")%></a></td>
    <td width="50%" bgcolor="#FFFFFF">[<%=ns.rs("newsid")%>]&nbsp;&nbsp;<a href="right.asp?ss=23&types=<%=ns.rs("types")%>&newsid=<%=ns.rs("newsid")%>" target="_self"><%=ns.rs("ntitle")%></a></td>
    <td width="20%" bgcolor="#FFFFFF">时间：<%=ns.rs("posttime")%></td>
    <td width="17%" bgcolor="#FFFFFF"><a href="right.asp?ss=23&types=<%=ns.rs("types")%>&newsid=<%=ns.rs("newsid")%>" target="_self">修改</a> | <a href="right.asp?fl=7&newsid=<%=ns.rs("newsid")%>">删除</a> | 
    <%
	if ns.rs("tuijian")=1 then
	response.Write "<a href='right.asp?fl=31&newsid="&ns.rs("newsid")&"' target='_self'><font color='red'>取消</font></a>"
	else
	response.Write "<a href='right.asp?fl=30&newsid="&ns.rs("newsid")&"' target='_self'>推荐</a>"
	end if
	%>
	|    <%
	if ns.rs("huandeng")=1 then
	response.Write "<a href='right.asp?fl=33&newsid="&ns.rs("newsid")&"' target='_self'><font color='red'>取消</font></a>"
	else
	response.Write "<a href='right.asp?fl=32&newsid="&ns.rs("newsid")&"' target='_self'>幻灯</a>"
	end if
	%>
    </td>
    <td width="3%" bgcolor="#FFFFFF"><input name="cate" type="checkbox" id="<%=ns.rs("newsid")%>" /></td>
  </tr>
  <%
  ns.rs.movenext()
  next
  %>
  <tr>
      <td colspan="5" bgcolor="#FFFFFF" >
      <table width="100%" border="0">
      <tr>
            <td align="center">
			<%
			if page=1 then
response.Write "首页&nbsp;上一页&nbsp;第"&page&"页&nbsp;<a href='right.asp?ss="&sys&"&cid="&cid&"&page="&page+1&"'>下一页</a>&nbsp;<a href='right.asp?ss="&sys&"&cid="&cid&"&page="&ns.rs.pagecount&"'>尾页</a>"
elseif page=ns.rs.pagecount then
response.Write "<a href='right.asp?ss="&sys&"&cid="&cid&"&page=1'>首页</a>&nbsp;<a href='right.asp?ss="&sys&"&cid="&cid&"&page="&page-1&"'>上一页</a>&nbsp;第"&page&"页&nbsp;下一页&nbsp;尾页"
else
response.Write "<a href='right.asp?ss="&sys&"&cid="&cid&"&page=1'>首页</a>&nbsp;<a href='right.asp?ss="&sys&"&cid="&cid&"&page="&page-1&"'>上一页</a>&nbsp;第"&page&"页&nbsp;<a href='right.asp?ss="&sys&"&cid="&cid&"&page="&page+1&"'>下一页</a>&nbsp;<a href='right.asp?ss="&sys&"&cid="&cid&"&page="&ns.rs.pagecount&"'>尾页</a>"
end if
			%></td>
          </tr>
          <tr>
            <td align="center"><input type="button" value="全选" onClick="sltall()" />
            <input type="button" value="清空" onClick="sltnull()" />
            <input name="tijiao" type="submit" value="删除" onClick="SelectNews()" />            </td>
          </tr>
      </table>      </td>
    </tr>
</table>
</form>
<%
ns.rs.close
set rs=nothing
end if
%>
<%
case 8
cid=request.QueryString("cid")
set ns=new news
rs=ns.pagelist(cid)
if ns.rs.eof then
response.Write "暂时没有任何记录！"
else
    ns.rs.PageSize = 20
	Page = CLng(Request("Page"))
    If Page < 1 Then Page = 1
    If Page > ns.rs.PageCount Then Page = ns.rs.PageCount
    ns.rs.AbsolutePage = Page
%>
<form method="post" name="form1" id="form1">
<table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999">
  <%
  for i=1 to ns.rs.PageSize
  if ns.rs.EOF then Exit For
  %>
  <tr>
    <td width="10%" bgcolor="#FFFFFF">[<a href="right.asp?ss=7&cid=<%=ns.rs("cid")%>"><%=ns.rs("cid")%></a>]</td>
    <td width="50%" bgcolor="#FFFFFF">编号：<%=ns.rs("newsid")%>&nbsp;&nbsp;<%=ns.rs("ntitle")%></td>
    <td width="20%" bgcolor="#FFFFFF">时间：<%=ns.rs("posttime")%></td>
    <td width="10%" bgcolor="#FFFFFF"><a href="right.asp?ss=23&types=<%=ns.rs("types")%>&newsid=<%=ns.rs("newsid")%>" target="_self">修改</a> | <a href="right.asp?fl=7&newsid=<%=ns.rs("newsid")%>">删除</a> </td>
    <td width="10%" bgcolor="#FFFFFF"><input name="cate" type="checkbox" id="<%=ns.rs("newsid")%>" /></td>
  </tr>
  <%
  ns.rs.movenext()
  next
  %>
  <tr>
      <td colspan="5" bgcolor="#FFFFFF" >
      <table width="100%" border="0">
      <tr>
            <td align="center">
			<%
			if page=1 then
response.Write "首页&nbsp;上一页&nbsp;第"&page&"页&nbsp;<a href='right.asp?ss="&sys&"&cid="&cid&"&page="&page+1&"'>下一页</a>&nbsp;<a href='right.asp?ss="&sys&"&cid="&cid&"&page="&ns.rs.pagecount&"'>尾页</a>"
elseif page=ns.rs.pagecount then
response.Write "<a href='right.asp?ss="&sys&"&cid="&cid&"&page=1'>首页</a>&nbsp;<a href='right.asp?ss="&sys&"&cid="&cid&"&page="&page-1&"'>上一页</a>&nbsp;第"&page&"页&nbsp;下一页&nbsp;尾页"
else
response.Write "<a href='right.asp?ss="&sys&"&cid="&cid&"&page=1'>首页</a>&nbsp;<a href='right.asp?ss="&sys&"&cid="&cid&"&page="&page-1&"'>上一页</a>&nbsp;第"&page&"页&nbsp;<a href='right.asp?ss="&sys&"&cid="&cid&"&page="&page+1&"'>下一页</a>&nbsp;<a href='right.asp?ss="&sys&"&cid="&cid&"&page="&ns.rs.pagecount&"'>尾页</a>"
end if
			%></td>
          </tr>
          <tr>
            <td align="center"><input type="button" value="全选" onClick="sltall()" />
            <input type="button" value="清空" onClick="sltnull()" />
            <input name="tijiao" type="submit" value="删除" onClick="SelectNews()" />            </td>
          </tr>
      </table>      </td>
    </tr>
</table>
</form>
<%
ns.rs.close
set rs=nothing
end if
%>
<%
case 9
types=request.QueryString("types")
cname=request.QueryString("cname")
cid=request.QueryString("cid")
if types=0 then
cp="新闻"
elseif types=1 then
cp="产品"
act = "Products"
'########################################################################################################未完成功能模块
end if
%>
<form action="right.asp?fl=9" method="post" name="jnr" id="jnr">
<table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999">
  <tr>
    <td width="10%" bgcolor="#FFFFFF"><%=cp%>分类：</td>
    <td bgcolor="#FFFFFF">
      <select name="cid" id="cid">
      <option value="<%=cid%>" selected="selected"><%=cname%></option>
      <%
set ct=new category
ct.gettypes(types)
do while not ct.rs.eof
	  %>
      <option value="<%=ct.rs("cid")%>"><%=ct.rs("cname")%></option>
      <%
	  ct.rs.movenext
	  loop
	  %>
    </select></td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF"><%=cp%>标题：</td>
    <td bgcolor="#FFFFFF">
      <input name="ntitle" type="text" id="ntitle" size="50" />
      <input name="posttime" type="hidden" id="posttime" value="<%=now()%>" />
      <input name="types" type="hidden" id="types" value="<%=types%>" />
      <input name="cname" type="hidden" id="cname" value="<%=cname%>" />
    </td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF">简介：</td>
    <td bgcolor="#FFFFFF">
      <input name="ninfo" type="text" id="ninfo" size="50" />
    </td>
  </tr>
  <tr>
    <td width="10%" rowspan="2" bgcolor="#FFFFFF"><%=cp%>小图：</td>
    <td bgcolor="#FFFFFF"><input name="img" type="text" id="img" size="50" /></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><iframe frameborder="0" width="400" height="30" scrolling="no" src="upload.htm" id="ff"></iframe></td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF"><%=cp%>内容：</td>
    <td bgcolor="#FFFFFF"><%  
  Set oFCKeditor = New FCKeditor 
  oFCKeditor.BasePath = "../fckeditor/"  
  oFCKeditor.ToolbarSet = "Default" 
  oFCKeditor.Width = "650" 
  oFCKeditor.Height = "400" 
  oFCKeditor.Value = "" 
  oFCKeditor.Create "ncontent" 
 %>      </td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF">关键字：</td>
    <td bgcolor="#FFFFFF">
      <input name="nkeyword" type="text" id="nkeyword" size="50" /> <%call gettag()%>
    </td>
  </tr>
  <tr>
    <td colspan="2" bgcolor="#FFFFFF"><table width="100%" border="0">
          <tr>
            <td align="center"><input type="submit" value="保存设置" />            </td>
          </tr>
    </table></td>
    </tr>
</table>
</form>
<%
case 10
set cn=new config
rs=cn.guest()
if cn.rs.eof then
response.Write "暂时没有任何记录！"
else
cn.rs.PageSize = 20
	Page = CLng(Request("Page"))
    If Page < 1 Then Page = 1
    If Page > cn.rs.PageCount Then Page = cn.rs.PageCount
    cn.rs.AbsolutePage = Page
%>
<table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999">
<tr>
    <td bgcolor="#FFFFFF">编号：</td>
    <td bgcolor="#FFFFFF">标题：</td>
    <td bgcolor="#FFFFFF">名字：</td>
    <td bgcolor="#FFFFFF">时间：</td>
    <td bgcolor="#FFFFFF">删除&nbsp;&nbsp;回复</td>
    <td bgcolor="#FFFFFF">是否回复</td>
</tr>
  <%
  for i=1 to cn.rs.pagesize
  if cn.rs.eof then exit for
  %>
  <tr>
    <td bgcolor="#FFFFFF"><%=cn.rs("id")%></td>
    <td bgcolor="#FFFFFF"><a href="right.asp?ss=24&id=<%=cn.rs("id")%>"><%=cn.rs("title")%></a></td>
    <td bgcolor="#FFFFFF"><%=cn.rs("name")%></td>
    <td bgcolor="#FFFFFF"><%=cn.rs("times")%></td>
    <td bgcolor="#FFFFFF"><a href="right.asp?fl=10&id=<%=cn.rs("id")%>">删除</a>&nbsp;&nbsp;<a href="right.asp?ss=17&id=<%=cn.rs("id")%>">回复</a></td>
    <td bgcolor="#FFFFFF">
	<%
	if cn.rs("hf")<>"" then
	response.Write "是"
	else
	response.Write "<font color='ff0000'>否</font>"
	end if
	%></td>
  </tr>
  <%
  cn.rs.movenext
  next
  %>
</table>
<%
end if
%>
<%
case 11
set cn=new config
rs=cn.biaoqian()
%>
<table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999">
  <tr>
    <td width="10%" bgcolor="#FFFFFF">编号</td>
    <td bgcolor="#FFFFFF">名称</td>
    <td bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
  <%
  do while not cn.rs.eof
  %>
  <tr>
    <td width="10%" bgcolor="#FFFFFF"><%=cn.rs("id")%></td>
    <td bgcolor="#FFFFFF"><%=cn.rs("names")%></td>
    <td bgcolor="#FFFFFF"><a href="right.asp?ss=25&id=<%=cn.rs("id")%>">修改</a> <a href="right.asp?fl=25&dos=del&id=<%=cn.rs("id")%>">删除</a></td>
  </tr>
  <%
  cn.rs.movenext
  loop
  %>
</table>
<%
case 12
%>
<form action="right.asp?fl=25&dos=add" method="post">
<table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999">
  <tr>
    <td width="10%" bgcolor="#FFFFFF">名称：</td>
    <td bgcolor="#FFFFFF">
      <input name="names" type="text" id="names" />
      
    </td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF">内容：</td>
    <td bgcolor="#FFFFFF"><label><%  
  Set oFCKeditor = New FCKeditor 
  oFCKeditor.BasePath = "../fckeditor/"  
  oFCKeditor.ToolbarSet = "Default" 
  oFCKeditor.Width = "650" 
  oFCKeditor.Height = "400" 
  oFCKeditor.Value = ""
  oFCKeditor.Create "bq" 
 %> 
    </label>
      
    </td>
  </tr>
  <tr>
    <td colspan="2" bgcolor="#FFFFFF"><table width="100%" border="0">
          <tr>
            <td align="center"><input type="submit" value="保存设置" />            </td>
          </tr>
      </table></td>
    </tr>
</table>
</form>
<%
case 13
set cn=new config
rs=cn.diylist()
%>
<table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999">
  <tr>
    <td width="10%" bgcolor="#FFFFFF">编号</td>
    <td bgcolor="#FFFFFF">名称</td>
    <td bgcolor="#FFFFFF">路径</td>
    <td bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
  <%
  do while not cn.rs.eof
  %>
  <tr>
    <td width="10%" bgcolor="#FFFFFF"><%=cn.rs("id")%></td>
    <%
	if html=0 then
	%>
    <td bgcolor="#FFFFFF"><a href="<%=install%>diy.asp?id=<%=cn.rs("id")%>" target="_blank"><%=cn.rs("names")%></a></td>
    <td bgcolor="#FFFFFF"><a href="<%=install%>diy.asp?id=<%=cn.rs("id")%>" target="_blank">/diy.asp?id=<%=cn.rs("id")%></a></td>
    <%
	elseif html=1 then
	%>
    <td bgcolor="#FFFFFF"><a href="<%=install%>diy-<%=cn.rs("id")%>.html" target="_blank"><%=cn.rs("names")%></a></td>
    <td bgcolor="#FFFFFF"><a href="<%=install%>diy-<%=cn.rs("id")%>.html" target="_blank"><%=install%>diy-<%=cn.rs("id")%>.html</a></td>
    <%
	elseif html=2 then
	%>
    <td bgcolor="#FFFFFF"><a href="<%=install%>diy-<%=cn.rs("id")%>.html" target="_blank"><%=cn.rs("names")%></a></td>
    <td bgcolor="#FFFFFF"><a href="<%=install%>diy-<%=cn.rs("id")%>.html" target="_blank"><%=install%>diy-<%=cn.rs("id")%>.html</a></td>
    <%
	end if
	%>
    <td bgcolor="#FFFFFF"><a href="right.asp?ss=26&id=<%=cn.rs("id")%>">修改</a> <a href="right.asp?fl=26&dos=del&id=<%=cn.rs("id")%>">删除</a></td>
  </tr>
  <%
  cn.rs.movenext
  loop
  %>
</table>
<%
case 14
%>
<form action="right.asp?fl=26&dos=add" method="post">
<table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999">
  <tr>
    <td width="10%" bgcolor="#FFFFFF">名称：</td>
    <td bgcolor="#FFFFFF">
      <input name="names" type="text" id="names" />
      
    </td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF">内容：</td>
    <td bgcolor="#FFFFFF"><label>
      <textarea name="diy" cols="80" rows="15" id="diy"></textarea>
    </label></td>
  </tr>
  <tr>
    <td colspan="2" bgcolor="#FFFFFF"><table width="100%" border="0">
          <tr>
            <td align="center"><input type="submit" value="保存设置" />            </td>
          </tr>
      </table></td>
    </tr>
</table>
</form>
<%
case 22
cid=request.QueryString("cid")
set ct=new category
rs=ct.getcategoryinfo(cid)
%>
<script language="JavaScript">
//绑定页面加载完成事件调用函数
window.onload=page_onload;
function page_onload(){
	//showhide();			//在打开页时，就是判断了
	showhide('linkture','w_link','r_link');
}
</script>
<form action="right.asp?fl=22" method="post" name="form2" id="form2">
<table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999">
  <tr>
    <td width="10%" bgcolor="#FFFFFF">分类名称：</td>
    <td bgcolor="#FFFFFF">
      <input name="cname" type="text" id="cname" value="<%=ct.rs("cname")%>" />
      <input name="cid" type="hidden" id="cid" value="<%=ct.rs("cid")%>" />
      <input name="types" type="hidden" id="types" value="<%=ct.rs("types")%>" />
    <%
	if ct.rs("ons")=0 then
	%>
    <input name="ons" type="checkbox" id="ons" value="1" />
    <%
	else
	%>
    <input name="ons" type="checkbox" id="ons" value="1" checked="checked" />
    <%
	end if
	%> 显示<%
	if ct.rs("linkture")=0 then
	%><input name="linkture" type="checkbox" id="linkture" value="1" onClick="showhide('linkture','w_link','r_link');"/>
    <%
	else
	%>
    <input name="linkture" type="checkbox" id="linkture" value="1" checked="checked" onClick="showhide('linkture','w_link','r_link');"/>
    <%
	end if
	%>外连<input name="px" type="text" id="px" value="<%=ct.rs("px")%>" size="2" />
    排序</td>
  </tr>
  <tr id="w_link">
    <td width="10%" bgcolor="#FFFFFF">外连地址：</td>
    <td bgcolor="#FFFFFF"><input name="link" type="text" id="link" value="<%=ct.rs("link")%>" size="30" /></td>
  </tr><table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999" id="r_link">
  <tr>
    <td bgcolor="#FFFFFF">分类模版：</td>
    <td bgcolor="#FFFFFF"><label>
      <input name="ctemp" type="text" id="ctemp" value="<%=ct.rs("ctemp")%>" size="50" />
    </label></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">内容模版：</td>
    <td bgcolor="#FFFFFF"><label>
      <input name="ntemp" type="text" id="ntemp" value="<%=ct.rs("ntemp")%>" size="50" />
    </label></td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF">分类关键字：</td>
    <td bgcolor="#FFFFFF"><input name="keyword" type="text" id="keyword" value="<%=ct.rs("keyword")%>" size="30" /></td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF">分类简介：</td>
    <td bgcolor="#FFFFFF"><textarea name="info" cols="60" rows="10" id="info"><%=ct.rs("info")%></textarea></td>
  </tr></table>
<tr>
      <td colspan="2" bgcolor="#FFFFFF" ><table width="100%" border="0">
          <tr>
            <td align="center"><input type="submit" value="保存设置" />            </td>
          </tr>
      </table></td>
    </tr>
</table>
</form>
<%
case 23
newsid=request.QueryString("newsid")
set ns=new news
rs=ns.getnewsinfo(newsid)
%>
<form action="right.asp?fl=23" method="post" name="jnr" id="jnr">
<table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999">
  <tr>
    <td width="10%" bgcolor="#FFFFFF">分类：</td>
    <td bgcolor="#FFFFFF"><select name="cid" id="cid">
      <option value="<%=ns.rs("cid")%>" selected="selected"><%=ns.rs("cname")%></option>
      <%
types=ns.rs("types")
set ct=new category
ct.gettypes(types)
do while not ct.rs.eof
	  %>
      <option value="<%=ct.rs("cid")%>"><%=ct.rs("cname")%></option>
      <%
	  ct.rs.movenext
	  loop
	  %>
    </select></td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF">标题：</td>
    <td bgcolor="#FFFFFF"><label>
      <input name="ntitle" type="text" id="ntitle" value="<%=ns.rs("ntitle")%>" size="50" />
      <input name="posttime" type="hidden" id="posttime" value="<%=ns.rs("posttime")%>" />
      <input name="newsid" type="hidden" id="newsid" value="<%=ns.rs("newsid")%>" />
    </label></td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF">关键字：</td>
    <td bgcolor="#FFFFFF">
    <%
	keyword=split(ns.rs("nkeyword"),",")
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
    <input name="nkeyword" type="text" id="nkeyword" value="<%=kk2%>" size="50" />
    </td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF">简介：</td>
    <td bgcolor="#FFFFFF"><label>
      <input name="ninfo" type="text" id="ninfo" value="<%=ns.rs("ninfo")%>" size="50" />
    </label></td>
  </tr>
  <tr>
    <td width="10%" rowspan="2" bgcolor="#FFFFFF">缩略图：</td>
    <td bgcolor="#FFFFFF"><label>
      <input name="img" type="text" id="img" value="<%=ns.rs("img")%>" size="50" />
    </label></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><iframe frameborder="0" width="400" height="30" scrolling="no" src="upload.htm" id="ff"></iframe></td>
  </tr>
  
  <tr>
    <td width="10%" bgcolor="#FFFFFF">内容：</td>
    <td bgcolor="#FFFFFF"><%  
  Set oFCKeditor = New FCKeditor 
  oFCKeditor.BasePath = "../fckeditor/"  
  oFCKeditor.ToolbarSet = "Default" 
  oFCKeditor.Width = "650" 
  oFCKeditor.Height = "400" 
  oFCKeditor.Value = ns.rs("ncontent") 
  oFCKeditor.Create "ncontent" 
 %></td>
  </tr>
  <tr>
    <td colspan="2" bgcolor="#FFFFFF"><table width="100%" border="0">
          <tr>
            <td align="center"><input type="submit" value="保存设置" />            </td>
          </tr>
      </table></td>
    </tr>
</table>
</form>
<%
case 24
id=request.QueryString("id")
set cn=new config
rs=cn.getguest(id)
%><table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999">
  <tr>
    <td width="10%" bgcolor="#FFFFFF">标题：</td>
    <td bgcolor="#FFFFFF"><%=cn.rs("title")%></td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF">姓名：</td>
    <td bgcolor="#FFFFFF"><%=cn.rs("name")%></td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF">电话：</td>
    <td bgcolor="#FFFFFF"><%=cn.rs("tel")%></td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF">Q&nbsp;Q：</td>
    <td bgcolor="#FFFFFF"><%=cn.rs("qq")%></td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF">邮箱</td>
    <td bgcolor="#FFFFFF"><%=cn.rs("email")%></td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF">IP地址</td>
    <td bgcolor="#FFFFFF"><%=cn.rs("ip")%></td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF">内容：</td>
    <td bgcolor="#FFFFFF"><%=cn.rs("content")%></td>
  </tr>
    <tr>
    <td width="10%" bgcolor="#FFFFFF"><font color="#FF0000">回复：</font></td>
    <td bgcolor="#FFFFFF"><%=cn.rs("hf")%></td>
  </tr>
  <tr>
    <td colspan="2" bgcolor="#FFFFFF"><table width="100%" border="0">
          <tr>
            <td align="center">[<a href="javascript:history.go(-1);">返回</a>] [<a href="right.asp?fl=10&id=<%=cn.rs("id")%>">删除</a>]</td>
          </tr>
    </table></td>
  </tr>
</table>
<%
case 25
id=request.QueryString("id")
set cn=new config
rs=cn.getbiaoqian(id)
%>
<form action="right.asp?fl=25&dos=edit&id=<%=cn.rs("id")%>" method="post">
<table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999">
  <tr>
    <td width="10%" bgcolor="#FFFFFF">名称：</td>
    <td bgcolor="#FFFFFF">
      <input name="names" type="text" id="names" value="<%=cn.rs("names")%>" />
      
    </td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF">内容：</td>
    <td bgcolor="#FFFFFF"><%  
  Set oFCKeditor = New FCKeditor 
  oFCKeditor.BasePath = "../fckeditor/"  
  oFCKeditor.ToolbarSet = "Default" 
  oFCKeditor.Width = "650" 
  oFCKeditor.Height = "400" 
  oFCKeditor.Value = cn.rs("bq")
  oFCKeditor.Create "bq" 
 %> 
    </td>
  </tr>
  <tr>
    <td colspan="2" bgcolor="#FFFFFF"><table width="100%" border="0">
          <tr>
            <td align="center"><input type="submit" value="保存设置" />            </td>
          </tr>
      </table></td>
    </tr>
</table>
</form>
<%
case 26
id=request.QueryString("id")
set cn=new config
rs=cn.getdiy(id)
%>
<form action="right.asp?fl=26&dos=edit&id=<%=cn.rs("id")%>" method="post">
<table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999">
  <tr>
    <td width="10%" bgcolor="#FFFFFF">名称：</td>
    <td bgcolor="#FFFFFF">
      <input name="names" type="text" id="names" value="<%=cn.rs("names")%>" />
      
    </td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF">内容：</td>
    <td bgcolor="#FFFFFF">
      <textarea name="diy" cols="80" rows="15" id="diy"><%=cn.rs("diy")%></textarea>

    </td>
  </tr>
  <tr>
    <td colspan="2" bgcolor="#FFFFFF"><table width="100%" border="0">
          <tr>
            <td align="center"><input type="submit" value="保存设置" />            </td>
          </tr>
      </table></td>
    </tr>
</table>
</form>
<%
case 15
%>
<table width="300" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999">
  <tr>
    <td width="50%" bgcolor="#FFFFFF">帐号</td>
    <td bgcolor="#FFFFFF">删除</td>
  </tr>
<%
set cn=new config
rs=cn.user()
do while not cn.rs.eof
%>
  <tr>
    <td bgcolor="#FFFFFF"><%=cn.rs("admin_name")%></td>
    <td bgcolor="#FFFFFF"><a href="right.asp?fl=15&id=<%=cn.rs("id")%>">删除</a></td>
  </tr>
  <%
  cn.rs.movenext
  loop
  %>
</table>
<%
case 16
%>
<table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999">
<tr>
<td bgcolor="#FFFFFF">网站地图</td>
</tr>
<form method="post" action="right.asp?fl=16">
<tr>
  <td bgcolor="#FFFFFF">选项</td>
</tr>
<tr>
  <td bgcolor="#FFFFFF">
  <%
  if html=1 then
  %>
  <input name="mode" type="radio" id="radio" value="0" checked="checked" />
    生成首页<input name="mode" type="radio" id="radio" value="1" />
    生成分类页<input name="mode" type="radio" id="radio" value="2" />
    生成内容页<input name="mode" type="radio" id="radio" value="3" />
    生成自定义页
    <%
	end if
	%>
    <input name="mode" type="radio" id="radio" value="4" />
    Google
      <input type="radio" name="mode" id="radio" value="5" />
      Baidu
      <input type="radio" name="mode" id="radio" value="6" />
      RSS
      <input type="radio" name="mode" id="radio" value="7" />
      分类RSS
      </td>
</tr>
<tr>
<td bgcolor="#FFFFFF"><input name="submit" type="submit" value="提交(S)"  class="EasySiteButton" /></td>
</tr></form>
</table>
<%
case 17
id=request.QueryString("id")
%>
<table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999">
  <tr>
    <td bgcolor="#FFFFFF">回复留言：</td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><form id="form3" name="form3" method="post" action="right.asp?fl=17&id=<%=id%>">
        <textarea name="hf" cols="50" rows="4" id="hf"></textarea><br><input name="" type="submit" value="提交" />
    </form>
    </td>
  </tr>
</table>
<%
case 18
set cn=new config
rs=cn.olink_list()
%>
<table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999">
  <tr>
    <td width="10%" bgcolor="#FFFFFF">编号</td>
    <td bgcolor="#FFFFFF">名称</td>
    <td bgcolor="#FFFFFF">路径</td>
    <td bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
  <%
  do while not cn.rs.eof
  %>
  <tr>
    <td width="10%" bgcolor="#FFFFFF"><%=cn.rs("id")%></td>
    <td bgcolor="#FFFFFF"><%=cn.rs("link_name")%></td>
    <td bgcolor="#FFFFFF"><%=cn.rs("link_url")%></td>
    <td bgcolor="#FFFFFF"><a href="right.asp?ss=20&id=<%=cn.rs("id")%>">修改</a> <a href="right.asp?fl=19&dos=del&id=<%=cn.rs("id")%>">删除</a></td>
  </tr>
  <%
  cn.rs.movenext
  loop
  %>
</table>
<%
case 19
%>
<form action="right.asp?fl=19&dos=add" method="post">
<table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999">
  <tr>
    <td width="10%" bgcolor="#FFFFFF">字符：</td>
    <td bgcolor="#FFFFFF"><input name="link_name" type="text" id="link_name" />
    </td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF">地址：</td>
    <td bgcolor="#FFFFFF"><input name="link_url" type="text" id="link_url" value="" size="50" />
    </td>
  </tr>
  <tr>
    <td colspan="2" bgcolor="#FFFFFF"><table width="100%" border="0">
      <tr>
        <td align="center"><input name="submit2" type="submit" value="保存设置" />
        </td>
      </tr>
    </table></td>
  </tr>
</table>
</form>
<%
case 20
link_id=request.QueryString("id")
set cn=new config
rs=cn.getolink(link_id)
%>

<form action="right.asp?fl=19&dos=edit&id=<%=link_id%>" method="post">
<table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999">
  <tr>
    <td width="10%" bgcolor="#FFFFFF">字符：</td>
    <td bgcolor="#FFFFFF">
      <input name="link_name" type="text" id="link_name" value="<%=cn.rs("link_name")%>" />
      
    </td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF">地址：</td>
    <td bgcolor="#FFFFFF">
      <input name="link_url" type="text" id="link_url" value="<%=cn.rs("link_url")%>" size="50" />
    </td>
  </tr>
  <tr>
    <td colspan="2" bgcolor="#FFFFFF"><table width="100%" border="0">
          <tr>
            <td align="center"><input type="submit" value="保存设置" />            </td>
          </tr>
      </table></td>
    </tr>
</table>
</form>
<%
case 21
act=request.QueryString("active")
if act="list" then
set cn=new config
rs=cn.jss_list()
%>
<table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999">
  <tr>
    <td width="10%" bgcolor="#FFFFFF">编号</td>
    <td bgcolor="#FFFFFF">名称</td>
    <td bgcolor="#FFFFFF">调用代码</td>
    <td bgcolor="#FFFFFF">操作</td>
  </tr>
  <%
  do while not cn.rs.eof
  %>
  <tr>
    <td width="10%" bgcolor="#FFFFFF"><%=cn.rs("id")%></td>
    <td bgcolor="#FFFFFF"><%=cn.rs("js_name")%></td>
    <td bgcolor="#FFFFFF"><input name="input" type="text" value="<script language=JavaScript src='http://<%=request.ServerVariables("http_HOST")%><%=install%>inc/js.asp?id=<%=cn.rs("id")%>'></script>" size="70" /></td>
    <td bgcolor="#FFFFFF"><a href="right.asp?ss=21&active=edit&id=<%=cn.rs("id")%>">修改</a> <a href="right.asp?fl=21&dos=del&id=<%=cn.rs("id")%>">删除</a></td>
  </tr>
  <%
  cn.rs.movenext
  loop
  %>
</table>
<%
elseif act="add" then
%>
<form action="right.asp?fl=21&dos=add" method="post">
<table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999">
  <tr>
    <td width="10%" bgcolor="#FFFFFF">字符：</td>
    <td bgcolor="#FFFFFF"><input name="js_name" type="text" id="js_name" />    </td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF">地址：</td>
    <td bgcolor="#FFFFFF"><textarea name="js_code" cols="60" rows="6" id="js_code"></textarea></td>
  </tr>
  <tr>
    <td colspan="2" bgcolor="#FFFFFF"><table width="100%" border="0">
      <tr>
        <td align="center"><input name="submit" type="submit" value="保存设置" />        </td>
      </tr>
    </table></td>
  </tr>
</table>
</form>
<%
elseif act="edit" then
js_id=request.QueryString("id")
set cn=new config
rs=cn.getjss(js_id)
%>
<form action="right.asp?fl=21&dos=edit&id=<%=js_id%>" method="post">
<table width="100%" border="1" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#999999">
  <tr>
    <td width="10%" bgcolor="#FFFFFF">字符：</td>
    <td bgcolor="#FFFFFF">
      <input name="js_name" type="text" id="js_name" value="<%=cn.rs("js_name")%>" />    </td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF">地址：</td>
    <td bgcolor="#FFFFFF"><textarea name="js_code" cols="60" rows="6" id="js_code"><%=js2html(cn.rs("js_code"))%></textarea></td>
  </tr>
  <tr>
    <td colspan="2" bgcolor="#FFFFFF"><table width="100%" border="0">
          <tr>
            <td align="center"><input type="submit" value="保存设置" />            </td>
          </tr>
      </table></td>
    </tr>
</table>
</form>
<%
end if
%>
<%
case 27
%>
<br/><br/><br/><br/><br/><br/><br/><br/>
<center>单页频道</center>
<%
end select
%>
<%
fl=request.QueryString("fl")
select case fl
case 1
set cn=new config
cn.webname=request.Form("webname")
cn.index=request.Form("index")
cn.itemp=request.Form("itemp")
cn.email=request.Form("email")
cn.keywords=request.Form("keywords")
cn.info=request.Form("info")
cn.copyright=request.Form("copyright")
cn.upconfig
xs_err("<font color='red'>保存失败！</font>")
jump=zx_url("保存成功","right.asp?ss=1")
case 2
admin_password=request.Form("admin_password")
set cn=new config
cn.admin_name=request.Form("admin_name")
cn.admin_password=md5(admin_password)
if cn.haveuser(admin_name)=true then
cn.adduser
xs_err("<font color='red'>添加失败！</font>")
jump=zx_url("添加成功","right.asp?ss=2")
else
jump=zx_url("不能添加相同的用户!","right.asp?ss=2")
end if
case 3
admin_password=request.Form("admin_password")
set cn=new config
cn.admin_password=md5(admin_password)
cn.edituser(session("id"))
xs_err("<font color='red'>修改失败！</font>")
jump=zx_url("修改成功!","right.asp?ss=3")
case 4
cid = Request.QueryString("cid")
set ns=new news
If Not ns.HaveCategory(cid) Then
Set ct = New Category
ct.DeleteCategory(cid)
xs_err("<font color='red'>删除失败！</font>")
jump=zx_url("删除分类成功!","right.asp?ss=4")
else
jump=zx_url("分类下有新闻，不能删除!","right.asp?ss=4")
end if
case 22
cid=request.Form("cid")
set ct=new category
ct.cname=request.Form("cname")
ct.linkture=request.Form("linkture")
if ct.linkture="" then ct.linkture=0
ct.link=request.Form("link")
ct.keyword=request.Form("keyword")
ct.info=request.Form("info")
ct.types=request.Form("types")
ct.ons=request.Form("ons")
if ct.ons="" then ct.ons=0
ct.px=request.Form("px")
ct.ctemp=request.Form("ctemp")
ct.ntemp=request.Form("ntemp")
if ct.ctemp="" then
ct.ctemp="\templist\newslist.html"
else
ct.ctemp=ct.ctemp
end if
if ct.ntemp="" then
ct.ntemp="\templist\view.html"
else
ct.ntemp=ct.ntemp
end if
ct.updatecategory(cid)
xs_err("<font color='red'>修改分类失败！</font>")
jump=zx_url("修改分类成功!","right.asp?ss=4")
case 6
set cn=new category
cn.cname=request.Form("cname")
cn.linkture=cint(request.Form("linkture"))
if cn.linkture="" then cn.linkture="0"
cn.link=request.Form("link")
cn.keyword=request.Form("keyword")
cn.info=request.Form("info")
cn.types=cint(request.Form("types"))
cn.ons=request.Form("ons")
cn.px=request.Form("px")
cn.ctemp=request.Form("ctemp")
cn.ntemp=request.Form("ntemp")
if cn.ctemp="" then
cn.ctemp=install&"templist\newslist.html"
else
cn.ctemp=cn.ctemp
end if
if cn.ntemp="" then
cn.ntemp=install&"templist\view.html"
else
cn.ntemp=cn.ntemp
end if
if cn.ons="" then
cn.ons="0"
else
cn.ons=cn.ons
end if
if cn.px="" then
cn.px="1"
else
cn.px=cn.px
end if
cn.pcid=request.Form("cid")
if cn.pcid="" then
cn.insertcategory()
xs_err("<font color='red'>添加分类失败！</font>")
jump=zx_url("添加分类成功!","right.asp?ss=6")
else
cn.sinsertcategory()
xs_err("<font color='red'>添加分类失败！</font>")
jump=zx_url("添加分类成功!","right.asp?ss=6")
end if
case 7
newsid=request.QueryString("newsid")
set ns=new news
ns.deletenews(newsid)
xs_err("<font color='red'>删除新闻失败！</font>")
jump=zx_url("删除新闻成功!","right.asp?ss=4")
case 23
newsid=request.Form("newsid")
set ns=new news
ns.cid=request.Form("cid")
ns.ntitle=request.Form("ntitle")
ns.posttime=request.Form("posttime")
ns.img=request.Form("img")
ns.ncontent=request.Form("ncontent")
aaa = request.Form("nkeyword")
aaa = Replace(aaa,"，",",")
aaa = Replace(aaa," ",",")
aaa = Replace(aaa,"|",",")
ns.nkeyword=keyword_edit(aaa)
ns.ninfo=request.Form("ninfo")
if ns.img="" then
ns.attpic=0
else
ns.attpic=1
end if
ns.updatenews(newsid)
xs_err("<font color='red'>修改新闻失败！</font>")
jump=zx_url("修改新闻成功!","right.asp?ss=7&cid="&cid)
case 9
cname=request.Form("cname")
set ns=new news
ns.cid=request.Form("cid")
ns.types=request.Form("types")
ns.ntitle=request.Form("ntitle")
aaa = request.Form("nkeyword")
aaa = Replace(aaa,"，",",")
aaa = Replace(aaa," ",",")
aaa = Replace(aaa,"|",",")
ns.nkeyword=keyword_add(aaa)
ns.ninfo=request.Form("ninfo")
ns.posttime=request.Form("posttime")
ns.img=request.Form("img")
ns.ncontent=request.Form("ncontent")
if ns.img="" then
ns.attpic=0
else
ns.attpic=1
end if
ns.insertnews()
xs_err("<font color='red'>添加新闻失败！</font>")
jump=zx_url("添加新闻成功!","right.asp?ss=9&types="&ns.types&"&cid="&ns.cid&"&cname="&cname)
case 10
id=request.QueryString("id")
set cn=new config
cn.delguest(id)
xs_err("<font color='red'>删除留言失败！</font>")
jump=zx_url("删除留言成功!","right.asp?ss=10")
case 25
dos=request.QueryString("dos")
id=request.QueryString("id")
set cn=new config
cn.bq=glyh(request.Form("bq"))
cn.names=request.Form("names")
if dos="edit" then
cn.editbiaoqian(id)
xs_err("<font color='red'>修改标签失败！</font>")
jump=zx_url2("修改标签成功!","right.asp?ss=11","window.parent.frames.item('menu').location.reload();")
elseif dos="del" then
cn.delbiaoqian(id)
xs_err("<font color='red'>删除标签失败！</font>")
jump=zx_url2("删除标签成功!","right.asp?ss=11","window.parent.frames.item('menu').location.reload();")
elseif dos="add" then
cn.addbiaoqian()
xs_err("<font color='red'>添加标签失败！</font>")
jump=zx_url2("添加标签成功!","right.asp?ss=11","window.parent.frames.item('menu').location.reload();")
end if
case 15
id=int(request.QueryString("id"))
if id=int(session("id")) then
jump=zx_url("不能删除自己!","right.asp?ss=15")
else
set cn=new config
cn.deluser(id)
xs_err("<font color='red'>删除用户失败！</font>")
jump=zx_url("删除用户成功!","right.asp?ss=15")
end if
case 26
dos=request.QueryString("dos")
id=request.QueryString("id")
set cn=new config
cn.diy=glyh(request.Form("diy"))
cn.names=request.Form("names")
if dos="edit" then
cn.editdiy(id)
xs_err("<font color='red'>修改自定义页面失败！</font>")
jump=zx_url("修改自定义页面成功!","right.asp?ss=13")
elseif dos="del" then
cn.deldiy(id)
xs_err("<font color='red'>删除自定义页面失败！</font>")
jump=zx_url("删除自定义页面成功!","right.asp?ss=13")
elseif dos="add" then
cn.addiy()
xs_err("<font color='red'>添加自定义页面失败！</font>")
jump=zx_url("添加自定义页面成功!","right.asp?ss=13")
end if
case 17
set cn=new config
hid=request.QueryString("id")
ghf=request.Form("hf")
cn.guesthf(hid)
xs_err("<font color='red'>回复失败！</font>")
jump=zx_url("回复成功!","right.asp?ss=10")
case 19
link_do=request.QueryString("dos")
link_id=request.QueryString("id")
set cn=new config
cn.link_name=request.Form("link_name")
cn.link_url=request.Form("link_url")
if link_do="add" then
cn.olink_add()
xs_err("<font color='red'>添加失败！</font>")
jump=zx_url("添加成功!","right.asp?ss=18")
elseif link_do="edit" then
cn.olink_edit(link_id)
xs_err("<font color='red'>修改失败！</font>")
jump=zx_url("修改成功!","right.asp?ss=18")
elseif link_do="del" then
cn.olink_del(link_id)
xs_err("<font color='red'>删除失败！</font>")
jump=zx_url("删除成功!","right.asp?ss=18")
end if
case 21
js_do=request.QueryString("dos")
js_id=request.QueryString("id")
set cn=new config
cn.js_name=request.Form("js_name")
cn.js_code=html2js(request.Form("js_code"))
if js_do="add" then
cn.jss_add()
xs_err("<font color='red'>添加失败！</font>")
jump=zx_url("添加成功!","right.asp?ss=21&active=list")
elseif js_do="edit" then
cn.jss_edit(js_id)
xs_err("<font color='red'>修改失败！</font>")
jump=zx_url("修改成功!","right.asp?ss=21&active=list")
elseif js_do="del" then
cn.jss_del(js_id)
xs_err("<font color='red'>删除失败！</font>")
jump=zx_url("删除成功!","right.asp?ss=21&active=list")
end if
case 30
newsid=request.QueryString("newsid")
set ns=new news
ns.updatecom(newsid)
xs_err("<font color='red'>推荐失败！</font>")
jump=zx_url("推荐成功!","right.asp?ss=7")
case 31
newsid=request.QueryString("newsid")
set ns=new news
ns.updateqx(newsid)
xs_err("<font color='red'>取消推荐文章失败！</font>")
jump=zx_url("取消推荐成功!","right.asp?ss=7")
case 32
newsid=request.QueryString("newsid")
set ns=new news
ns.updatehuandeng(newsid)
xs_err("<font color='red'>设置幻灯失败</font>")
jump=zx_url("设置幻灯成功!","right.asp?ss=7")
case 33
newsid=request.QueryString("newsid")
set ns=new news
ns.updateqxhuandeng(newsid)
xs_err("<font color='red'>取消推荐文章失败！</font>")
jump=zx_url("取消幻灯成功!","right.asp?ss=7")
case 16
act=request.Form("mode")
if act=0 then

set fso = Server.CreateObject("Scripting.FileSystemObject")
set f = fso.CreateTextFile(Server.MapPath("../index.html"),,true)
f.WriteLine(index_temp())
xs_err("<font color='red'>生成首页失败！</font>")
response.Write "<font color='#ff0000'><b>生成首页成功</b></font>"
f.close


elseif act=1 then

set ctl=new category
ctl.getlistdir()
do while not ctl.rs.eof
MakeNewsDir(install&ctl.rs("cid"))
classid=ctl.rs("cid")
pnum=pagenum(install&ctl.rs("ctemp"),classid)
for j =1 to pnum
pagel=j
call htmllist(classid,list_temp(classid),j)
next
response.Write "<font color='ff0000'>分类"&classid&"已生成<br></font>"
response.Flush()
ctl.rs.movenext
loop
response.Write "<font color='ff0000'><b>已经生成全部分类</b></font>"
elseif act=2 then
set nst=new news
nst.allnews()
do while not nst.rs.eof
newsid=nst.rs("newsid")
MakeNewsDir(install&nst.rs("cid"))
zz=htmlnews(nst.rs("cid"),view_temp(newsid),newsid)
nst.rs.movenext
loop
response.Write "<font color='ff0000'><b>全部内容生成完成</b></font>"
elseif act=3 then
set cnt=new config
cnt.diylist()
do while not cnt.rs.eof
zz=htmldiy(diy_temp(cnt.rs("id")),cnt.rs("id"))
cnt.rs.movenext
loop
response.Write "<font color='ff0000'><b>全部自定义页面生成完成</b></font>"
elseif act=4 then
	session("count")=0
	strURL = siteUrl
	
	dim strXML
	strXML = strXML + "<?xml version=""1.0"" encoding=""UTF-8""?>"
	strXML = strXML + "<!--Google Site Map File Generated by " & strURL & " " & RFC822(now,"GMT") & "-->"
	strXML = strXML + "<urlset xmlns=""http://www.google.com/schemas/sitemap/0.84"">"
	strXML = strXML + "<url>"
	strXML = strXML + "<loc>" & strURL & "</loc>"
	strXML = strXML + "</url>"
	session("count")=session("count")+"1"
	
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	set folderFso = fso.GetFolder(server.MapPath("/"))
	set filesFso = folderFso.files
	for each file in filesFso
	strXML = strXML + "<url>"
	strXML = strXML + "<loc>" & strURL & "" & File.Name & "</loc>"
	strXML = strXML + "</url>"
	session("count")=session("count")+"1"
	next
	dim rs,sql
	set rs = server.CreateObject("ADODB.RecordSet")
	sql = "select * from news order by newsid asc"    '修改你要生成的数据表名
	set rs = conn.execute (sql)
	do until rs.eof
	ID=""&rs("newsID")&""
	cid=""&rs("cid")&""
	strXML = strXML + "<url>"
	strXML = strXML + "<loc>" & strURL & "" & cid & "/view-" & ID & ".html</loc>"  '修改为你的文件名称和id
	strXML = strXML + "</url>"
	session("count")=session("count")+"1"
	rs.movenext
	loop
	rs.close
	set rs = nothing
	strXML = strXML + "</urlset>" 
	strXML = "" + strXML + ""
	strXML = "" & strXML & ""
	FolderPath = Server.MapPath(install)
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	Set fout = fso.CreateTextFile(FolderPath&"\google_sitemap.xml")
	fout.writeLine strXML
	fout.close
	set fout = nothing
	conn.close
	set conn = nothing
	
	Function RFC822(byVal myDate, byVal TimeZone)
	Dim myDay, myDays, myMonth, myYear
	Dim myHours, myMinutes, mySeconds
	myDate = CDate(myDate)
	myDay = EnWeekDayName(myDate)
	myDays = Right("00" & Day(myDate),2)
	myMonth = EnMonthName(myDate)
	myYear = Year(myDate)
	myHours = Right("00" & Hour(myDate),2)
	myMinutes = Right("00" & Minute(myDate),2)
	  
	RFC822 = myDay&", "& _
	myDays&" "& _
	myMonth&" "& _ 
	myYear&" "& _
	myHours&":"& _
	myMinutes&":"& _
	mySeconds&" "& _ 
	" " & TimeZone
	End Function
	
	Function EnWeekDayName(InputDate)
	Dim Result
	Select Case WeekDay(InputDate,1)
	Case 1:Result="Sun"
	Case 2:Result="Mon"
	Case 3:Result="Tue"
	Case 4:Result="Wed"
	Case 5:Result="Thu"
	Case 6:Result="Fri"
	Case 7:Result="Sat"
	End Select
	EnWeekDayName = Result
	End Function
	
	Function EnMonthName(InputDate)
	Dim Result
	Select Case Month(InputDate)
	Case 1:Result="Jan"
	Case 2:Result="Feb"
	Case 3:Result="Mar"
	Case 4:Result="Apr"
	Case 5:Result="May"
	Case 6:Result="Jun"
	Case 7:Result="Jul"
	Case 8:Result="Aug"
	Case 9:Result="Sep"
	Case 10:Result="Oct"
	Case 11:Result="Nov"
	Case 12:Result="Dec"
	End Select
	EnMonthName = Result
	End Function
	response.Write "<font color='#ff0000'>GOOGLE地图生成成功</font>"
	elseif act=5 then
	session("count")=0
	strURL = siteUrl
	
	strXML = strXML + "<?xml version=""1.0"" encoding=""UTF-8""?>"
	strXML = strXML + "<document xmlns:bbs=""http://www.baidu.com/search/bbs_sitemap.xsd"">"
	strXML = strXML + "<webSite>" & strURL & "</webSite>"
	strXML = strXML + "<webMaster>" & siteEmail & "</webMaster>"
	strXML = strXML + "<updatePeri>3</updatePeri>"
	strXML = strXML + "<updateTime>" & RFC822(now,"GMT") & "</updateTime>"
	session("count")=session("count")+"1"

	set rs = server.CreateObject("ADODB.RecordSet")
	sql = "select * from news order by newsid asc"    '修改你要生成的数据表名
	set rs = conn.execute (sql)
	do until rs.eof
		ID=""&rs("newsID")&""
		cid=""&rs("cid")&""
		dateandtime=""&rs("posttime")&""
		strXML = strXML + "<item>"
		strXML = strXML + "<link>" & strURL & "" & cid & "/view-" & ID & ".html</link>"  '修改为你的文件名称和id
		strXML = strXML + "<pubDate>" & dateandtime & "</pubDate>"  '修改为你的文件名称和id
		strXML = strXML + "</item>"
		session("count")=session("count")+"1"
	rs.movenext
	loop
	rs.close
	set rs = nothing
	strXML = strXML + "</document>" 
	strXML = "" + strXML + ""
	strXML = "" & strXML & ""
	FolderPath = Server.MapPath(install)
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	Set fout = fso.CreateTextFile(FolderPath&"\baidu_sitemap.xml")
	fout.writeLine strXML
	fout.close
	set fout = nothing
	conn.close
	set conn = nothing
	
	Function RFC822(byVal myDate, byVal TimeZone)
		Dim myDay, myDays, myMonth, myYear
		Dim myHours, myMinutes, mySeconds
		myDate = CDate(myDate)
		myDay = EnWeekDayName(myDate)
		myDays = Right("00" & Day(myDate),2)
		myMonth = EnMonthName(myDate)
		myYear = Year(myDate)
		myHours = Right("00" & Hour(myDate),2)
		myMinutes = Right("00" & Minute(myDate),2)
		  
		RFC822 = myDay&", "& _
		myDays&" "& _
		myMonth&" "& _ 
		myYear&" "& _
		myHours&":"& _
		myMinutes&":"& _
		mySeconds&" "& _ 
		" " & TimeZone
	End Function
	
	Function EnWeekDayName(InputDate)
		Dim Result
		Select Case WeekDay(InputDate,1)
		Case 1:Result="Sun"
		Case 2:Result="Mon"
		Case 3:Result="Tue"
		Case 4:Result="Wed"
		Case 5:Result="Thu"
		Case 6:Result="Fri"
		Case 7:Result="Sat"
		End Select
		EnWeekDayName = Result
	End Function
	
	Function EnMonthName(InputDate)
		Dim Result
		Select Case Month(InputDate)
		Case 1:Result="Jan"
		Case 2:Result="Feb"
		Case 3:Result="Mar"
		Case 4:Result="Apr"
		Case 5:Result="May"
		Case 6:Result="Jun"
		Case 7:Result="Jul"
		Case 8:Result="Aug"
		Case 9:Result="Sep"
		Case 10:Result="Oct"
		Case 11:Result="Nov"
		Case 12:Result="Dec"
		End Select
		EnMonthName = Result
	End Function
	response.Write "<font color='#ff0000'>BAIDU地图生成成功</font>"
	elseif act=6 then
	MakeNewsDir(install&"rss")
	rssall=opentxt(server.mappath(install&"templist/all_rss.txt"),1,x)
    set fso = Server.CreateObject("Scripting.FileSystemObject")
    set f = fso.CreateTextFile(Server.MapPath(install&"rss/0.xml"),true)
    f.WriteLine(rssall)
	response.Write "<a href='"&install&"rss/0.xml' target='_blank'>"&install&"rss/0.xml</a><br>"
    response.Write "<font color='red'><b>生成总RSS成功</b></font>"
    f.close
elseif act=7 then

set ctl=new category
ctl.getlistdir()
do while not ctl.rs.eof
MakeNewsDir(install&ctl.rs("cid"))
classid=ctl.rs("cid")
pnum=1
for j =1 to pnum
pagel=j
call list_rss(classid,opentxt(server.mappath(install&"templist/list_rss.txt"),1,x),j)
next
response.Flush()
ctl.rs.movenext
loop
response.Write "<font color='ff0000'><b>已经生成全部分类XML</b></font>"

end if
%>
<%
end select
%>
<!--频道管理 -->
<%
    '定义第一级分类
    Sub MainFl()
    Dim Rs
	types=0
	set rs=server.CreateObject("adodb.recordset")
	set rs.activeconnection=conn
	rs.cursortype=3
	sql="select * from category where pcid is null"
	rs.open(sql)
	If Not Rs.Eof Then
	Do While Not Rs.Eof
%>
<li>[<%=rs("cid")%>]
	<a href="right.asp?ss=22&cid=<%=rs("cid")%>" title="链接 right.asp?ss=22&cid=<%=rs("cid")%>&ensp;&ensp;" target="_blank"><%=rs("cname")%></a><font color='red'>(<%=rs("px")%>)</font>
<% if rs("linkture")=1 then %>
    --------<font color='cccccc'>添加子分类 | 添加内容</font> | <a href="right.asp?ss=22&cid=<%=rs("cid")%>" target="_self">分类设置</a> | <a href="right.asp?fl=4&cid=<%=rs("cid")%>">删除分类</a>
<% else %>
	--------<a href="right.asp?ss=6&types=<%=rs("types")%>&cid=<%=rs("cid")%>&ctemp=<%=rs("ctemp")%>&ntemp=<%=rs("ntemp")%>">添加子分类</a> | <a href="right.asp?ss=9&types=<%=rs("types")%>&cid=<%=rs("cid")%>&cname=<%=rs("cname")%>">添加内容</a> | <a href="right.asp?ss=22&cid=<%=rs("cid")%>" target="_self">分类设置</a> | <a href="right.asp?fl=4&cid=<%=rs("cid")%>">删除分类</a>
<% End if %>
	<input style="vertical-align:middle;height: 14px;" name="Cate" type="checkbox" id="<%=rs("cid")%>" />
</li>
<%
	Call Subfl(Rs("cid")) '循环子级分类                    
	Rs.MoveNext
	Loop
	End If
	Set Rs = Nothing
        End Sub
        '定义子级分类
        Sub SubFl(FID)
	set rs1=server.CreateObject("adodb.recordset")
	set rs1.activeconnection=conn
	rs1.cursortype=3
	Set Rs1 = Conn.Execute("SELECT * FROM category WHERE pcid = " & FID )
	If Not Rs1.Eof Then
%>
	<ul>
<%
	Do While Not Rs1.Eof
%>
<li>[<%=rs1("cid")%>]
    <a href="right.asp?ss=22&cid=<%=rs1("cid")%>" title="链接 right.asp?ss=22&cid=<%=rs1("cid")%>&ensp;&ensp;" target="_blank"><%=rs1("cname")%></a><font color='green'>(<%=rs1("px")%>)</font>
    --------<a href="right.asp?ss=6&types=<%=rs1("types")%>&cid=<%=rs1("cid")%>">添加子分类</a> | <a href="right.asp?ss=9&types=<%=rs1("types")%>&cid=<%=rs1("cid")%>&cname=<%=rs1("cname")%>">添加内容</a> | <a href="right.asp?ss=22&cid=<%=rs1("cid")%>" target="_self">分类设置</a> | <a href="right.asp?fl=4&cid=<%=rs1("cid")%>">删除分类</a>
    <input style="vertical-align:middle;height: 14px;" name="Cate" type="checkbox" id="<%=rs1("cid")%>" />
</li>
<%
	Call SubFl(Trim(Rs1("cid"))) '递归子级分类
	Rs1.Movenext:Loop
%>
	</ul>
<% 
	If Rs1.Eof Then
	Rs1.Close
	Exit Sub
	End If
	End If
	Set Rs1 = Nothing
  End Sub

        '最后直接调用MainFl()就行了
'生成列表
sub htmllist(tcid,tstr,j)
set fso = Server.CreateObject("Scripting.FileSystemObject")
set f = fso.CreateTextFile(Server.MapPath(install&tcid&"/list-"&j&".html"),,true)
f.WriteLine(tstr)
f.close
response.Write "正在生成../"&tcid&"/list-"&j&".html<br>"
response.Flush()
end sub

'生成列表XML
sub list_rss(tcid,tstr,j)
set fso = Server.CreateObject("Scripting.FileSystemObject")
set f = fso.CreateTextFile(Server.MapPath(install&"rss/"&tcid&".xml"),,true)
f.WriteLine(tstr)
f.close
response.Write "<a href='"&install&"rss/"&tcid&".xml' target='_blank'>"&request.ServerVariables("HTTP_HOST")&install&"rss/"&tcid&".xml</a><br>"
response.Flush()
end sub

'生成内容
function htmlnews(tcid,tstr,j)
set fso = Server.CreateObject("Scripting.FileSystemObject")
set f = fso.CreateTextFile(Server.MapPath(install&tcid&"/view-"&j&".html"),,true)
f.WriteLine(tstr)
f.close
response.Write install&tcid&"/view-"&j&".html已生成<br>"
response.Flush()
end function

'获取页数
function pagenum(code,ccid)
set fso=server.createobject("scripting.filesystemobject") 
set mytext=fso.OpenTextFile(server.mappath(code)) 
str=mytext.readall
mytext.close
set reg=new regexp
reg.IgnoreCase = False 
reg.Global = True 
reg.Pattern = "{pagelist([\s\S.]*?)}([\s\S.]*?){pageend}"
set aa=reg.execute(str)
for each match in aa
bq=match.submatches(0)
size=nbq(bq,"size")
if size="" then
size=20
else
size=size
end if
next
set ns=new news
rs=ns.pagelist(ccid)
if not ns.rs.eof then
ns.rs.pagesize=size
pagenum=ns.rs.pagecount
else
pagenum=1
end if
end Function

'生成内容
function htmldiy(tstr,k)
set fso = Server.CreateObject("Scripting.FileSystemObject")
set f = fso.CreateTextFile(Server.MapPath(install&"diy-"&k&".html"),,true)
f.WriteLine(tstr)
f.close
response.Write install&"diy-"&k&".html已生成<br>"
response.Flush()
end function


 %>
</body>
</html>