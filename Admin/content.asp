<!--#include file="../inc/utf.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/news.asp"-->
<!--#include file="../inc/sys.asp"-->
<!--#include file="../inc/category.asp"-->
<!--#include file="../inc/tfunction.asp"-->
<!--#include file="isadmin.asp"-->
<html>
<head>
<link rel=stylesheet href="styles/advanced/style.css" />
<meta http-equiv="Content-Type" content="text/html; charset=<%=q_Charset%>">
<script type="text/javascript">
<!--
function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->
</script>
</head>
<body>

<%
select case request.QueryString("action")
case 32
newsid=request.QueryString("newsid")
set ns=new news
ns.updatehuandeng(newsid)
xs_err("<font color='red'>设置幻灯失败</font>")
jump=zx_url("设置幻灯成功!","content.asp?active=list")
set ns = nothing
case 33
newsid=request.QueryString("newsid")
set ns=new news
ns.updateqxhuandeng(newsid)
xs_err("<font color='red'>取消推荐文章失败！</font>")
jump=zx_url("取消幻灯成功!","content.asp?active=list")
set ns = nothing
end select

act=request.QueryString("active")
if act="list" Then
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">内容管理</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<table width="100%" border="0" cellspacing="0" cellpadding="0" class="table">
  <tr>
    <td height="40" width="70%"><form name="form" id="form">
      <select name="jumpMenu" id="jumpMenu" onChange="MM_jumpMenu('self',this,0)" class="kuangy">
      <option value="">分类导航</option>
<%call jumpfl()%>
      </select>
    </form></td>
    <td width="30%"><form name="form2" method="get" action="content.asp">
      <input type="text" name="keyword" id="keyword" class="kuangy">&nbsp;&nbsp;<input name="" type="submit" value="搜索">
      <input name="active" type="hidden" id="active" value="search">
    </form></td>
  </tr>
</table>

</div>
<br />
<br />
<%
cid=request.QueryString("cid")
set ns=new news
rs=ns.pagelist(cid)
If Not ns.rs.eof then
    ns.rs.PageSize = 20
	Page = CLng(Request("Page"))
    If Page < 1 Then Page = 1
    If Page > ns.rs.PageCount Then Page = ns.rs.PageCount
    ns.rs.AbsolutePage = Page
%>
<script language="JavaScript">
//检查选择的新闻，并执行批量设置
function EditNews()
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
  strurl = "content.asp?active=edit_all&id=" + strid;
  if(!s) {
    alert("请选择要设置的新闻!");
    return false;
  }	
  if (confirm("你确定要设置这些新闻吗？")) {
    form1.action = strurl;
    form1.submit();
  }
}
//检查选择的新闻类别，并执行转移操作
function MoveNews()
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
  strurl = "content.asp?active=move&id=" + strid;
  if(!s) {
    alert("请选择要移动的新闻!");
    return false;
  }	
  if (confirm("你确定要移动这些新闻吗？")) {
    form1.action = strurl;
    form1.submit();
  }
}
//检查选择的新闻类别，并执行删除操作
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
  strurl = "content.asp?cz=del&id=" + strid;
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
</script>


<div id="divMain2">
<form method="post" name="form1" id="form1">
<table width="100%" border="0" cellspacing="0" class="table">
<tr bgcolor="#DAE2E8"><td>选择</td><td>标题</td><td>操作</td><td>时间</td></tr>
  <%
  for i=1 to ns.rs.PageSize
  if ns.rs.EOF then Exit For
  %>
  <%
  if i mod 2=0 then
  %>
   <tr bgcolor="#F1F4F7">
   <%else%>
  <tr>
  <%end if%>
    <td width="5%"><input name="cate" type="checkbox" id="<%=ns.rs("newsid")%>" /></td>
    <td width="50%">[&nbsp;<a href="content.asp?active=list&cid=<%=ns.rs("news.cid")%>"><%=ns.rs("cname")%></a>&nbsp;]&nbsp;&nbsp;<%=ns.rs("ntitle")%>&nbsp;&nbsp;
	<%
	if ns.rs("tuijian")=1 then response.Write("[<font color='#ff0000'>荐</font>]")
	%></td>
    
    <td width="30%"><a href="content.asp?active=edit&id=<%=ns.rs("newsid")%>&cid=<%=ns.rs("news.cid")%>" target="_self">修改</a> | <a href="content.asp?cz=del&id=<%=ns.rs("newsid")%>&cid=<%=ns.rs("news.cid")%>">删除</a> | <%
	if ns.rs("huandeng")=1 then
	response.Write "<a href='?action=33&newsid="&ns.rs("newsid")&"' target='_self'><font color='red'>取消</font></a>"
	else
	response.Write "<a href='?action=32&newsid="&ns.rs("newsid")&"' target='_self'>幻灯</a>"
	end if
	%> 
    <%
	if ns.rs("tuijian")=1 then
	%>
    <a href="content.asp?cz=ntuijian&id=<%=ns.rs("newsid")%>&cid=<%=ns.rs("news.cid")%>" target="_self">取消</a>
    <%else%>
	    <a href="content.asp?cz=tuijian&id=<%=ns.rs("newsid")%>&cid=<%=ns.rs("news.cid")%>" target="_self">推荐</a>
    <%end if%>
    </td><td width="15%">时间：<%=formatdatetime(ns.rs("posttime"),2)%></td>
  </tr>
  <%
  ns.rs.movenext()
  next
  %>
  <tr>
      <td colspan="4" align="center"><%
			if page=1 then
response.Write "首页&nbsp;上一页&nbsp;第"&page&"页&nbsp;<a href='content.asp?active=list&cid="&cid&"&page="&page+1&"'>下一页</a>&nbsp;<a href='content.asp?active=list&cid="&cid&"&page="&ns.rs.pagecount&"'>尾页</a>"
elseif page=ns.rs.pagecount then
response.Write "<a href='content.asp?active=list&cid="&cid&"&page=1'>首页</a>&nbsp;<a href='content.asp?active=list&cid="&cid&"&page="&page-1&"'>上一页</a>&nbsp;第"&page&"页&nbsp;下一页&nbsp;尾页"
else
response.Write "<a href='content.asp?active=list&cid="&cid&"&page=1'>首页</a>&nbsp;<a href='content.asp?active=list&cid="&cid&"&page="&page-1&"'>上一页</a>&nbsp;第"&page&"页&nbsp;<a href='content.asp?active=list&cid="&cid&"&page="&page+1&"'>下一页</a>&nbsp;<a href='content.asp?active=list&cid="&cid&"&page="&ns.rs.pagecount&"'>尾页</a>"
end if
			%></td>
    </tr>
<tr><td align="center" colspan="4" height="40"><input type="button" value=" 全 选 " onClick="sltall()" style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input type="button" value=" 清 空 " onClick="sltnull()" style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input name="tijiao" type="submit" value=" 删 除 " onClick="SelectNews()" style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input name="tijiao2" type="submit" value=" 移 动 " onClick="MoveNews()" style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input name="tijiao4" type="submit" value=" 设 置 " onClick="EditNews()" style="border:1px #000000 solid;vertical-align:middle;height:25px"/></td></tr>
</table>
</form></div>
<%
Else
%>
<table width="100%" border="0" cellspacing="0" class="table">
<tr bgcolor="#F1F4F7"><td height="40">暂时无任何内容！</td></tr></table>
<%
End if
ns.rs.close
set rs=nothing
%>
<%
elseif act="search" then
skey=request.QueryString("keyword")
set ns=new news
rs=ns.news_search(skey)
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">内容管理</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<table width="100%" border="0" cellspacing="0" cellpadding="0" class="table">
  <tr>
    <td height="40" width="70%"><form name="form" id="form">
      <select name="jumpMenu" id="jumpMenu" onChange="MM_jumpMenu('self',this,0)" class="kuangy">
      <option value="">分类导航</option>
<%call jumpfl()%>
      </select>
    </form></td>
    <td width="30%"><form name="form2" method="get" action="content.asp?active=search">
      <input type="text" name="keyword" id="keyword" class="kuangy">&nbsp;&nbsp;<input name="" type="submit" value="搜索">
      <input name="active" type="hidden" id="active" value="search">
    </form></td>
  </tr>
</table>

</div>
<br />
<br />
<%
If Not ns.rs.eof then
    ns.rs.PageSize = 20
	Page = CLng(Request("Page"))
    If Page < 1 Then Page = 1
    If Page > ns.rs.PageCount Then Page = ns.rs.PageCount
    ns.rs.AbsolutePage = Page
%>
<script language="JavaScript">
//检查选择的新闻，并执行批量设置
function EditNews()
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
  strurl = "content.asp?active=edit_all&id=" + strid;
  if(!s) {
    alert("请选择要设置的新闻!");
    return false;
  }	
  if (confirm("你确定要设置这些新闻吗？")) {
    form1.action = strurl;
    form1.submit();
  }
}
//检查选择的新闻类别，并执行转移操作
function MoveNews()
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
  strurl = "content.asp?active=move&id=" + strid;
  if(!s) {
    alert("请选择要移动的新闻!");
    return false;
  }	
  if (confirm("你确定要移动这些新闻吗？")) {
    form1.action = strurl;
    form1.submit();
  }
}
//检查选择的新闻类别，并执行删除操作
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
  strurl = "content.asp?cz=del&id=" + strid;
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
</script>

<div id="divMain2">
<form method="post" name="form1" id="form1">
<table width="100%" border="0" cellspacing="0" class="table">
<tr bgcolor="#DAE2E8"><td>选择</td><td>标题</td><td>操作</td><td>时间</td></tr>
  <%
  for i=1 to ns.rs.PageSize
  if ns.rs.EOF then Exit For
  %>
  <%
  if i mod 2=0 then
  %>
   <tr bgcolor="#F1F4F7">
   <%else%>
  <tr>
  <%end if%>
    <td width="5%"><input name="cate" type="checkbox" id="<%=ns.rs("newsid")%>" /></td>
    <td width="50%">[&nbsp;<a href="content.asp?active=list&cid=<%=ns.rs("news.cid")%>"><%=ns.rs("cname")%></a>&nbsp;]&nbsp;&nbsp;<%=ns.rs("ntitle")%>&nbsp;&nbsp;
	<%
	if ns.rs("tuijian")=1 then response.Write("[<font color='#ff0000'>荐</font>]")
	%></td>
    
    <td width="30%"><a href="content.asp?active=edit&id=<%=ns.rs("newsid")%>&cid=<%=ns.rs("news.cid")%>" target="_self">修改</a> | <a href="content.asp?cz=del&id=<%=ns.rs("newsid")%>&cid=<%=ns.rs("news.cid")%>">删除</a> | 
    <%
	if ns.rs("tuijian")=1 then
	%>
	    <a href="content.asp?cz=ntuijian&id=<%=ns.rs("newsid")%>&cid=<%=ns.rs("news.cid")%>" target="_self">取消</a>
    <%else%>
    <a href="content.asp?cz=tuijian&id=<%=ns.rs("newsid")%>&cid=<%=ns.rs("news.cid")%>" target="_self">推荐</a>
    <%end if%>
    </td><td width="15%">时间：<%=formatdatetime(ns.rs("posttime"),2)%></td>
  </tr>
  <%
  ns.rs.movenext()
  next
  %>
  <tr>
      <td colspan="4" align="center"><%
			if page=1 then
response.Write "首页&nbsp;上一页&nbsp;第"&page&"页&nbsp;<a href='content.asp?active=search&cid="&cid&"&page="&page+1&"&keyword="&skey&"'>下一页</a>&nbsp;<a href='content.asp?active=search&cid="&cid&"&page="&ns.rs.pagecount&"&keyword="&skey&"'>尾页</a>"
elseif page=ns.rs.pagecount then
response.Write "<a href='content.asp?active=search&cid="&cid&"&page=1&keyword="&skey&"'>首页</a>&nbsp;<a href='content.asp?active=search&cid="&cid&"&page="&page-1&"&keyword="&skey&"'>上一页</a>&nbsp;第"&page&"页&nbsp;下一页&nbsp;尾页"
else
response.Write "<a href='content.asp?active=search&cid="&cid&"&page=1&keyword="&skey&"'>首页</a>&nbsp;<a href='content.asp?active=search&cid="&cid&"&page="&page-1&"&keyword="&skey&"'>上一页</a>&nbsp;第"&page&"页&nbsp;<a href='content.asp?active=search&cid="&cid&"&page="&page+1&"&keyword="&skey&"'>下一页</a>&nbsp;<a href='content.asp?active=search&cid="&cid&"&page="&ns.rs.pagecount&"&keyword="&skey&"'>尾页</a>"
end if
			%></td>
    </tr>
<tr><td align="center" colspan="4" height="40"><input type="button" value=" 全 选 " onClick="sltall()" style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input type="button" value=" 清 空 " onClick="sltnull()" style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input name="tijiao" type="submit" value=" 删 除 " onClick="SelectNews()" style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input name="tijiao2" type="submit" value=" 移 动 " onClick="MoveNews()" style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input name="tijiao3" type="submit" value=" 设 置 " onClick="EditNews()" style="border:1px #000000 solid;vertical-align:middle;height:25px"/></td></tr>
</table>
</form></div>
<%
Else
%>
<table width="100%" border="0" cellspacing="0" class="table">
<tr bgcolor="#F1F4F7"><td height="40">暂时无任何内容！</td></tr></table>
<%
End if
ns.rs.close
set rs=nothing
%>
<%elseif act="add" then
qesy=int(request.QueryString("cid"))
%>
<script type="text/javascript" charset="utf-8" src="kindeditor/kindeditor.js"></script>
<script type="text/javascript">

       KE.show({
           id : 'content_1',

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
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">添加内容</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<form action="content.asp?cz=add" method="post" name="jnr" id="jnr">
<table width="100%" border="0" cellspacing="0" class="table">
  <tr bgcolor="#F1F4F7">
    <td width="10%" height="40">分类：</td>
    <td>
      <select name="cid" id="cid" class="kuangy">
<%call addfl(qesy)%>
    </select></td>
  </tr>
  <tr>
    <td width="10%" height="40">标题：</td>
    <td>
      <input name="ntitle" type="text" id="ntitle" class="kuang" onBlur="this.className='kuang'" onFocus="this.className='kuang1'" />
      <input name="posttime" type="hidden" id="posttime" value="<%=now()%>"  /></td>
  </tr>
  <tr bgcolor="#F1F4F7">
    <td width="10%" height="40">关键字：</td>
    <td><input name="nkeyword" type="text" id="nkeyword" class="kuang" onBlur="this.className='kuang'" onFocus="this.className='kuang1'" /></td>
  </tr>
  <tr>
    <td width="10%" height="40">简介：</td>
    <td>
      <input name="ninfo" type="text" id="ninfo" class="kuang" onBlur="this.className='kuang'" onFocus="this.className='kuang1'" />
    </td>
  </tr>
  <tr bgcolor="#F1F4F7">
    <td width="10%" height="60">小图：</td>
    <td><input name="img" type="text" id="img" class="kuang" onBlur="this.className='kuang'" onFocus="this.className='kuang1'" /><br>
      <iframe frameborder="0" width="600" height="25" scrolling="no" src="upload.htm.asp" id="ff"></iframe></td>
  </tr>
  <tr>
    <td width="10%">内容：</td>
    <td><div id="cont">
      <textarea id="content_1" name="ncontent" style="width:650px;height:300px;visibility:hidden;"></textarea>
      </div></td>
  </tr>
  <tr>
    <td colspan="2" align="center" height="40"><input type="submit" value=" 保 存 " style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;

        <input type="reset" name="button" id="button" value=" 重 置 " style="border:1px #000000 solid;vertical-align:middle;height:25px">
</td>
    </tr>
</table>
</form>
</div>
<%
ElseIf act="Replace" Then
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">批量替第</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2"><form action="content.asp?cz=replace" method="post">
<table width="100%" border="0" cellspacing="0" class="table">
  <tr bgcolor="#F1F4F7">
    <td width="10%" height="40">替换选择：</td>
    <td><select name="tixiang" id="tixiang" class="kuangy">
      <option value="title">标题</option>
      <option value="content">内容</option>
    </select>
    </td>
  </tr>
  <tr>
    <td width="10%" height="40">将字符：</td>
    <td><input name="aa" type="text" class="kuang" id="ntitle" onFocus="this.className='kuang1'" onBlur="this.className='kuang'" >*替换前的内容</td>
  </tr>
  <tr bgcolor="#F1F4F7">
    <td width="10%" height="40">替换成：</td>
    <td><input name="bb" type="text" class="kuang" id="bb" onFocus="this.className='kuang1'" onBlur="this.className='kuang'" >*替换后的内容</td>
  </tr>
  <tr>
    <td height="40" colspan="2" align="center"><input type="submit" value=" 保 存 " style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;

        <input type="reset" name="button" id="button" value=" 重 置 " style="border:1px #000000 solid;vertical-align:middle;height:25px"></td>
    </tr>
</table>
</form>
</div>
<%
elseif act="hiRep" then
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">高级替换</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<%
step=request.QueryString("step")
if step="" or step=1 then
%>
<form action="content.asp?active=hiRep&step=2" method="post"><table width="100%" border="0" cellspacing="0" class="table">
  <tr bgcolor="#F1F4F7">
    <td width="10%" height="40">表名选择：</td>
    <td>
    <select name="biao" id="biao" class="kuangy">
    <%Set rs=Conn.OpenSchema(20)
      While not rs.EOF
         
if rs(3)="TABLE" then
response.Write "<option value='"&rs(2)&"'>"&rs(2)&"</option>"
end if
         rs.MoveNext
      Wend
	  %>
      
    </select>
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" value=" 下一步 " style="border:1px #000000 solid;vertical-align:middle;height:25px"/></td>
  </tr></table></form>
  <%
  elseif step=2 then
  biao=request.Form("biao")
  %>
  <form action="content.asp?cz=hirep" method="post"><table width="100%" border="0" cellspacing="0" class="table">
  <tr bgcolor="#F1F4F7">
    <td width="10%" height="40">所选表名：</td>
    <td>
    <select name="biao" id="biao" class="kuangy">

<option value="<%=biao%>"><%=biao%></option>

      
    </select>
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  </tr>
  <tr>
    <td height="40">字段选择：</td>
    <td><select name="ziduan" id="ziduan" class="kuangy">
      <%
       Set rs=Server.CreateObject("ADODB.Recordset") 
       Sql="select * from "&biao&" where 1<>1"
       rs.open sql,Conn,1,1
       j=rs.Fields.count
       For i=0 to (j-1) 
	  Response.Write "<option value='"&rs.Fields(i).Name&"'>"&rs.Fields(i).Name&"</option>"
       Next
     %>
    </select></td>
  </tr>
  <tr bgcolor="#F1F4F7">
    <td height="40">将字符：</td>
    <td><input name="aa" type="text" class="kuang" id="aa" onFocus="this.className='kuang1'" onBlur="this.className='kuang'" >*替换前的内容</td>
  </tr>
  <tr>
    <td height="40">替换成：</td>
    <td><input name="bb" type="text" class="kuang" id="bb" onFocus="this.className='kuang1'" onBlur="this.className='kuang'" >*替换后的内容</td>
  </tr>
  <tr bgcolor="#F1F4F7">
    <td height="40" colspan="2" align="center"><input type="submit" value=" 替 换 " style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" name="button" value=" 上一步 " style="border:1px #000000 solid;vertical-align:middle;height:25px"  onclick="history.go(-1)" />
</td>
    </tr>
  </table>
  </form>
  <%end if%>
</div>
<%
elseif act="edit" then
qesy=int(request.QueryString("cid"))
newsid=request.QueryString("id")
set ns=new news
rs=ns.getnewsinfo(newsid)
%>
<script type="text/javascript" charset="utf-8" src="kindeditor/kindeditor.js"></script>
<script type="text/javascript">
       KE.show({
           id : 'content_1',

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
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">修改内容</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<form action="content.asp?cz=edit" method="post" name="jnr" id="jnr">
<table width="100%" border="0" cellspacing="0" class="table">
  <tr bgcolor="#F1F4F7">
    <td width="10%" height="40">分类：</td>
    <td>
      <select name="cid" id="cid" class="kuangy">
<%call addfl(qesy)%>
    </select></td>
  </tr>
  <tr>
    <td width="10%" height="40">标题：</td>
    <td>
      <input name="ntitle" type="text" class="kuang" id="ntitle" onFocus="this.className='kuang1'" onBlur="this.className='kuang'" value="<%=ns.rs("ntitle")%>" />
      <input name="posttime" type="hidden" id="posttime" value="<%=now()%>"  /><input name="newsid" type="hidden" id="newsid" value="<%=ns.rs("newsid")%>" /></td>
  </tr>
  <tr bgcolor="#F1F4F7">
    <td width="10%" height="40">关键字：</td>
    <td><input name="nkeyword" type="text" class="kuang" id="nkeyword" onFocus="this.className='kuang1'" onBlur="this.className='kuang'" value="<%=nkeyword_add(ns.rs("nkeyword"))%>" /></td>
  </tr>
  <tr>
    <td width="10%" height="40">简介：</td>
    <td>
      <input name="ninfo" type="text" class="kuang" id="ninfo" onFocus="this.className='kuang1'" onBlur="this.className='kuang'" value="<%=ns.rs("ninfo")%>" />
    </td>
  </tr>
  <tr bgcolor="#F1F4F7">
    <td width="10%" height="60">小图：</td>
    <td><input name="img" type="text" class="kuang" id="img" onFocus="this.className='kuang1'" onBlur="this.className='kuang'" value="<%=ns.rs("img")%>" /><br>
      <iframe frameborder="0" width="600" height="25" scrolling="no" src="upload.htm.asp" id="ff"></iframe></td>
  </tr>
  <tr>
    <td width="10%">内容：</td>
    <td><div id="cont">
      <textarea id="content_1" name="ncontent" style="width:650px;height:300px;visibility:hidden;"><%=ns.rs("ncontent")%></textarea>
      </div></td>
  </tr>
  <tr>
    <td colspan="2" align="center" height="40"><input type="submit" value=" 保 存 " style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name="" type="reset" style="border:1px #000000 solid;vertical-align:middle;height:25px" value=" 重 置 "></td>
    </tr>
</table>
</form>
</div>
<%elseif act="move" then
id=request.QueryString("id")
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">移动内容</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<form action="content.asp?cz=move" method="post" name="jnr" id="jnr">
<table width="100%" border="0" cellspacing="0" class="table">
  <tr bgcolor="#F1F4F7">
    <td width="10%" height="40">文章ID：</td>
    <td><input name="id" type="text" class="kuang" id="id" onFocus="this.className='kuang1'" onBlur="this.className='kuang'" value="<%=id%>" /></td>
  </tr>
  <tr>
    <td width="10%" height="40">目标分类：</td>
    <td><select name="cid" id="cid" class="kuangy">
      <%call addfl(qesy)%>
      </select></td>
  </tr>
 
  <tr>
    <td colspan="2" align="center" height="40"><input type="submit" value=" 提 交 " style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name="重置" type="reset" value=" 重 置 " style="border:1px #000000 solid;vertical-align:middle;height:25px"/></td>
    </tr>
</table>
</form>
</div>
<%elseif act="edit_all" then
id=request.QueryString("id")
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">设置内容</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<form action="content.asp?cz=edit_all" method="post" name="jnr" id="jnr">
<table width="100%" border="0" cellspacing="0" class="table">
  <tr bgcolor="#F1F4F7">
    <td width="10%" height="40">文章ID：</td>
    <td><input name="id" type="text" class="kuang" id="id" onFocus="this.className='kuang1'" onBlur="this.className='kuang'" value="<%=id%>" /></td>
  </tr>
  <tr>
    <td width="10%" height="40">是否推荐：</td>
    <td><select name="tuijian" id="tuijian" class="kuangy">
      <option value="0">否</option>
      <option value="1">是</option>
    </select></td>
  </tr>
 
  <tr>
    <td colspan="2" align="center" height="40"><input type="submit" value=" 提 交 " style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name="重置" type="reset" value=" 重 置 " style="border:1px #000000 solid;vertical-align:middle;height:25px"/></td>
    </tr>
</table>
</form>
</div>
<%
end if
%>
<br />
<br />
<table width="100%" border="0" cellpadding="0" cellspacing="0" align="center" class="table">
  <tr>
    <td align="center" height="40">Powered By <a href="http://www.q-cms.cn"><%=Version%></a>&nbsp;&nbsp;<%=UCase(q_Charset)%>&nbsp;&nbsp;<%=UCase(data_type)%> &copy; 2009 All Rights Reserved  <div class="no_view" style="display:none;"><script src="http://s87.cnzz.com/stat.php?id=1707573&web_id=1707573&show=pic" language="JavaScript" charset="gb2312"></script></div>
</td>
</table>
</body>
</html>
<%
cz=request.QueryString("cz")
if cz="add" then
set ns=new news
ns.cid=request.Form("cid")
ns.ntitle=request.Form("ntitle")
ns.nkeyword=keyword_add(request.Form("nkeyword"))
ns.ninfo=request.Form("ninfo")
ns.pinyin=pinyin(ns.ntitle)
ns.posttime=request.Form("posttime")
ns.img=request.Form("img")
ns.ncontent=request.Form("ncontent")
if ns.img="" then
ns.attpic=0
else
ns.attpic=1
end if
ns.insertnews()

set nx=new news
rs=nx.get_one_news()
nurl=cont_html_url(nx.rs("newsid"))
id=nx.rs("newsid")
call nx.update_url_news(nurl,id)
xs_err("<font color='red'>添加失败！</font>")
jump=zx_url("添加成功!","content.asp?active=add&cid="&ns.cid&"")

elseif cz="edit" then
newsid=request.Form("newsid")
set ns=new news
ns.cid=request.Form("cid")
ns.ntitle=request.Form("ntitle")
ns.nkeyword=keyword_edit(request.Form("nkeyword"))
ns.ninfo=request.Form("ninfo")
ns.posttime=request.Form("posttime")
ns.img=request.Form("img")
ns.ncontent=request.Form("ncontent")
if ns.img="" then
ns.attpic=0
else
ns.attpic=1
end if
ns.updatenews(newsid)

set nx=new news
rs=nx.getnewsinfo(newsid)
nurl=cont_html_url(nx.rs("newsid"))
id=nx.rs("newsid")
call nx.update_url_news(nurl,id)

xs_err("<font color='red'>修改新闻失败！</font>")
jump=zx_url("修改新闻成功!","content.asp?active=list&cid="&ns.cid&"")
elseif cz="del" then
newsid=request.QueryString("id")
cid=request.QueryString("cid")
set ns=new news
ns.deletenews(newsid)
xs_err("<font color='red'>删除新闻失败！</font>")
jump=zx_url("删除新闻成功!","content.asp?active=list&cid="&cid&"")
elseif cz="move" then
id=request.Form("id")
set ns=new news
ns.cid=request.Form("cid")
ns.news_move(id)
xs_err("<font color='red'>移动新闻失败！</font>")
jump=zx_url("移动新闻成功!","content.asp?active=list")
elseif cz="tuijian" then
id=request.QueryString("id")
set ns=new news
ns.updatecom(id)
xs_err("<font color='red'>推荐失败！</font>")
jump=zx_url("推荐成功!","content.asp?active=list")
elseif cz="ntuijian" then
id=request.QueryString("id")
set ns=new news
ns.updateqx(id)
xs_err("<font color='red'>取消失败！</font>")
jump=zx_url("取消成功!","content.asp?active=list")
elseif cz="edit_all" then
id=request.Form("id")
set ns=new news
ns.tuijian=request.Form("tuijian")
ns.news_edit_all(id)
xs_err("<font color='red'>设置失败！</font>")
jump=zx_url("设置成功!","content.asp?active=list")
elseif cz="replace" then
tixiang=request.Form("tixiang")
aa=request.Form("aa")
bb=request.Form("bb")
if tixiang="title" then
set ns=new news
rs=ns.news_search_q(aa)
do while not ns.rs.eof
cc=replace(ns.rs("ntitle"),aa,bb)
id=ns.rs("newsid")
call ns.update_title(cc,id)
ns.rs.movenext()
loop
xs_err("<font color='red'>替换失败！</font>")
jump=zx_url("替换成功,共替换了"&ns.rs.recordcount&"行!","content.asp?active=Replace")
elseif tixiang="content" then
set ns=new news
rs=ns.news_search_qc(aa)
do while not ns.rs.eof
cc=replace(ns.rs("ncontent"),aa,bb)
id=ns.rs("newsid")
call ns.update_content(cc,id)
ns.rs.movenext()
loop
xs_err("<font color='red'>替换失败！</font>")
jump=zx_url("替换成功,共替换了"&ns.rs.recordcount&"行!","content.asp?active=Replace")
end if
elseif cz="hirep" then
biao=request.Form("biao")
ziduan=request.Form("ziduan")
aa=request.Form("aa")
bb=request.Form("bb")
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql="select * from "&biao&" where "&ziduan&" like '%"&aa&"%'"
rs.open(sql)
do while not rs.eof
cc=replace(rs(""&ziduan&""),aa,bb)
strsql="update "&biao&" set "&ziduan&"='"&cc&"' where newsid="&rs.Fields(0)
conn.execute(strsql)
rs.movenext()
loop
xs_err("<font color='red'>替换失败！</font>")
jump=zx_url("替换成功,共替换了"&rs.recordcount&"行!","content.asp?active=Replace")
end if
%>
