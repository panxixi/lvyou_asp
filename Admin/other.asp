<!--#include file="../inc/utf.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/news.asp"-->
<!--#include file="../inc/sys.asp"-->
<!--#include file="../inc/category.asp"-->
<!--#include file="../inc/review.asp"-->
<!--#include file="../inc/do.asp"-->
<!--#include file="isadmin.asp"-->
<html>
<head>
<link rel=stylesheet href="styles/advanced/style.css" />
<meta http-equiv="Content-Type" content="text/html; charset=<%=q_Charset%>">
</head>
<body>
<%
op=request.QueryString("open")
if op="bq" then
act=request.QueryString("act")
if act="list" Then
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">标签管理 [ <a href="other.asp?open=bq&act=add">添加标签</a> ]</p></td>
  </tr>
</table>
<br />
<br />
<%
set cn=new config
rs=cn.biaoqian()
If cn.rs.eof Then
response.write("<table width='100%' border='0' cellspacing='0' class='table'><tr bgcolor='#F1F4F7'><td height='40'>暂时无任何标签!</td></tr></table>")
else
    cn.rs.PageSize = 20
	Page = CLng(Request("Page"))
    If Page < 1 Or page="" Then Page = 1
    If Page > cn.rs.PageCount Then Page = cn.rs.PageCount
    cn.rs.AbsolutePage = Page
%>
<script language="JavaScript">
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
  strurl = "other.asp?cz=bq&done=del&id=" + strid;
  if(!s) {
    alert("请选择要删除的标签!");
    return false;
  }	
  if (confirm("你确定要删除这些标签吗？")) {
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
  <tr bgcolor="#DAE2E8">
    <td width="10%">编号</td>
    <td width="50%">名称</td>
    <td width="35%">操作</td>
    <td width="5%" align="left"><input name="" type="checkbox" value=""></td>
  </tr>
  <%
  for i=1 to cn.rs.PageSize
  if cn.rs.EOF then Exit For
  %>
  <%
  if i mod 2=0 then
  %>
   <tr bgcolor="#F1F4F7">
   <%else%>
  <tr>
  <%end if%>
    <td><%=cn.rs("id")%></td>
    <td><%=cn.rs("names")%></td>
    <td><a href="other.asp?open=bq&act=edit&id=<%=cn.rs("id")%>">修改</a> <a href="other.asp?cz=bq&done=del&id=<%=cn.rs("id")%>">删除</a></td>
    <td><input name="cate" type="checkbox" id="<%=cn.rs("id")%>" /></td>
  </tr>
  <%
  cn.rs.movenext()
  next
  %>
  <tr><td colspan="4" align="center"><%
			if page=1 then
response.Write "首页&nbsp;上一页&nbsp;第"&page&"页&nbsp;<a href='other.asp?open=bq&act=list&page="&page+1&"'>下一页</a>&nbsp;<a href='other.asp?open=bq&act=list&page="&cn.rs.pagecount&"'>尾页</a>"
elseif page=cn.rs.pagecount then
response.Write "<a href='other.asp?open=bq&act=list&page=1'>首页</a>&nbsp;<a href='other.asp?open=bq&act=list&page="&page-1&"'>上一页</a>&nbsp;第"&page&"页&nbsp;下一页&nbsp;尾页"
else
response.Write "<a href='other.asp?open=bq&act=list&page=1'>首页</a>&nbsp;<a href=other.asp?open=bq&act=list&page="&page-1&"'>上一页</a>&nbsp;第"&page&"页&nbsp;<a href='other.asp?open=bq&act=list&page="&page+1&"'>下一页</a>&nbsp;<a href=other.asp?open=bq&act=list&page="&cn.rs.pagecount&"'>尾页</a>"
end if
			%></td></tr>
            <tr><td align="center" colspan="4" height="40"><input type="button" value=" 全 选 " onClick="sltall()" style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input type="button" value=" 清 空 " onClick="sltnull()" style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input name="tijiao" type="submit" value=" 删 除 " onClick="SelectNews()" style="border:1px #000000 solid;vertical-align:middle;height:25px"/></td></tr>
</table>
</form>
</div>
<%end if%>
<%
elseif act="add" then
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">添加标签</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<script type="text/javascript" charset="utf-8" src="kindeditor/kindeditor.js"></script>
<script type="text/javascript">
       KE.init({
           id : 'bq',

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
<form action="other.asp?cz=bq&done=add" method="post">
<table width="100%" border="0" cellspacing="0" class="table">
  <tr>
    <td width="10%">名称：</td>
    <td>
      <input name="names" type="text" id="names" class="kuang" />
      
    </td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF">内容：</td>
    <td bgcolor="#FFFFFF"><div id="cont">
     <textarea id="bq" name="bq" style="width:700px;height:300px;"></textarea> 
     </div>    
    </td>
  </tr>
  <tr>
    <td colspan="2" align="center" height="40"><input type="submit" value="保存设置" style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" name="button" value="加载编辑器" onClick="javascript:KE.create('bq');" style="border:1px #000000 solid;vertical-align:middle;height:25px"/></td>
    </tr>
</table>
</form>
</div>
<%
elseif act="edit" then
id=request.QueryString("id")
set cn=new config
rs=cn.getbiaoqian(id)
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">修改标签</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<script type="text/javascript" charset="utf-8" src="kindeditor/kindeditor.js"></script>
<script type="text/javascript">
       KE.init({
           id : 'bq',

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
<form action="other.asp?cz=bq&done=edit" method="post">
<table width="100%" border="0" cellspacing="0" class="table">
  <tr>
    <td width="10%">名称：</td>
    <td>
      <input name="names" type="text" class="kuang" id="names" value="<%=cn.rs("names")%>" /><input name="id" type="hidden" id="id" value="<%=cn.rs("id")%>">
      
    </td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF">内容：</td>
    <td bgcolor="#FFFFFF"><div id="cont">
     <textarea id="bq" name="bq" style="width:700px;height:300px;"><%=cn.rs("bq")%></textarea> 
     </div>    
    </td>
  </tr>
  <tr>
    <td colspan="2" align="center" height="40"><input type="submit" value="保存设置" style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" name="button" value="加载编辑器" onClick="javascript:KE.create('bq');" style="border:1px #000000 solid;vertical-align:middle;height:25px"/></td>
    </tr>
</table>
</form>
</div>
<%end if
%>
<%
elseif op="diy" then
act=request.QueryString("act")
if act="list" then
set cn=new config
rs=cn.diylist()
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">自定义页面管理 [ <a href="other.asp?open=diy&act=add">添加自定义页面</a> ]</p></td>
  </tr>
</table>
<br />
<br />
<%
If cn.rs.eof Then
response.write("<table width='100%' border='0' cellspacing='0' class='table'><tr bgcolor='#F1F4F7'><td height='40'>暂时无任何自定义页面!</td></tr></table>")
else
    cn.rs.PageSize = 20
	Page = CLng(Request("Page"))
    If Page < 1 Then Page = 1
    If Page > cn.rs.PageCount Then Page = cn.rs.PageCount
    cn.rs.AbsolutePage = Page
%>
<script language="JavaScript">
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
  strurl = "other.asp?cz=bq&done=del&id=" + strid;
  if(!s) {
    alert("请选择要删除的自定义页面!");
    return false;
  }	
  if (confirm("你确定要删除这些自定义页面吗？")) {
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
  <tr bgcolor="#DAE2E8">
    <td width="10%">编号</td>
    <td width="50%">名称</td>
    <td width="35%">操作</td>
    <td width="5%" align="left"><input name="" type="checkbox" value=""></td>
  </tr>
  <%
  for i=1 to cn.rs.PageSize
  if cn.rs.EOF then Exit For
  %>
  <%
  if i mod 2=0 then
  %>
   <tr bgcolor="#F1F4F7">
   <%else%>
  <tr>
  <%end if%>
    <td><%=cn.rs("id")%></td>
    <td><%=cn.rs("names")%></td>
    <td><a href="other.asp?open=diy&act=edit&id=<%=cn.rs("id")%>">修改</a> <a href="other.asp?cz=diy&done=del&id=<%=cn.rs("id")%>">删除</a></td>
    <td><input name="cate" type="checkbox" id="<%=cn.rs("id")%>" /></td>
  </tr>
  <%
  cn.rs.movenext()
  next
  %>
  <tr><td colspan="4" align="center"><%
			if page=1 then
response.Write "首页&nbsp;上一页&nbsp;第"&page&"页&nbsp;<a href='other.asp?open=diy&act=list&page="&page+1&"'>下一页</a>&nbsp;<a href='other.asp?open=diy&act=list&page="&cn.rs.pagecount&"'>尾页</a>"
elseif page=cn.rs.pagecount then
response.Write "<a href='other.asp?open=diy&act=list&page=1'>首页</a>&nbsp;<a href='other.asp?open=diy&act=list&page="&page-1&"'>上一页</a>&nbsp;第"&page&"页&nbsp;下一页&nbsp;尾页"
else
response.Write "<a href='other.asp?open=diy&act=list&page=1'>首页</a>&nbsp;<a href=other.asp?open=diy&act=list&page="&page-1&"'>上一页</a>&nbsp;第"&page&"页&nbsp;<a href='other.asp?open=diy&act=list&page="&page+1&"'>下一页</a>&nbsp;<a href=other.asp?open=diy&act=list&page="&cn.rs.pagecount&"'>尾页</a>"
end if
			%></td></tr>
            <tr><td align="center" colspan="4" height="40"><input type="button" value=" 全 选 " onClick="sltall()" style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input type="button" value=" 清 空 " onClick="sltnull()" style="border:1px #000000 solid;vertical-align:middle;height:25px"/>
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input name="tijiao2" type="submit" value=" 删 除 " onClick="SelectNews()" style="border:1px #000000 solid;vertical-align:middle;height:25px"/></td></tr>
</table>
</form>
</div>
<%End if%>
<%
elseif act="add" then
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">添加自定义页面</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<script type="text/javascript" charset="utf-8" src="kindeditor/kindeditor.js"></script>
<script type="text/javascript">
       KE.init({
           id : 'diy',

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
<form action="other.asp?cz=diy&done=add" method="post">
<table width="100%" border="0" cellspacing="0" class="table">
  <tr>
    <td width="10%">名称：</td>
    <td>
      <input name="names" type="text" id="names" class="kuang" />
      
    </td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF">内容：</td>
    <td bgcolor="#FFFFFF"><div id="cont">
     <textarea id="diy" name="diy" style="width:700px;height:300px;"></textarea> 
     </div>    
    </td>
  </tr>
  <tr>
    <td colspan="2" align="center" height="40"><input type="submit" value="保存设置" style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" name="button" value="加载编辑器" onClick="javascript:KE.create('diy');" style="border:1px #000000 solid;vertical-align:middle;height:25px"/></td>
    </tr>
</table>
</form>
</div>
<%
elseif act="edit" then
id=request.QueryString("id")
set cn=new config
rs=cn.getdiy(id)
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">修改自定义页面</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<script type="text/javascript" charset="utf-8" src="kindeditor/kindeditor.js"></script>
<script type="text/javascript">
       KE.init({
           id : 'diy',

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
<form action="other.asp?cz=diy&done=edit" method="post">
<table width="100%" border="0" cellspacing="0" class="table">
  <tr>
    <td width="10%">名称：</td>
    <td>
      <input name="names" type="text" class="kuang" id="names" value="<%=cn.rs("names")%>" /><input name="id" type="hidden" id="id" value="<%=cn.rs("id")%>">
      
    </td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF">内容：</td>
    <td bgcolor="#FFFFFF"><div id="cont">
     <textarea id="diy" name="diy" style="width:700px;height:300px;"><%=cn.rs("diy")%></textarea> 
     </div>    
    </td>
  </tr>
  <tr>
    <td colspan="2" align="center" height="40"><input type="submit" value="保存设置" style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" name="button" value="加载编辑器" onClick="javascript:KE.create('diy');" style="border:1px #000000 solid;vertical-align:middle;height:25px"/></td>
    </tr>
</table>
</form>
</div>
<%end if%>
<%
elseif op="js" then
act=request.QueryString("act")
if act="list" then
set cn=new config
rs=cn.jss_list()
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">JS管理 [ <a href="other.asp?open=js&act=add">添加JS</a> ]</p></td>
  </tr>
</table>
<br />
<br />
<%
If cn.rs.eof Then
response.write("<table width='100%' border='0' cellspacing='0' class='table'><tr bgcolor='#F1F4F7'><td height='40'>暂时无任何外部调用!</td></tr></table>")
else
    cn.rs.PageSize = 20
	Page = CLng(Request("Page"))
    If Page < 1 Then Page = 1
    If Page > cn.rs.PageCount Then Page = cn.rs.PageCount
    cn.rs.AbsolutePage = Page
%>
<script language="JavaScript">
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
  strurl = "other.asp?cz=js&done=del&id=" + strid;
  if(!s) {
    alert("请选择要删除的JS!");
    return false;
  }	
  if (confirm("你确定要删除这些JS吗？")) {
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
  <tr bgcolor="#DAE2E8">
    <td width="10%">编号</td>
    <td width="20%">名称</td>
    <td width="50%">调用代码</td>
    <td width="15%">操作</td>
    <td width="5%" align="left"><input name="" type="checkbox" value=""></td>
  </tr>
  <%
  for i=1 to cn.rs.PageSize
  if cn.rs.EOF then Exit For
  %>
  <%
  if i mod 2=0 then
  %>
   <tr bgcolor="#F1F4F7">
   <%else%>
  <tr>
  <%end if%>
    <td><%=cn.rs("id")%></td>
    <td><%=cn.rs("js_name")%></td>
    <td><input name="input" type="text" value="<script language=JavaScript src='http://<%=request.ServerVariables("http_HOST")%><%=install%>inc/js.asp?id=<%=cn.rs("id")%>'></script>" size="70" /></td>
    <td><a href="other.asp?open=js&act=edit&id=<%=cn.rs("id")%>">修改</a> <a href="other.asp?cz=js&done=del&id=<%=cn.rs("id")%>">删除</a></td>
    <td><input name="cate" type="checkbox" id="<%=cn.rs("id")%>" /></td>
  </tr>
  <%
  cn.rs.movenext()
  next
  %>
  <tr><td colspan="5" align="center"><%
			if page=1 then
response.Write "首页&nbsp;上一页&nbsp;第"&page&"页&nbsp;<a href='other.asp?open=js&act=list&page="&page+1&"'>下一页</a>&nbsp;<a href='other.asp?open=js&act=list&page="&cn.rs.pagecount&"'>尾页</a>"
elseif page=cn.rs.pagecount then
response.Write "<a href='other.asp?open=js&act=list&page=1'>首页</a>&nbsp;<a href='other.asp?open=js&act=list&page="&page-1&"'>上一页</a>&nbsp;第"&page&"页&nbsp;下一页&nbsp;尾页"
else
response.Write "<a href='other.asp?open=js&act=list&page=1'>首页</a>&nbsp;<a href=other.asp?open=js&act=list&page="&page-1&"'>上一页</a>&nbsp;第"&page&"页&nbsp;<a href='other.asp?open=js&act=list&page="&page+1&"'>下一页</a>&nbsp;<a href=other.asp?open=js&act=list&page="&cn.rs.pagecount&"'>尾页</a>"
end if
			%></td></tr>
            <tr><td align="center" colspan="5" height="40"><input type="button" value=" 全 选 " onClick="sltall()" style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input type="button" value=" 清 空 " onClick="sltnull()" style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input name="tijiao3" type="submit" value=" 删 除 " onClick="SelectNews()" style="border:1px #000000 solid;vertical-align:middle;height:25px"/></td></tr>
</table>
</form>
</div>
<%
End if
elseif act="add" then
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">添加JS</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<script type="text/javascript" charset="utf-8" src="kindeditor/kindeditor.js"></script>
<script type="text/javascript">
       KE.init({
           id : 'js_code',

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
<form action="other.asp?cz=js&done=add" method="post">
<table width="100%" border="0" cellspacing="0" class="table">
  <tr>
    <td width="10%">名称：</td>
    <td>
      <input name="js_name" type="text" id="js_name" class="kuang" />
      
    </td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF">内容：</td>
    <td bgcolor="#FFFFFF"><div id="cont">
     <textarea id="js_code" name="js_code" style="width:700px;height:300px;"></textarea> 
     </div>    
    </td>
  </tr>
  <tr>
    <td colspan="2" align="center" height="40"><input type="submit" value="保存设置"  style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" name="button" value="加载编辑器" onClick="javascript:KE.create('js_code');" style="border:1px #000000 solid;vertical-align:middle;height:25px"/></td>
    </tr>
</table>
</form>
</div>
<%
elseif act="edit" then
id=request.QueryString("id")
set cn=new config
rs=cn.getjss(id)
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">修改JS</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<script type="text/javascript" charset="utf-8" src="kindeditor/kindeditor.js"></script>
<script type="text/javascript">
       KE.init({
           id : 'js_code',

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
<form action="other.asp?cz=js&done=edit" method="post">
<table width="100%" border="0" cellspacing="0" class="table">
  <tr>
    <td width="10%">名称：</td>
    <td>
      <input name="js_name" type="text" class="kuang" id="js_name" value="<%=cn.rs("js_name")%>" /><input name="id" type="hidden" id="id" value="<%=cn.rs("id")%>">
      
    </td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF">内容：</td>
    <td bgcolor="#FFFFFF"><div id="cont">
     <textarea id="js_code" name="js_code" style="width:700px;height:300px;"><%=js2html(cn.rs("js_code"))%></textarea> 
     </div>    
    </td>
  </tr>
  <tr>
    <td colspan="2" align="center" height="40"><input type="submit" value="保存设置" style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" name="button" value="加载编辑器" onClick="javascript:KE.create('js_code');" style="border:1px #000000 solid;vertical-align:middle;height:25px"/></td>
    </tr>
</table>
</form>
</div>
<%end if%>
<%
elseif op="ilink" then
act=request.QueryString("act")
if act="list" then
set cn=new config
rs=cn.olink_list()
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">站内连接管理 [ <a href="other.asp?open=ilink&act=add">添加站内连接</a> ]</p></td>
  </tr>
</table>
<br />
<br />
<%
If cn.rs.eof Then
response.write("<table width='100%' border='0' cellspacing='0' class='table'><tr bgcolor='#F1F4F7'><td height='40'>暂时无任何站内连接!</td></tr></table>")
else
    cn.rs.PageSize = 20
	Page = CLng(Request("Page"))
    If Page < 1 Then Page = 1
    If Page > cn.rs.PageCount Then Page = cn.rs.PageCount
    cn.rs.AbsolutePage = Page
%>
<script language="JavaScript">
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
  strurl = "other.asp?cz=ilink&done=del&id=" + strid;
  if(!s) {
    alert("请选择要删除的JS!");
    return false;
  }	
  if (confirm("你确定要删除这些JS吗？")) {
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
  <tr bgcolor="#DAE2E8">
    <td width="10%">编号</td>
    <td width="20%">名称</td>
    <td width="50%">路径</td>
    <td width="15%">操作</td>
    <td width="5%" align="left"><input name="" type="checkbox" value=""></td>
  </tr>
  <%
  for i=1 to cn.rs.PageSize
  if cn.rs.EOF then Exit For
  %>
  <%
  if i mod 2=0 then
  %>
   <tr bgcolor="#F1F4F7">
   <%else%>
  <tr>
  <%end if%>
    <td><%=cn.rs("id")%></td>
    <td><%=cn.rs("link_name")%></td>
    <td><%=cn.rs("link_url")%></td>
    <td><a href="other.asp?open=ilink&act=edit&id=<%=cn.rs("id")%>">修改</a> <a href="other.asp?cz=ilink&done=del&id=<%=cn.rs("id")%>">删除</a></td>
    <td><input name="cate" type="checkbox" id="<%=cn.rs("id")%>" /></td>
  </tr>
  <%
  cn.rs.movenext()
  next
  %>
  <tr><td colspan="5" align="center"><%
			if page=1 then
response.Write "首页&nbsp;上一页&nbsp;第"&page&"页&nbsp;<a href='other.asp?open=ilink&act=list&page="&page+1&"'>下一页</a>&nbsp;<a href='other.asp?open=ilink&act=list&page="&cn.rs.pagecount&"'>尾页</a>"
elseif page=cn.rs.pagecount then
response.Write "<a href='other.asp?open=ilink&act=list&page=1'>首页</a>&nbsp;<a href='other.asp?open=ilink&act=list&page="&page-1&"'>上一页</a>&nbsp;第"&page&"页&nbsp;下一页&nbsp;尾页"
else
response.Write "<a href='other.asp?open=ilink&act=list&page=1'>首页</a>&nbsp;<a href=other.asp?open=ilink&act=list&page="&page-1&"'>上一页</a>&nbsp;第"&page&"页&nbsp;<a href='other.asp?open=ilink&act=list&page="&page+1&"'>下一页</a>&nbsp;<a href=other.asp?open=ilink&act=list&page="&cn.rs.pagecount&"'>尾页</a>"
end if
			%></td></tr>
            <tr><td align="center" colspan="5" height="40"><input type="button" value=" 全 选 " onClick="sltall()" style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input type="button" value=" 清 空 " onClick="sltnull()" style="border:1px #000000 solid;vertical-align:middle;height:25px"/>
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name="tijiao" type="submit" value=" 删 除 " onClick="SelectNews()" style="border:1px #000000 solid;vertical-align:middle;height:25px"/></td></tr>
</table>
</form>
</div>
<%End if%>
<%
elseif act="add" then
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">添加站内连接</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<form action="other.asp?cz=ilink&done=add" method="post">
<table width="100%" border="0" cellspacing="0" class="table">
  <tr>
    <td width="10%">名称：</td>
    <td>
      <input name="link_name" type="text" id="link_name" class="kuang" />
      
    </td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF">地址：</td>
    <td bgcolor="#FFFFFF"><input name="link_url" type="text" id="link_url" class="kuang" />  
    </td>
  </tr>
  <tr>
    <td colspan="2" align="center" height="40"><input type="submit" value=" 保 存 " style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name="" type="reset" value=" 重 置 " style="border:1px #000000 solid;vertical-align:middle;height:25px"></td>
    </tr>
</table>
</form>
</div>
<%
elseif act="edit" then
id=request.QueryString("id")
set cn=new config
rs=cn.getolink(id)
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">修改站内连接</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<form action="other.asp?cz=ilink&done=edit" method="post">
<table width="100%" border="0" cellspacing="0" class="table">
  <tr>
    <td width="10%">名称：</td>
    <td>
      <input name="link_name" type="text" class="kuang" id="link_name" value="<%=cn.rs("link_name")%>" /><input name="id" type="hidden" id="id" value="<%=cn.rs("id")%>">
      
    </td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF">内容：</td>
    <td bgcolor="#FFFFFF"><input name="link_url" type="text" class="kuang" id="link_url" value="<%=cn.rs("link_url")%>" />
    </td>
  </tr>
  <tr>
    <td colspan="2" align="center" height="40"><input type="submit" value=" 保 存 " style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name="" type="reset" value=" 重 置 " style="border:1px #000000 solid;vertical-align:middle;height:25px"></td>
    </tr>
</table>
</form>
</div>
<%end if%>
<%
elseif op="temp" then
act=request.QueryString("act")
if act="list" then
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">模版管理 [ <a href="other.asp?open=temp&act=add">添加模版</a> ]</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">

  <table width="100%" border="0" cellspacing="0" class="table">
  <tr bgcolor="#DAE2E8">
    <td width="5%">编号</td>
    <td width="70%">名称</td>
    <td width="25%">操作</td>
  </tr>
  <%
kk=sort(split(mblist(),","))
for j=1 to ubound(kk)
  %>
  <%
  if j mod 2=0 then
  %>
   <tr bgcolor="#F1F4F7">
   <%else%>
  <tr>
  <%end if%>
    <td><%=j%></td>
    <td><%=kk(j)%></td>
    <td><a href="other.asp?open=temp&act=edit&t_name=<%=kk(j)%>">修改</a> <a href="other.asp?open=temp&act=copy&t_name=<%=kk(j)%>">复制</a> <a href="other.asp?cz=temp&done=del&t_name=<%=kk(j)%>">删除</a></td>
  </tr>
  <%
  next
  %>
</table>
</div>
<%
elseif act="add" then
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">添加模版</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<script type="text/javascript" charset="utf-8" src="kindeditor/kindeditor.js"></script>
<script type="text/javascript">
       KE.init({
           id : 't_code',

		   cssPath : 'kindeditor/index.css',skinType: 'tinymce',
        items : [
            'source', 'preview',  'fullscreen', 'undo', 'redo', 'cut', 'copy', 'paste',
            'plainpaste', 'wordpaste', 'justifyleft', 'justifycenter', 'justifyright',
            'justifyfull', 'insertorderedlist', 'insertunorderedlist', 'indent', 'outdent', 'subscript',
            'superscript', 'date', 'time', 'specialchar', 'emoticons', 'link', 'unlink', '-',
            'title', 'fontname', 'fontsize', 'textcolor', 'bgcolor', 'bold',
            'italic', 'underline', 'strikethrough', 'removeformat', 'selectall', 'image',
            'flash', 'media', 'layer', 'table', 'hr', 'about'
        ]		   
       });
   </script>
<form action="other.asp?cz=temp&done=add" method="post">
<table width="100%" border="0" cellspacing="0" class="table">
  <tr>
    <td width="10%">名称：</td>
    <td>
      <input name="t_name" type="text" id="t_name" class="kuang" />
      
    </td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF">内容：</td>
    <td bgcolor="#FFFFFF"><div id="cont">
     <textarea id="t_code" name="t_code" style="width:700px;height:300px;"></textarea> 
     </div>    
    </td>
  </tr>
  <tr>
    <td colspan="2" align="center" height="40"><input type="submit" value="保存设置" style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" name="button" value="加载编辑器" onClick="javascript:KE.create('t_code');" style="border:1px #000000 solid;vertical-align:middle;height:25px"/></td>
    </tr>
</table>
</form>
</div>
<%
elseif act="edit" then
t_name=request.QueryString("t_name")
t_code=read_html(server.mappath(install&"templist/"&temp_url&"/"&t_name),q_Charset)
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">修改模版</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<script type="text/javascript" charset="utf-8" src="kindeditor/kindeditor.js"></script>
<script type="text/javascript">
       KE.init({
           id : 't_code',

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
<form action="other.asp?cz=temp&done=edit" method="post">
<table width="100%" border="0" cellspacing="0" class="table">
  <tr>
    <td width="10%">名称：</td>
    <td>
      <input name="t_name" type="text" class="kuang" id="t_name" value="<%=t_name%>" /></td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF">内容：</td>
    <td bgcolor="#FFFFFF"><div id="cont">
     <textarea id="t_code" name="t_code" style="width:700px;height:300px;"><%=server.HTMLEncode(t_code)%></textarea> 
     </div>    
    </td>
  </tr>
  <tr>
    <td colspan="2" align="center" height="40"><input type="submit" value="保存设置" style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" name="button" value="加载编辑器" onClick="javascript:KE.create('t_code');" style="border:1px #000000 solid;vertical-align:middle;height:25px"/></td>
    </tr>
</table>
</form>
</div>
<%
elseif act="copy" then
t_name=request.QueryString("t_name")
t_code=read_html(server.mappath(install&"templist/"&temp_url&"/"&t_name),q_Charset)
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">复制模版</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<script type="text/javascript" charset="utf-8" src="kindeditor/kindeditor.js"></script>
<script type="text/javascript">
       KE.init({
           id : 't_code',

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
<form action="other.asp?cz=temp&done=copy" method="post">
<table width="100%" border="0" cellspacing="0" class="table">
  <tr>
    <td width="10%">名称：</td>
    <td>
      <input name="t_name" type="text" class="kuang" id="t_name" value="<%=t_name%>" /></td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF">内容：</td>
    <td bgcolor="#FFFFFF"><div id="cont">
     <textarea id="t_code" name="t_code" style="width:700px;height:300px;"><%=server.HTMLEncode(t_code)%></textarea> 
     </div>    
    </td>
  </tr>
  <tr>
    <td colspan="2" align="center" height="40"><input type="submit" value="保存设置" style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" name="button" value="加载编辑器" onClick="javascript:KE.create('t_code');" style="border:1px #000000 solid;vertical-align:middle;height:25px"/></td>
    </tr>
</table>
</form>
</div>
<%end if%>
<%
elseif op="guest" then
act=request.QueryString("act")
if act="list" then
%>
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
  strurl = "other.asp?cz=guest&done=del&id=" + strid;
  if(!s) {
    alert("请选择要删除的留言!");
    return false;
  }	
  if (confirm("确定要删除这些留言吗？")) {
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
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">留言列表</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<form method="post" name="form1" id="form1">
<table width="100%" border="0" cellspacing="0" class="table">
<%
set cn=new config
rs=cn.guest()
if cn.rs.eof then
%>
<tr bgcolor="#F1F4F7">
    <td width="10%"  height="40" colspan="3">暂无任何留言！</td></tr>
    <%
else
    cn.rs.PageSize = 20
	Page = CLng(Request("Page"))
    If Page < 1 or page="" Then Page = 1
    If Page > cn.rs.PageCount Then Page = cn.rs.PageCount
    cn.rs.AbsolutePage = Page
    for i=1 to cn.rs.pagesize
    if cn.rs.eof then exit for
%>
<tr bgcolor="#DAE2E8">
  <td width="8%"><input name="cate" type="checkbox" id="<%=cn.rs("id")%>" /></td>
  <td width="70%"><font color="#666666">标题：</font><%=cn.rs("title")%>&nbsp;&nbsp;</td>
  <td>&nbsp;&nbsp;<a href="other.asp?open=guest&act=edit&id=<%=cn.rs("id")%>">修改</a>&nbsp;&nbsp;<a href="other.asp?open=guest&act=hf&id=<%=cn.rs("id")%>">回复</a>&nbsp;&nbsp;<a href="other.asp?cz=guest&done=del&id=<%=cn.rs("id")%>">删除</a></td></tr>
   <tr>
     <td>留言：</td>
     <td colspan="2"><%=cn.rs("content")%></td>
   </tr>
   <tr>
     <td >回复：</td>
     <td colspan="2"><%
	 
	 if cn.rs("hf")<>"" then
	 response.Write (cn.rs("hf"))
	 else
	 response.Write("<font color='#ff0000'>暂无回复！</font>")
	 end if
	 %></td>
   </tr>

  <tr>
      <td colspan="3"><font color="#666666">帐号：</font><%=cn.rs("name")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#666666">电话：</font><%=cn.rs("tel")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#666666">邮箱：</font><%=cn.rs("email")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#666666">QQ：</font><%=cn.rs("qq")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#666666">IP：</font><%=cn.rs("ip")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#666666">时间：</font><%=cn.rs("times")%></td>
      </tr>
      <%
  cn.rs.movenext()
  next
  
  %>
  <tr>
    <td colspan="3" align="center">
      <%
			if page=1 then
response.Write "首页&nbsp;上一页&nbsp;第"&page&"页&nbsp;<a href='other.asp?open=guest&act=list&page="&page+1&"'>下一页</a>&nbsp;<a href='other.asp?open=guest&act=list&page="&cn.rs.pagecount&"'>尾页</a>"
elseif page=cn.rs.pagecount then
response.Write "<a href='other.asp?open=guest&act=list&page=1'>首页</a>&nbsp;<a href='other.asp?open=guest&act=list&page="&page-1&"'>上一页</a>&nbsp;第"&page&"页&nbsp;下一页&nbsp;尾页"
else
response.Write "<a href='other.asp?open=guest&act=list&page=1'>首页</a>&nbsp;<a href='other.asp?open=guest&act=list&page="&page-1&"'>上一页</a>&nbsp;第"&page&"页&nbsp;<a href='other.asp?open=guest&act=list&page="&page+1&"'>下一页</a>&nbsp;<a href='other.asp?open=guest&act=list&page="&cn.rs.pagecount&"'>尾页</a>"

			%></td>
  </tr>
<tr><td align="center" colspan="3" height="40"><input type="button" value=" 全 选 " onClick="sltall()" style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input type="button" value=" 清 空 " onClick="sltnull()" style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input name="tijiao" type="submit" value=" 删 除 " onClick="SelectChk()" style="border:1px #000000 solid;vertical-align:middle;height:25px"/></td></tr>
	    <%
	    end if
		end if
	    %>
</table>
</form>
</div>
<%
elseif act="hf" then
id=request.QueryString("id")
set cn=new config
cn.getguest(id)
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">回复留言</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<form action="other.asp?cz=guest&done=hf" method="post">
<table width="100%" border="0" cellspacing="0" class="table">
<tr bgcolor="#F1F4F7">
  <td height="140" width="10%">内容：</td>
  <td>
    <textarea name="hf" cols="30" rows="6" id="hf" class="kuangx" onBlur="this.className='kuangx'" onFocus="this.className='kuangx1'"><%=cn.rs("hf")%></textarea><input name="id" type="hidden" id="id" value="<%=id%>"></td>
</tr>
<tr bgcolor="#F1F4F7">
  <td height="40" colspan="2" align="center"><input name="bot" type="submit" id="bot" value=" 保 存 "  style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
       
        <input type="reset" name="but" id="but" value=" 重 置 " style="border:1px #000000 solid;vertical-align:middle;height:25px">
        </td>
  </tr>
</table>
</form></div>
<%
elseif act="edit" then
id=request.QueryString("id")
set cn=new config
cn.getguest(id)
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">修改留言</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<form action="other.asp?cz=guest&done=edit" method="post">
<table width="100%" border="0" cellspacing="0" class="table">
<tr bgcolor="#F1F4F7"><td height="40">标题：</td>
  <td colspan="3"><input name="title" type="text" class="kuang" id="title" onFocus="this.className='kuang'" onBlur="this.className='kuang'" value="<%=cn.rs("title")%>"></td>
</tr>
<tr><td width="10%" height="40">名称：</td>
  <td width="40%"><input name="name" type="text" class="kuangy" id="name" onFocus="this.className='kuangy1'" onBlur="this.className='kuangy'" value="<%=cn.rs("name")%>"></td>
  <td width="10%">QQ：</td>
  <td width="40%"><input name="qq" type="text" class="kuangy" id="qq" onFocus="this.className='kuangy1'" onBlur="this.className='kuangy'" value="<%=cn.rs("qq")%>"></td></tr>
<tr bgcolor="#F1F4F7"><td height="40">电话：</td>
  <td><input name="tel" type="text" class="kuangy" id="tel" onFocus="this.className='kuangy1'" onBlur="this.className='kuangy'" value="<%=cn.rs("tel")%>"></td>
  <td>时间：</td>
  <td><input name="times" type="text" class="kuangy" id="times" onFocus="this.className='kuangy1'" onBlur="this.className='kuangy'" value="<%=cn.rs("times")%>"></td></tr>
<tr><td height="40">EMAIL：</td>
  <td><input name="email" type="text" class="kuangy" id="email" onFocus="this.className='kuangy1'" onBlur="this.className='kuangy'" value="<%=cn.rs("email")%>"></td>
  <td>IP：</td>
  <td><input name="ip" type="text" class="kuangy" id="ip" onFocus="this.className='kuangy1'" onBlur="this.className='kuangy'" value="<%=cn.rs("ip")%>"></td></tr>
<tr bgcolor="#F1F4F7"><td >留言：</td>
  <td colspan="3" height="140"><textarea name="content" cols="30" rows="6" id="content" class="kuangx" onBlur="this.className='kuangx'" onFocus="this.className='kuangx1'"><%=cn.rs("content")%></textarea>
    <input name="id" type="hidden" id="id" value="<%=id%>"></td>
</tr><tr><td colspan="4" height="40" align="center"><input name="bot1" type="submit" id="bot1" value=" 保 存 "  style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
       
          <input type="reset" name="but2" id="but2" value=" 重 置 " style="border:1px #000000 solid;vertical-align:middle;height:25px">
        </td></tr>
</table>
</form>
</div>
<%
end if
%>
<%
elseif op="review" then
%>
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
  strurl = "other.asp?cz=review&done=del&id=" + strid;
  if(!s) {
    alert("请选择要删除的回复!");
    return false;
  }	
  if (confirm("确定要删除这些回复吗？")) {
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
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">回复管理</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<form method="post" name="form1" id="form1">
<table width="100%" border="0" cellspacing="0" class="table">
<%
set rv=new review
rs=rv.did_all()
if rv.rs.eof then
%>
<tr bgcolor="#F1F4F7"><td width="10%"  height="40" colspan="2">暂无任何回复！</td></tr>
<%
else
    rv.rs.PageSize = 20
	Page = CLng(Request("Page"))
    If Page < 1 or page="" Then Page = 1
    If Page > rv.rs.PageCount Then Page = rv.rs.PageCount
    rv.rs.AbsolutePage = Page
    for i=1 to rv.rs.pagesize
    if rv.rs.eof then exit for
%>
<tr bgcolor="#DAE2E8">
  <td width="8%"><input name="cate" type="checkbox" id="<%=rv.rs("did")%>" /></td>
  <td><font color="#666666">标题：</font><%=rv.rs("ntitle")%>&nbsp;&nbsp;&nbsp;</td>
  </tr>
   <tr>
     <td>留言：</td>
     <td><%=rv.rs("dcontent")%></td>
   </tr>
   

  <tr>
      <td colspan="2"><font color="#666666">帐号：</font><%=rv.rs("poster")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#666666">时间：</font><%=rv.rs("discuss.posttime")%></td>
      </tr>
      <%
  rv.rs.movenext()
  next
  
  %>
  <tr>
    <td colspan="2" align="center">
      </td>
  </tr>
<tr><td align="center" colspan="2" height="40"><input type="button" value=" 全 选 " onClick="sltall()" style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input type="button" value=" 清 空 " onClick="sltnull()" style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input name="tijiao" type="submit" value=" 删 除 " onClick="SelectChk()" style="border:1px #000000 solid;vertical-align:middle;height:25px"/></td></tr>
	    <%
	    end if
	    %>
</table>
</form>
</div>
<%
elseif op="plus" then
act=request.QueryString("act")
if act="list" then
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">插件中心 [<a href="other.asp?open=plus&act=add">添加插件</a>]</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<table width="100%" border="0" cellspacing="0" class="table">
  <tr bgcolor="#DAE2E8">
    <td width="5%" >编号</td>
    <td width="50%">名称</td>
    <td width="45">管理</td>
    </tr>
    <%
	set cn=new config
	rs=cn.plus_list()
	do while not cn.rs.eof
	%>
  <tr bgcolor="#F1F4F7">
    <td width="5%" ><%=cn.rs("id")%></td>
    <td width="5%"><%=cn.rs("plus_name")%></td>
    <td><a href="other.asp?open=plus&act=edit&id=<%=cn.rs("id")%>">修改</a> 
    <%
	if cn.rs("plus_on")=1 then
	%>
    <font color="#cccccc">启动</font> <a href="other.asp?cz=plus&done=Disable&id=<%=cn.rs("id")%>">禁用</a>
    <%else%>
    <a href="other.asp?cz=plus&done=Enable&id=<%=cn.rs("id")%>">启动</a> <font color="#cccccc">禁用</font>
    <%end if%>
    
     <a href="other.asp?cz=plus&done=del&id=<%=cn.rs("id")%>">删除</a></td>
    </tr>
    <%
	cn.rs.movenext
	loop
	%>
</table>
</div>
<%
elseif act="add" then
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">添加插件</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<form action="other.asp?cz=plus&done=add" method="post">
<table width="100%" border="0" cellspacing="0" class="table">
  <tr bgcolor="#F1F4F7">
    <td width="10%" height="40">名称：</td>
    <td>
      <input name="plus_name" type="text" class="kuang" id="plus_name" value="" />  <select name="plus_on" id="plus_on">
        <option value="1" selected>开启</option>
        <option value="0">禁用</option>
      </select>
      
    </td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF" height="40">管理地址：</td>
    <td bgcolor="#FFFFFF"><input name="plus_manage" type="text" id="plus_manage" class="kuang" />  
    </td>
  </tr>
    <tr bgcolor="#F1F4F7">
    <td width="10%"  height="40">安装地址：</td>
    <td><input name="plus_setup" type="text" id="plus_setup" class="kuang" />  
    </td>
  </tr>
  <tr>
    <td colspan="2" align="center" height="40"><input type="submit" value=" 保 存 " style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name="" type="reset" value=" 重 置 " style="border:1px #000000 solid;vertical-align:middle;height:25px"></td>
    </tr>
</table>
</form>
</div>
<%
elseif act="edit" then
id=request.QueryString("id")
set cn=new config
rs=cn.plus_get(id)
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">修改插件</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<form action="other.asp?cz=plus&done=edit&id=<%=id%>" method="post">
<table width="100%" border="0" cellspacing="0" class="table">
  <tr bgcolor="#F1F4F7">
    <td width="10%" height="40">名称：</td>
    <td>
      <input name="plus_name" type="text" class="kuang" id="plus_name" value="<%=cn.rs("plus_name")%>" />
      <select name="plus_on" id="plus_on">
      <%
	  if cn.rs("plus_on")=1 then
	  %>
        <option value="1" selected>开启</option>
        <option value="0">禁用</option>
        <%else%>        
        <option value="0" selected>禁用</option>
        <option value="1">开启</option>
        <%end if%>
      </select></td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFFFFF" height="40">管理地址：</td>
    <td bgcolor="#FFFFFF"><input name="plus_manage" type="text" class="kuang" id="plus_manage" value="<%=cn.rs("plus_manage")%>" />  
    </td>
  </tr>
    <tr bgcolor="#F1F4F7">
    <td width="10%" height="40">安装地址：</td>
    <td><input name="plus_setup" type="text" class="kuang" id="plus_setup" value="<%=cn.rs("plus_setup")%>" />  
    </td>
  </tr>
  <tr>
    <td colspan="2" align="center" height="40"><input type="submit" value=" 保 存 " style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name="" type="reset" value=" 重 置 " style="border:1px #000000 solid;vertical-align:middle;height:25px"></td>
    </tr>
</table>
</form>
</div>
<%
end if
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
'标签操作开始
if cz="bq" then
done=request.QueryString("done")
if done="add" then
set cn=new config
cn.names=request.Form("names")
cn.bq=glyh(request.Form("bq"))
cn.addbiaoqian()
xs_err("<font color='red'>添加失败！</font>")
jump=zx_url("添加成功!","other.asp?open=bq&act=list")
elseif done="edit" then
id=request.Form("id")
set cn=new config
cn.names=request.Form("names")
cn.bq=glyh(request.Form("bq"))
cn.editbiaoqian(id)
xs_err("<font color='red'>修改失败！</font>")
jump=zx_url("修改成功!","other.asp?open=bq&act=list")
elseif done="del" then
id=request.QueryString("id")
set cn=new config
cn.delbiaoqian(id)
xs_err("<font color='red'>删除失败！</font>")
jump=zx_url("删除成功!","other.asp?open=bq&act=list")
end if
'自定义页面操作开始
elseif cz="diy" then
done=request.QueryString("done")
if done="add" then
set cn=new config
cn.names=request.Form("names")
cn.diy=glyh(request.Form("diy"))
cn.addiy()
xs_err("<font color='red'>添加失败！</font>")
jump=zx_url("添加成功!","other.asp?open=diy&act=list")
elseif done="edit" then
id=request.Form("id")
set cn=new config
cn.names=request.Form("names")
cn.diy=glyh(request.Form("diy"))
cn.editdiy(id)
xs_err("<font color='red'>修改失败！</font>")
jump=zx_url("修改成功!","other.asp?open=diy&act=list")
elseif done="del" then
id=request.QueryString("id")
set cn=new config
cn.deldiy(id)
xs_err("<font color='red'>删除失败！</font>")
jump=zx_url("删除成功!","other.asp?open=diy&act=list")
end if
'JS操作开始
elseif cz="js" then
done=request.QueryString("done")
if done="add" then
set cn=new config
cn.js_name=request.Form("js_name")
cn.js_code=html2js(request.Form("js_code"))
cn.jss_add()
xs_err("<font color='red'>添加失败！</font>")
jump=zx_url("添加成功!","other.asp?open=js&act=list")
elseif done="edit" then
id=request.Form("id")
set cn=new config
cn.js_name=request.Form("js_name")
cn.js_code=html2js(request.Form("js_code"))
cn.jss_edit(id)
xs_err("<font color='red'>修改失败！</font>")
jump=zx_url("修改成功!","other.asp?open=js&act=list")
elseif done="del" then
id=request.QueryString("id")
set cn=new config
cn.jss_del(id)
xs_err("<font color='red'>删除失败！</font>")
jump=zx_url("删除成功!","other.asp?open=js&act=list")
end if
'站内连接操作开始
elseif cz="ilink" then
done=request.QueryString("done")
if done="add" then
set cn=new config
cn.link_name=request.Form("link_name")
cn.link_url=request.Form("link_url")
cn.olink_add()
xs_err("<font color='red'>添加失败！</font>")
jump=zx_url("添加成功!","other.asp?open=ilink&act=list")
elseif done="edit" then
id=request.Form("id")
set cn=new config
cn.link_name=request.Form("link_name")
cn.link_url=request.Form("link_url")
cn.olink_edit(id)
xs_err("<font color='red'>修改失败！</font>")
jump=zx_url("修改成功!","other.asp?open=ilink&act=list")
elseif done="del" then
id=request.QueryString("id")
set cn=new config
cn.olink_del(id)
xs_err("<font color='red'>删除失败！</font>")
jump=zx_url("删除成功!","other.asp?open=ilink&act=list")
end if
'插件操作开始
elseif cz="plus" then
done=request.QueryString("done")
if done="add" then
set cn=new config
cn.plus_name=request.Form("plus_name")
cn.plus_manage=request.Form("plus_manage")
cn.plus_setup=request.Form("plus_setup")
cn.plus_on=int(request.Form("plus_on"))
cn.plus_add()
xs_err("<font color='red'>添加失败！</font>")
jump=zx_url("添加成功!","other.asp?open=plus&act=list")
elseif done="edit" then
id=request.QueryString("id")
set cn=new config
cn.plus_name=request.Form("plus_name")
cn.plus_manage=request.Form("plus_manage")
cn.plus_setup=request.Form("plus_setup")
cn.plus_on=int(request.Form("plus_on"))
cn.plus_edit(id)
xs_err("<font color='red'>修改失败！</font>")
jump=zx_url("修改成功!","other.asp?open=plus&act=list")
elseif done="del" then
id=request.QueryString("id")
set cn=new config
cn.plus_del(id)
xs_err("<font color='red'>删除失败！</font>")
jump=zx_url("删除成功!","other.asp?open=plus&act=list")
elseif done="Enable" then
id=request.QueryString("id")
set cn=new config
cn.plus_Enable(id)
xs_err("<font color='red'>启用失败！</font>")
jump=zx_url("启用成功!","other.asp?open=plus&act=list")
elseif done="Disable" then
id=request.QueryString("id")
set cn=new config
cn.plus_Disable(id)
xs_err("<font color='red'>禁用失败！</font>")
jump=zx_url("禁用成功!","other.asp?open=plus&act=list")
end if
'留言操作开始
elseif cz="guest" then
done=request.QueryString("done")
if done="edit" then
id=request.Form("id")
set cn=new config
cn.title=request.Form("title")
cn.name=request.Form("name")
cn.tel=request.Form("tel")
cn.email=request.Form("email")
cn.qq=request.Form("qq")
cn.ip=request.Form("ip")
cn.content=request.Form("content")
cn.times=request.Form("times")
cn.guest_edit(id)
xs_err("<font color='red'>修改失败！</font>")
jump=zx_url("修改成功!","other.asp?open=guest&act=list")
elseif done="del" then
id=request.QueryString("id")
set cn=new config
cn.delguest(id)
xs_err("<font color='red'>删除失败！</font>")
jump=zx_url("删除成功!","other.asp?open=guest&act=list")

elseif done="hf" then
id=request.Form("id")
set cn=new config
ghf=request.Form("hf")
cn.guesthf(id)
xs_err("<font color='red'>回复失败！</font>")
jump=zx_url("回复成功!","other.asp?open=guest&act=list")
end if
elseif cz="review" then
done=request.QueryString("done")
id=request.QueryString("id")
if done="del" then
set rv=new review
rv.did_del(id)
xs_err("<font color='red'>删除失败！</font>")
jump=zx_url("删除成功!","other.asp?open=review&act=list")
end if
'模版管理操作开始
elseif cz="temp" then
done=request.QueryString("done")
if done="add" then
t_name=request.Form("t_name")
t_code=request.Form("t_code")
if chkname(t_name)=true then
response.write "<SCRIPT>alert('模板名 "&t_name&" 已被占用！请修改模板名称');history.go(-1)</SCRIPT>"
else
call do_html(t_code,Server.MapPath(install&"templist/"&temp_url&"/"&t_name),q_Charset)
xs_err("<font color='red'>添加失败！</font>")
jump=zx_url("添加成功!","other.asp?open=temp&act=list")
end if
elseif done="edit" then
t_name=request.Form("t_name")
t_code=request.Form("t_code")
call do_html(t_code,Server.MapPath(install&"templist/"&temp_url&"/"&t_name),q_Charset)
xs_err("<font color='red'>修改失败！</font>")
jump=zx_url("修改成功!","other.asp?open=temp&act=list")
elseif done="del" then
t_name=request.QueryString("t_name")
t_url=Server.MapPath(install&"templist/"&temp_url&"/"&t_name)
call delfile(t_url)
xs_err("<font color='red'>删除失败！</font>")
jump=zx_url("删除成功!","other.asp?open=temp&act=list")
elseif done="copy" then
t_name=request.Form("t_name")
t_code=request.Form("t_code")
if chkname(t_name)=true then
response.write "<SCRIPT>alert('模板名 "&t_name&" 已被占用！请修改模板名称');history.go(-1)</SCRIPT>"
else
call do_html(t_code,Server.MapPath(install&"templist/"&temp_url&"/"&t_name),q_Charset)
xs_err("<font color='red'>复制失败！</font>")
jump=zx_url("复制成功!","other.asp?open=temp&act=list")
end if
end if
end if
%>
