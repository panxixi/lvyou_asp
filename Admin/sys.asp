<!--#include file="../inc/utf.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/news.asp"-->
<!--#include file="../inc/sys.asp"-->
<!--#include file="../inc/category.asp"-->
<!--#include file="../inc/do.asp"-->
<!--#include file="isadmin.asp"-->
<html>
<head>
<link rel=stylesheet href="styles/advanced/style.css" />
<meta http-equiv="Content-Type" content="text/html; charset=<%=q_Charset%>">
<style type="text/css">
<!--
body,td,th {
	font-family: Arial, Helvetica, sans-serif;
}
-->
</style></head>
<body>
<%
act=request.querystring("id")
if act=2 then
%>
<script language="JavaScript">
function SelectUsers()
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
  strurl = "sys.asp?do=12&uid=" + strid;
  if(!s) {
    alert("请选择要删除的用户!");
    return false;
  }	
  if (confirm("你确定要删除这些用户吗？")) {
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
    <td align="center"><p class="pagetitle">用户管理&nbsp;&nbsp;[&nbsp;<a href="sys.asp?id=16">添加用户</a>&nbsp;]</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<form action="sys.asp?id=2&do=12" method="post" name="form1" id="form1">
  <table width="100%" border="0" cellspacing="0" class="table">
  <tr bgcolor="#DAE2E8">
    <td width="5%">ID</td>
    <td width="35%">用户名</td>
    <td width="15%">性别</td>
    <td width="15%">积分</td>
    <td width="15%">身份证</td>
    <td width="15%">修改 删除</td>
  </tr>
  <%
  set cn=new config
  rs=cn.user_normal()
  if cn.rs.eof then
  %>
    <tr>
    <td colspan="6">暂时没有任何用户！</td>
    </tr>
    <%
    else
    cn.rs.PageSize = 20
	Page = Int(Request.querystring("Page"))
    If Page <= 0 Then Page = 1
	if request.QueryString("page")="" then page=1
	if page>cn.rs.pagecount then page=cn.rs.pagecount
    cn.rs.AbsolutePage = Page
    for i=1 to cn.rs.PageSize
    if cn.rs.EOF then Exit For
    %>
    <%
    if i mod 2=0 then
    %>
  <tr bgcolor="#F1F4F7">
  <%
  else
  %>
  <tr>
  <%
  end if
  %>
    <td><%=cn.rs("userId")%></td>
    <td><%=cn.rs("username")%></td>
    <td><%=cn.rs("sex")%></td>
    <td><%=cn.rs("jifen")%></td>
    <td><%=cn.rs("idcard")%></td>
    <td><a href="sys.asp?id=11&uid=<%=cn.rs("userId")%>">修改</a> <a href="sys.asp?do=12&uid=<%=cn.rs("userId")%>">删除</a> <input name="cate" type="checkbox" id="<%=cn.rs("userId")%>" value="<%=cn.rs("userId")%>"></td>
  </tr>
  <%
  cn.rs.movenext()
  next
  end if
  %>
   <tr>
    <td colspan="6" align="center">
    <%
if page=1 then
response.Write "首页&nbsp;上一页&nbsp;第"&page&"页&nbsp;<a href='sys.asp?id=2&page="&page+1&"'>下一页</a>&nbsp;<a href='sys.asp?id=2&page="&cn.rs.pagecount&"'>尾页</a>"
elseif page=cn.rs.pagecount then
response.Write "<a href='sys.asp?id=2&page=1'>首页</a>&nbsp;<a href='sys.asp?id=2&page="&page-1&"'>上一页</a>&nbsp;第"&page&"页&nbsp;下一页&nbsp;尾页"
else
response.Write "<a href='sys.asp?id=2&page=1'>首页</a>&nbsp;<a href='sys.asp?id=2&page="&page-1&"'>上一页</a>&nbsp;第"&page&"页&nbsp;<a href='sys.asp?id=2&page="&page+1&"'>下一页</a>&nbsp;<a href='sys.asp?id=2&page="&cn.rs.pagecount&"'>尾页</a>"
end if
    %>
    </td>
    </tr>
     <tr>
    <td colspan="6" align="center" height="40"><input type="button" value=" 全 选 " onClick="sltall()" style="border:1px #000000 solid;vertical-align:middle;height:25px" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input type="button" value=" 清 空 " onClick="sltnull()" style="border:1px #000000 solid;vertical-align:middle;height:25px" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input name="tijiao" type="submit" value=" 删 除 " onClick="SelectUsers()" style="border:1px #000000 solid;vertical-align:middle;height:25px" /></td>
    </tr>
</table>

  </form>
</div>
<%
elseif act=3 then
%>
<script language="JavaScript">
function SelectUsers()
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
  strurl = "sys.asp?do=13&uid=" + strid;
  if(!s) {
    alert("请选择要删除的用户!");
    return false;
  }	
  if (confirm("你确定要删除这些用户吗？")) {
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
    <td align="center"><p class="pagetitle">管理员管理&nbsp;&nbsp;[&nbsp;<a href="sys.asp?id=15">添加管理员</a>&nbsp;]</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<form action="sys.asp?id=3&do=13" method="post" name="form1" id="form1">
  <table width="100%" border="0" cellspacing="0" class="table">
  <tr bgcolor="#DAE2E8">
    <td width="5%">ID</td>
    <td>用户名</td>
    <td width="15%">修改 删除</td>
  </tr>
  <%
  set cn=new config
  rs=cn.user()
  if cn.rs.eof then
  %>
    <tr>
    <td colspan="3">暂时没有任何用户！</td>
    </tr>
    <%
    else
    cn.rs.PageSize = 20
	Page = Int(Request.querystring("Page"))
    If Page <= 0 Then Page = 1
	if request.QueryString("page")="" then page=1
	if page>cn.rs.pagecount then page=cn.rs.pagecount
    cn.rs.AbsolutePage = Page
    for i=1 to cn.rs.PageSize
    if cn.rs.EOF then Exit For
    %>
    <%
    if i mod 2=0 then
    %>
  <tr bgcolor="#F1F4F7">
  <%
  else
  %>
  <tr>
  <%
  end if
  %>
    <td><%=cn.rs("id")%></td>
    <td><%=cn.rs("admin_name")%></td>
    <td><a href="sys.asp?id=14&uid=<%=cn.rs("id")%>">修改</a> <a href="sys.asp?do=13&uid=<%=cn.rs("id")%>">删除</a> <input name="cate" type="checkbox" id="<%=cn.rs("id")%>" value="<%=cn.rs("id")%>"></td>
  </tr>
  <%
  cn.rs.movenext()
  next
  end if
  %>
   <tr>
    <td colspan="3" align="center">
    <%
if page=1 then
response.Write "首页&nbsp;上一页&nbsp;第"&page&"页&nbsp;<a href='sys.asp?id=3&page="&page+1&"'>下一页</a>&nbsp;<a href='sys.asp?id=3&page="&cn.rs.pagecount&"'>尾页</a>"
elseif page=cn.rs.pagecount then
response.Write "<a href='sys.asp?id=3&page=1'>首页</a>&nbsp;<a href='sys.asp?id=3&page="&page-1&"'>上一页</a>&nbsp;第"&page&"页&nbsp;下一页&nbsp;尾页"
else
response.Write "<a href='sys.asp?id=3&page=1'>首页</a>&nbsp;<a href='sys.asp?id=3&page="&page-1&"'>上一页</a>&nbsp;第"&page&"页&nbsp;<a href='sys.asp?id=3&page="&page+1&"'>下一页</a>&nbsp;<a href='sys.asp?id=3&page="&cn.rs.pagecount&"'>尾页</a>"
end if
    %>
    </td>
    </tr>
     <tr>
    <td colspan="3" align="center" height="40"><input type="button" value=" 全 选 " onClick="sltall()" style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input type="button" value=" 清 空 " onClick="sltnull()" style="border:1px #000000 solid;vertical-align:middle;height:25px"/>
           &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
           <input name="tijiao2" type="submit" value=" 删 除 " onClick="SelectUsers()" style="border:1px #000000 solid;vertical-align:middle;height:25px"/></td>
    </tr>
</table>

  </form>
</div>
<%
elseif act=4 then
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">修复数据库</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<table width="100%" border="0" cellspacing="0" class="table">
  <tr bgcolor="#F1F4F7">
    <td height="40">
<%
conn.close()
mdbpath = server.mappath(install&dbpath)
Set fso=Server.CreateObject("Scripting.FileSystemObject")
If fso.FileExists(mdbPath) Then
Set Engine = CreateObject("JRO.JetEngine")

	If request("boolIs") = "97" Then
		Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mdbpath, _
		"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mdbpath & "_temp.asp;" _
		& "Jet OLEDB:Engine Type=" & JET_3X
	Else
		Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mdbpath, _
		"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mdbpath & "_temp.asp"
	End If

fso.CopyFile mdbpath & "_temp.asp",mdbpath
fso.DeleteFile(mdbpath & "_temp.asp")
Set fso = nothing
Set Engine = nothing


response.write "成功修复数据库!"
else
response.write "错误:数据库名称不正确.请重试!"	
end if
%>
    </td>
    </tr>

</table>
</div>
<%
elseif act=7 then
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">重置路径</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<table width="100%" border="0" cellspacing="0" class="table">
  <tr bgcolor="#F1F4F7">
    <td width="15%" height="auto" align="center"><p>当路径发生改变的时，需要重置路径<br>重置路径会占用比较大的资源，和需要比较多的时间<br>如果无特殊需要，请不要重置路径！</p>
      <p><a href="url.asp?active=class">重置分类路径</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="url.asp?active=cont">重置内容路径</a></p></td></tr></table>
</div>
<%
ElseIf act=8 Then
Application.Contents.RemoveAll
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">更新缓存</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<table width="100%" border="0" cellspacing="0" class="table">
<tr bgcolor="#F1F4F7"><td height="40">更新缓存成功！</td></tr></table></div>
<%
elseif act=14 then
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">修改管理员密码</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<%
id=request.querystring("uid")
set cn=new config
rs=cn.admin_info(id)
%>
<form action="sys.asp?do=14" method="post">
<table width="100%" border="0" cellspacing="0" class="table">
  <tr bgcolor="#F1F4F7">
    <td width="15%" height="40">用户名：</td>
    <td><%=cn.rs("admin_name")%><input name="id" type="hidden" id="id" value="<%=cn.rs("id")%>"></td>
  </tr>
  <tr>
    <td height="40">密码：</td>
    <td><input type="password" name="admin_password" id="admin_password" class="kuangy" onBlur="this.className='kuangy'" onFocus="this.className='kuangy1'"></td>
  </tr>
  <tr>
    <td height="40" colspan="2" align="center"><input name="bot1" type="submit" id="bot1" value=" 保 存 "  style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
       
          <input type="reset" name="but2" id="but2" value=" 重 置 " style="border:1px #000000 solid;vertical-align:middle;height:25px"></td>
    </tr>
</table>
</form>
</div>
<%
elseif act=11 then
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">修改用户信息</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<%
id=request.querystring("uid")
set cn=new config
rs=cn.users_info(id)
%>
<form action="sys.asp?do=11" method="post">
<table width="100%" border="0" cellspacing="0" class="table">
  <tr bgcolor="#F1F4F7">
    <td width="15%" height="40">用户名：</td>
    <td><%=cn.rs("username")%><input name="id" type="hidden" id="id" value="<%=cn.rs("id")%>"></td>
  </tr>
  <tr>
    <td height="40">密码：</td>
    <td><input type="password" name="password" id="password" class="kuangy" onBlur="this.className='kuangy'" onFocus="this.className='kuangy1'"> * 密码不修改请留空！</td>
  </tr>
  <tr bgcolor="#F1F4F7">
    <td height="40">性别：</td>
    <td><input name="sex" type="text" class="kuangy" id="sex" onFocus="this.className='kuangy1'" onBlur="this.className='kuangy'" value="<%=cn.rs("sex")%>"></td>
  </tr>
  <tr>
    <td height="40">Email：</td>
    <td><input name="email" type="text" class="kuangy" id="email" onFocus="this.className='kuangy1'" onBlur="this.className='kuangy'" value="<%=cn.rs("email")%>"></td>
  </tr>
  <tr bgcolor="#F1F4F7">
    <td height="40">QQ：</td>
    <td><input name="qq" type="text" class="kuangy" id="qq" onFocus="this.className='kuangy1'" onBlur="this.className='kuangy'" value="<%=cn.rs("qq")%>"></td>
  </tr>
  <tr>
    <td height="40">电话：</td>
    <td><input name="tel" type="text" class="kuangy" id="tel" onFocus="this.className='kuangy1'" onBlur="this.className='kuangy'" value="<%=cn.rs("tel")%>"></td>
  </tr>
  <tr bgcolor="#F1F4F7">
    <td height="40">真实名字：</td>
    <td><input name="truename" type="text" class="kuangy" id="truename" onFocus="this.className='kuangy1'" onBlur="this.className='kuangy'" value="<%=cn.rs("truename")%>"></td>
  </tr>
  <tr>
    <td height="40">地址：</td>
    <td><input name="address" type="text" class="kuangy" id="address" onFocus="this.className='kuangy1'" onBlur="this.className='kuangy'" value="<%=cn.rs("address")%>"></td>
  </tr>
  <tr>
    <td height="40" colspan="2" align="center"><input name="bot1" type="submit" id="bot1" value=" 保 存 "  style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
       
          <input type="reset" name="but2" id="but2" value=" 重 置 " style="border:1px #000000 solid;vertical-align:middle;height:25px"></td>
    </tr>
</table>
</form>
</div>
<%
elseif act=16 then
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">添加用户</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<form action="sys.asp?do=16" method="post">
<table width="100%" border="0" cellspacing="0" class="table">
  <tr bgcolor="#F1F4F7">
    <td width="15%" height="40">用户名：</td>
    <td><input name="username" type="text" class="kuangy" id="username" onFocus="this.className='kuangy1'" onBlur="this.className='kuangy'"></td>
  </tr>
  <tr>
    <td height="40">密码：</td>
    <td><input type="password" name="password" id="password" class="kuangy" onBlur="this.className='kuangy'" onFocus="this.className='kuangy1'"> * 密码不修改请留空！</td>
  </tr>
  <tr bgcolor="#F1F4F7">
    <td height="40">性别：</td>
    <td><input name="sex" type="text" class="kuangy" id="sex" onFocus="this.className='kuangy1'" onBlur="this.className='kuangy'"></td>
  </tr>
  <tr>
    <td height="40">Email：</td>
    <td><input name="email" type="text" class="kuangy" id="email" onFocus="this.className='kuangy1'" onBlur="this.className='kuangy'"></td>
  </tr>
  <tr bgcolor="#F1F4F7">
    <td height="40">QQ：</td>
    <td><input name="qq" type="text" class="kuangy" id="qq" onFocus="this.className='kuangy1'" onBlur="this.className='kuangy'"></td>
  </tr>
  <tr>
    <td height="40">电话：</td>
    <td><input name="tel" type="text" class="kuangy" id="tel" onFocus="this.className='kuangy1'" onBlur="this.className='kuangy'"></td>
  </tr>
  <tr bgcolor="#F1F4F7">
    <td height="40">真实名字：</td>
    <td><input name="truename" type="text" class="kuangy" id="truename" onFocus="this.className='kuangy1'" onBlur="this.className='kuangy'"></td>
  </tr>
  <tr>
    <td height="40">地址：</td>
    <td><input name="address" type="text" class="kuangy" id="address" onFocus="this.className='kuangy1'" onBlur="this.className='kuangy'"></td>
  </tr>
  <tr>
    <td height="40" colspan="2" align="center"><input name="bot1" type="submit" id="bot1" value=" 保 存 "  style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
       
          <input type="reset" name="but2" id="but2" value=" 重 置 " style="border:1px #000000 solid;vertical-align:middle;height:25px"></td>
    </tr>
</table>
</form>
</div>
<%
elseif act=15 then
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">添加管理员</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<form action="sys.asp?do=15" method="post">
<table width="100%" border="0" cellspacing="0" class="table">
  <tr bgcolor="#F1F4F7">
    <td width="15%" height="40">用户名：</td>
    <td><input type="text" name="admin_name" id="admin_name" class="kuangy" onBlur="this.className='kuangy'" onFocus="this.className='kuangy1'"></td>
  </tr>
  <tr>
    <td height="40">密码：</td>
    <td><input type="password" name="admin_password" id="admin_password" class="kuangy" onBlur="this.className='kuangy'" onFocus="this.className='kuangy1'"></td>
  </tr>
  <tr>
    <td height="40" colspan="2" align="center"><input name="bot1" type="submit" id="bot1" value=" 保 存 "  style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
       
          <input type="reset" name="but2" id="but2" value=" 重 置 " style="border:1px #000000 solid;vertical-align:middle;height:25px"></td>
    </tr>
</table>
</form>
</div>
<%
elseif act=88 then
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">参数设置</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<form action="sys.asp?do=88" method="post" name="form9" id="form9">
  <table width="100%" border="0" cellspacing="0" class="table">
  <tr bgcolor="#F1F4F7"><td width="15%" height="40">数据库：</td>
  <td>
    <input name="dbpath" type="text" class="kuang" id="dbpath" onFocus="this.className='kuang1'" onBlur="this.className='kuang'" value="<%=dbpath%>" ></td>
</tr>
<tr><td width="15%" height="40">缓存时间：</td>
  <td><input name="huancun" type="text" class="kuang" id="huancun" onFocus="this.className='kuang1'" onBlur="this.className='kuang'" value="<%=huancun%>" >

      <select name="c_open" id="c_open">
      <%
	  if c_open=0 then
	  %>
        <option value="1">开启缓存</option>
        <option value="0" selected>关闭缓存</option>
    <%else%>
                <option value="1" selected>开启缓存</option>
        <option value="0">关闭缓存</option>
        <%end if%>
      </select>

    *以分钟为单位</td>
</tr>
<tr bgcolor="#F1F4F7"><td width="15%" height="40">运行模式：</td>
  <td><input name="html" type="text" class="kuang" id="html" onFocus="this.className='kuang1'" onBlur="this.className='kuang'" value="<%=html%>" >
    *0=动态，1=静态，2=伪静态</td>
</tr>
<tr><td width="15%" height="40">模版路径</td>
  <td><input name="temp_url" type="text" class="kuang" id="temp_url" onFocus="this.className='kuang1'" onBlur="this.className='kuang'" value="<%=temp_url%>" >*模版文件夹</td>
</tr>
<tr bgcolor="#F1F4F7"><td width="15%" height="40">验证码：</td>
  <td><input name="yzcode" type="text" class="kuang" id="yzcode" onFocus="this.className='kuang1'" onBlur="this.className='kuang'" value="<%=yzcode%>" >*后台登陆验证码</td>
</tr>
<tr><td width="15%" height="60" rowspan="2">分类命名规则：</td>
  <td height="40"><input name="class_url" type="text" class="kuang" id="class_url" onFocus="this.className='kuang1'" onBlur="this.className='kuang'" value="<%=class_url%>" >*更改后需要重置路径</td>
</tr>
<tr>
  <td>全拼：{pinyin}，{pin_yin}，{py}，{p_y}，{cid}</td>
</tr>
<tr bgcolor="#F1F4F7"><td width="15%" height="40" rowspan="2">内容页命名规则：</td>
  <td height="40"><input name="content_url" type="text" class="kuang" id="content_url" onFocus="this.className='kuang1'" onBlur="this.className='kuang'" value="<%=content_url%>" >*更改后需要重置路径</td>
</tr>
<tr bgcolor="#F1F4F7">
  <td>全拼：{pinyin}，{pin_yin}，{py}，{p_y}，{id}</td>
</tr>
<tr>
      <td colspan="2" align="center" height="40"><input name="bot1" type="submit" id="bot1" value=" 保 存 "  style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
       
          <input type="reset" name="but2" id="but2" value=" 重 置 " style="border:1px #000000 solid;vertical-align:middle;height:25px">
        </td>
    </tr>
</table>
</form>
</div>
<%
else
set cn=new config
rs=cn.getconfig()
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">基本设置</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<form action="sys.asp?do=1" method="post" name="form1" id="form1">
  <table width="100%" border="0" cellspacing="0" class="table">
    <tr bgcolor="#F1F4F7">
      <td width="20%" height="40">网站名称：</td>
      <td><input name="webname" type="text" id="webname" value="<%=cn.rs("webname")%>" class="kuang" onBlur="this.className='kuang'" onFocus="this.className='kuang1'" /></td>
    </tr>
    <tr>
      <td height="40">网站首页：</td>
      <td><input name="index" type="text" id="index" value="<%=cn.rs("index")%>" class="kuang" onBlur="this.className='kuang'" onFocus="this.className='kuang1'"/></td>
    </tr>
    <tr bgcolor="#F1F4F7">
      <td height="40">首页模版：</td>
      <td><input name="itemp" type="text" id="itemp" value="<%=cn.rs("itemp")%>" class="kuang" onBlur="this.className='kuang'" onFocus="this.className='kuang1'"/></td>
    </tr>
    <tr>
      <td height="40">站长邮箱：</td>
      <td><input name="email" type="text" id="email" value="<%=cn.rs("email")%>" class="kuang" onBlur="this.className='kuang'" onFocus="this.className='kuang1'"/></td>
    </tr>
    <tr bgcolor="#F1F4F7">
      <td height="40">关 键 字 ：</td>
      <td><input name="keywords" type="text" id="keywords" value="<%=cn.rs("keywords")%>" class="kuang" onBlur="this.className='kuang'" onFocus="this.className='kuang1'"/></td>
    </tr>
    <tr>
      <td height="40">网站简介：</td>
      <td><input name="info" type="text" id="info" value="<%=cn.rs("info")%>" class="kuang" onBlur="this.className='kuang'" onFocus="this.className='kuang1'"/></td>
    </tr>
    <tr bgcolor="#F1F4F7">
      <td height="140">网站版权：</td>
      <td><textarea name="copyright" cols="30" rows="6" id="copyright" class="kuangx" onBlur="this.className='kuangx'" onFocus="this.className='kuangx1'"><%=cn.rs("copyright")%></textarea></td>
    </tr>
    <tr>
      <td colspan="2" align="center" height="40"><input name="bot1" type="submit" id="bot1" value=" 保 存 "  style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
       
          <input type="reset" name="but2" id="but2" value=" 重 置 " style="border:1px #000000 solid;vertical-align:middle;height:25px">
        </td>
    </tr>
  </table></form>
</div>

<%
end if
caozuo=request.querystring("do")
select case caozuo
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
jump=zx_url("保存成功","sys.asp?id=1")
case 11
set cn=new config
id=request.form("id")
cn.userpwd=md5(request.form("password"))
cn.sex=request.form("sex")
cn.email=request.form("email")
cn.qq=request.form("qq")
cn.tel=request.Form("tel")
cn.truename=request.Form("truename")
cn.address=request.Form("address")
if cn.userpwd="" then
cn.uinfo_nopass(id)
xs_err("<font color='red'>修改失败！</font>")
jump=zx_url("修改成功","sys.asp?id=2")
else
cn.uinfo_pass(id)
xs_err("<font color='red'>修改失败！</font>")
jump=zx_url("修改成功","sys.asp?id=2")
end if
case 16
set cn=new config
cn.username=request.Form("username")
cn.userpwd=md5(request.form("password"))
cn.sex=request.form("sex")
cn.email=request.form("email")
cn.qq=request.form("qq")
cn.tel=request.Form("tel")
cn.truename=request.Form("truename")
cn.address=request.Form("address")
cn.add_users()
xs_err("<font color='red'>添加失败！</font>")
jump=zx_url("添加成功","sys.asp?id=2")
case 12
set cn=new config
cn.deluser_n(request.querystring("uid"))
xs_err("<font color='red'>删除失败！</font>")
jump=zx_url("删除成功","sys.asp?id=2")
case 13
uid=trim(request.querystring("uid"))
if uid=trim(session("id")) then
jump=zx_url("不能删除自己！","sys.asp?id=3")
else
set cn=new config
cn.deluser(uid)
xs_err("<font color='red'>删除失败！</font>")
jump=zx_url("删除成功","sys.asp?id=3")
end if
case 14
set cn=new config
cn.admin_password=md5(trim(request.form("admin_password")))
cn.edituser(request.form("id"))
xs_err("<font color='red'>修改失败！</font>")
jump=zx_url("修改成功","sys.asp?id=3")
case 15
set cn=new config
cn.admin_name=request.form("admin_name")
cn.admin_password=md5(request.form("admin_password"))
if cn.admin_name="" or cn.admin_password="" then
jump=zx_url("用户名和密码不能为空！","sys.asp?id=15")
end if
if cn.haveuser(admin_name)=true then
cn.adduser
xs_err("<font color='red'>添加失败！</font>")
jump=zx_url("添加成功","sys.asp?id=3")
else
jump=zx_url("不能添加相同的用户!","sys.asp?id=15")
end if
case 88
cs=read_html(server.mappath(install&"inc/const.asp"),q_Charset)
cs=replace(cs,"temp_url="""&temp_url&"""","temp_url="""&request.Form("temp_url")&"""")
cs=replace(cs,"dbpath="""&dbpath&"""","dbpath="""&request.Form("dbpath")&"""")
cs=replace(cs,"yzcode="""&yzcode&"""","yzcode="""&request.Form("yzcode")&"""")
cs=replace(cs,"huancun="""&huancun&"""","huancun="""&request.Form("huancun")&"""")
cs=replace(cs,"html="""&html&"""","html="""&request.Form("html")&"""")
cs=replace(cs,"class_url="""&class_url&"""","class_url="""&request.Form("class_url")&"""")
cs=replace(cs,"content_url="""&content_url&"""","content_url="""&request.Form("content_url")&"""")
cs=replace(cs,"c_open="""&c_open&"""","c_open="""&request.Form("c_open")&"""")
Call do_html(cs,Server.MapPath(install&"inc/const.asp"),q_Charset)
xs_err("<font color='red'>修改失败！</font>")
jump=zx_url("修改成功","sys.asp?id=88")
end select
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