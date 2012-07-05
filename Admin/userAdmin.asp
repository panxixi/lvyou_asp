<!--#include file="../inc/utf.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/news.asp"-->
<!--#include file="../inc/sys.asp"-->
<!--#include file="../inc/users.asp"-->
<!--#include file="../inc/tfunction.asp"-->
<!--#include file="../inc/category.asp"-->
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
  strurl = "?action=del&uid=" + strid;
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
<p align="left">
<form action="?id=2" name="search" method="GET">
	<input type="text" name="searchKey" value="姓名|电话|身份证|网名" style="width:200px;height:24px;" onfocus="if(this.value=='姓名|电话|身份证|网名'){this.value='';}" onblur="if(this.value==''){this.value=='姓名|电话|身份证|网名';}"><input type="hidden" name="id" value="2"><input type="submit" value="查找" />
</form>
</p>
<form action="" method="post" name="form1" id="form1">
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
  searchKey = request.querystring("searchKey")
  set cn=new config
  if searchKey<>"" then
	  rs = cn.user_search(searchKey)
  else
	  rs=cn.user_normal()
  end if
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
    <td><a href="?action=edit&uid=<%=cn.rs("userId")%>">修改</a> <a href="?action=del&uid=<%=cn.rs("userId")%>">删除</a> <input name="cate" type="checkbox" id="<%=cn.rs("userId")%>" value="<%=cn.rs("userId")%>"></td>
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
response.Write "首页&nbsp;上一页&nbsp;第"&page&"页&nbsp;<a href='?id=2&page="&page+1&"'>下一页</a>&nbsp;<a href='?id=2&page="&cn.rs.pagecount&"'>尾页</a>"
elseif page=cn.rs.pagecount then
response.Write "<a href='?id=2&page=1'>首页</a>&nbsp;<a href='?id=2&page="&page-1&"'>上一页</a>&nbsp;第"&page&"页&nbsp;下一页&nbsp;尾页"
else
response.Write "<a href='?id=2&page=1'>首页</a>&nbsp;<a href='?id=2&page="&page-1&"'>上一页</a>&nbsp;第"&page&"页&nbsp;<a href='?id=2&page="&page+1&"'>下一页</a>&nbsp;<a href='?id=2&page="&cn.rs.pagecount&"'>尾页</a>"
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
end if
Dim Action
Set u = New Users
Action = Request.QueryString("Action")
Select case Action
	Case "edit"
		call edit()
	Case "editAct"
		call editAct()
	Case "del"
		call del()
End select

Public Sub edit()
	uid=Request.QueryString("uid")
	u.getuserById(uid)
%>
<form action="?Action=editAct&uid=<%=uid%>" method="POST">
用户名:<input type="text" name="userName" value="<%=u.rs("userName")%>" /><br/>
密 码:
<input type="password" name="password" value="<%=u.rs("password")%>" /><br/>
真 名:
<input type="text" name="RealName" value="<%=u.rs("RealName")%>" /><br/>
性 别:<%=u.rs("sex")%><br/>
电 话:
<input type="text" name="Phone" value="<%=u.rs("Phone")%>" /><br/>
身份证:<input type="text" name="IDCard" value="<%=u.rs("IDCard")%>" /><br/>
积 分:
<input type="text" name="Jifen" value="<%=u.rs("Jifen")%>" /><br/>
<input type="submit" value="确定提交"/>
</form>
<%
End Sub

Public Sub editAct()
	uid=Request.QueryString("uid")
	userName=Request.Form("userName")
	password=md5(Request.Form("password"))
	RealName=Request.Form("RealName")
	Phone=Request.Form("Phone")
	IDCard=Request.Form("IDCard")
	Jifen=Request.Form("Jifen")
	u.userId = uid
	u.userName = userName
	u.password = password
	u.RealName = RealName
	u.phone = Phone
	u.idcard = IDCard
	u.Jifen = Jifen
	u.updateUserHt()
	ToUrl("?id=2")
End Sub

Public Sub del()
	uid=Request.QueryString("uid")
	u.delusers(uid)
	ToUrl("?id=2")
End Sub
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