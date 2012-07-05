<!--#include file="../inc/utf.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/news.asp"-->
<!--#include file="../inc/baoming.asp"-->
<!--#include file="../inc/sys.asp"-->
<!--#include file="../inc/category.asp"-->
<!--#include file="../inc/do.asp"-->
<!--#include file="isadmin.asp"-->
<html>
<head>
<link rel=stylesheet href="styles/advanced/style.css" />
<meta http-equiv="Content-Type" content="text/html; charset=<%=q_Charset%>">
</head>
<body>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">后台管理</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
  <table width="100%" border="0" cellspacing="0" class="table">
      <tr bgcolor="#DAE2E8">
      <td colspan="4"><b><center>站内统计摘要</center></b></td>
    </tr>
    <tr bgcolor="#F1F4F7">
      <td width="20%">当前用户：</td>
      <td width="30%"><%=session("admin_name")%></td>
      <td width="20%">当前版本：</td>
      <td width="30%"><%=Version%>&nbsp;&nbsp;<%=UCase(q_Charset)%>&nbsp;&nbsp;<%=UCase(data_type)%></td>
    </tr>
    <tr>
      <td>分类总数：</td>
      <td>
	  <%
	  Set ct=new category
	  response.write ct.cate_count()
	  %></td>
      <td>当前域名：</td>
      <td><%=request.ServerVariables("HTTP_HOST")%></td>
    </tr>
    <tr bgcolor="#F1F4F7">
      <td>文章总数：</td>
      <td>	  <%
	  Set ns=new news
	  response.write ns.news_count()
	  %></td>
      <td>当前时间：</td>
      <td><%=Now()%></td>
    </tr>
    <tr>
      <td>用户总数：</td>
      <td>	  <%
	  Set cn=new config
	  response.write cn.user_count()
	  %></td>
      <td>留言总数：</td>
      <td><%
	  Set cn=new config
	  response.write cn.guest_count()
	  %></td>
    </tr>
    <tr bgcolor="#F1F4F7">
      <td>当前模式：</td>
      <td><%
	  If html=0 Then
	  response.write "动态"
	  Elseif html=1 Then
	  response.write "静态"
	  ElseIf html=2 Then
	  response.write "伪静态"
	  End if
	  %></td>
      <td>标签总数：</td>
      <td><%
	  Set cn=new config
	  response.write cn.bq_count()
	  %></td>
    </tr>
    <tr>
      <td>当前样式：</td>
      <td><%=temp_url%></td>
      <td>TAG总数：</td>
      <td><%
	  Set cn=new config
	  response.write cn.tag_count()
	  %></td>
    </tr>
    <tr bgcolor="#F1F4F7">
      <td>当前路径：</td>
      <td><%=install%></td>
      <td>评论总数：</td>
      <td><%
	  Set cn=new config
	  response.write cn.pinglun_count()
	  %></td>
    </tr>
     <tr bgcolor="#999999">
      <td>所有会员所拥有的积分：</td>
      <td><%
      set bm = new baoming
      response.write bm.getAlljifen()
      %></td>
      <td>导出所有参加活动会员电话号码：</td>
      <td><input type="button" onclick="document.getElementById('phone').style.display='';" value="导出电话"><%
	  phone = bm.getAllphone()
	  set bm = nothing
	  set cn = nothing
	  %></td>
    </tr>
  </table>
  <div id="phone" style="display:none;"><textarea cols="35" rows="20"><%=phone%></textarea><br/><input type="button" onclick="document.getElementById('phone').style.display='none';" value="关闭"></div>
  <div id="idcard" style="display:none;"><textarea cols="35" rows="20"><%=idcard%></textarea><br/><input type="button" onclick="document.getElementById('idcard').style.display='none';" value="关闭"></div>
</div>
<br />
<br />
<table width="100%" border="0" cellpadding="0" cellspacing="0" align="center" class="table">
  <tr>
    <td align="center" height="40">Powered By <a href="http://www.q-cms.cn"><%=Version%></a> &nbsp;&nbsp;<%=UCase(q_Charset)%>&nbsp;&nbsp;<%=UCase(data_type)%>&copy; 2009 All Rights Reserved  <div class="no_view" style="display:none;"><script src="http://s87.cnzz.com/stat.php?id=1707573&web_id=1707573&show=pic" language="JavaScript" charset="gb2312"></script></div>
</td>
</table>
</body>
</html>
