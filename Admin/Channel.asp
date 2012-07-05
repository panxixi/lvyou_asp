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
</head>
<body>
<%
act=request.querystring("id")
if act=1 then
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
  strurl = "Channel.asp?do=del&cid=" + strid;
  if(!s) {
    alert("请选择要删除的分类!");
    return false;
  }	
  if (confirm("确定要删除这些分类吗？")) {
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
    <td align="center"><p class="pagetitle">分类管理</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<form method="post" name="form1" id="form1">
<table width="100%" border="0" cellspacing="0" class="table">
<tr bgcolor="#DAE2E8">
  <td width="5%"><b>ID</b></td><td width="52%"><b>名称</b></td><td width="30%"><b>操作</b></td><td width="8%"><b>排序</b></td><td width="5%"><input name="" type="checkbox" value=""></td></tr>
<%call MainFl()%>
<tr>
      <td colspan="5" align="center" height="40"><input type="button" value=" 全 选 " onClick="sltall()" style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input type="button" value=" 清 空 " onClick="sltnull()" style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input name="tijiao" type="submit" value=" 删 除 " onClick="SelectChk()" style="border:1px #000000 solid;vertical-align:middle;height:25px"/></td>
    </tr>
</table>
</form>
</div>
<%
elseif act=2 then
cz=request.querystring("cz")
if cz="add" then
aa=int(request.QueryString("cid"))
%>
<script language="JavaScript">
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
<script language="JavaScript">
//绑定页面加载完成事件调用函数
window.onload=page_onload;
function page_onload(){
	//showhide();			//在打开页时，就是判断了
	showhide('linkture','w_link','r_link');
}
</script>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">添加分类</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<form action="Channel.asp?do=add" method="post" name="jnr" id="jnr">
<table width="100%" border="0" cellspacing="0" class="table">
  <tr bgcolor="#F1F4F7">
    <td width="10%" height="40">上级分类<%=qesy%>：</td>
    <td>
      <select name="pcid" id="pcid" class="kuangy">
      <option value="0">顶级分类</option>
        <%call addfl(aa)%>
      </select>
    <tr>
    <td width="10%" height="40">分类名称：</td>
    <td>
      <input name="cname" type="text" id="cname" class="kuang" onBlur="this.className='kuang'" onFocus="this.className='kuang1'"/>
      <select name="types" size="1" id="types">
        <option value="0">分类</option>
        <option value="1">单页</option>
      </select>
      <input name="ons" type="checkbox" id="ons" value="1" />
      显示<input name="linkture" type="checkbox" id="linkture" value="1" onClick="showhide('linkture','w_link','r_link');"/>
      外连
      <input name="px" type="text" id="px" value="0" size="2" />
      排序</td>
  </tr>
    <tr bgcolor="#F1F4F7">
      <td height="60">缩略图：</td>
      <td><input name="img" type="text" id="img" class="kuang" onBlur="this.className='kuang'" onFocus="this.className='kuang1'"/>        <iframe frameborder="0" width="600" height="25" scrolling="no" src="upload.htm.asp" id="ff"></iframe></td>
    </tr>
    <tr id="w_link">
      <td width="10%" height="40">外连地址：</td>
      <td><input name="link" type="text" id="link" class="kuang" onBlur="this.className='kuang'" onFocus="this.className='kuang1'" /></td>
    </tr>
<tr>
  <td colspan="2">
  <table width="100%" border="0" cellspacing="0" id="r_link" class="table">
  <tr>
    <td height="40">分类模版：</td>
    <td><input name="ctemp" type="text" id="ctemp" value="templist/newslist.html" class="kuang" onBlur="this.className='kuang'" onFocus="this.className='kuang1'" /></td>
  </tr>
  <tr bgcolor="#F1F4F7">
    <td height="40">内容模版：</td>
    <td>
      <input name="ntemp" type="text" id="ntemp" value="templist/view.html" class="kuang" onBlur="this.className='kuang'" onFocus="this.className='kuang1'" />
    </td>
  </tr>
  <tr>
    <td width="10%" height="40">分类关键字：</td>
    <td><input name="keyword" type="text" id="keyword" class="kuang" onBlur="this.className='kuang'" onFocus="this.className='kuang1'" /></td>
  </tr>
  <tr bgcolor="#F1F4F7">
    <td width="10%" height="140">分类简介：</td>
    <td><textarea name="info" cols="30" rows="6" id="info" class="kuangx" onBlur="this.className='kuangx'" onFocus="this.className='kuangx1'"></textarea></td>
  </tr></table>
</td></tr>
<tr>
      <td colspan="2" align="center" height="40"><input type="submit" value=" 保 存 " style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name="" type="reset" value=" 重 置 " style="border:1px #000000 solid;vertical-align:middle;height:25px"></td>
    </tr>
</table>
</form>
</div>
<%
elseif cz="edit" then
aa=int(request.QueryString("pcid"))
bb=int(request.QueryString("cid"))
set ct=new category
rs=ct.getcategoryinfo(bb)
%>
<script language="JavaScript">
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
<script language="JavaScript">
//绑定页面加载完成事件调用函数
window.onload=page_onload;
function page_onload(){
	//showhide();			//在打开页时，就是判断了
	showhide('linkture','w_link','r_link');
}
</script>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="header">
  <tr>
    <td align="center"><p class="pagetitle">修改分类</p></td>
  </tr>
</table>
<br />
<br />
<div id="divMain2">
<form action="Channel.asp?do=edit" method="post" name="jnr" id="jnr">
<table width="100%" border="0" cellspacing="0" class="table">
  <tr bgcolor="#F1F4F7">
    <td width="10%" height="40">上级分类<%=qesy%>：</td>
    <td>
      <select name="pcid" id="pcid" class="kuangy">
      <option value="0">顶级分类</option>
        <%call addfl(aa)%>
      </select>
    <tr>
    <td width="10%" height="40">分类名称：</td>
    <td>
      <input name="cname" type="text" class="kuang" id="cname" onFocus="this.className='kuang1'" onBlur="this.className='kuang'" value="<%=ct.rs("cname")%>"/>
      <select name="types" size="1" id="types">
      <%
	  if ct.rs("types")=1 then
	  %>
        <option value="1" selected="selected">单页</option>
        <option value="0">分类</option>
        <%else%>
        <option value="0" selected="selected">分类</option>
        <option value="1">单页</option>
        <%end if%>
      </select>
      <input name="cid" type="hidden" id="cid" value="<%=ct.rs("cid")%>" />
      <%
	  if ct.rs("ons")=1 then
	  %>
      <input name="ons" type="checkbox" id="ons" value="1" checked />
      <%else%>
      <input name="ons" type="checkbox" id="ons" value="1" />
      <%end if%>
      显示
      <%
	  if ct.rs("linkture")=1 then
	  %>
      <input name="linkture" type="checkbox" id="linkture" onClick="showhide('linkture','w_link','r_link');" value="1" checked/>
      <%else%>
      <input name="linkture" type="checkbox" id="linkture" value="1" onClick="showhide('linkture','w_link','r_link');"/>
      <%end if%>
      外连
      <input name="px" type="text" id="px" value="<%=ct.rs("px")%>" size="2" />
      排序</td>
  </tr>
    <tr bgcolor="#F1F4F7">
      <td height="60">缩略图</td>
      <td><input name="img" type="text" class="kuang" id="img" onFocus="this.className='kuang1'" onBlur="this.className='kuang'" value="<%=ct.rs("cimg")%>"/>        <iframe frameborder="0" width="600" height="25" scrolling="no" src="upload.htm.asp" id="ff"></iframe></td>
    </tr>
  
  <tr id="w_link">
    <td width="10%" height="40">外连地址：</td>
    <td><input name="link" type="text" class="kuang" id="link" onFocus="this.className='kuang1'" onBlur="this.className='kuang'" value="<%=ct.rs("link")%>" /></td>
  </tr>
<tr>
  <td colspan="2">
  <table width="100%" border="0" cellspacing="0" id="r_link" class="table">
  <tr>
    <td height="40">分类模版：</td>
    <td><input name="ctemp" type="text" id="ctemp" value="<%=ct.rs("ctemp")%>" class="kuang" onBlur="this.className='kuang'" onFocus="this.className='kuang1'" /></td>
  </tr>
  <tr bgcolor="#F1F4F7">
    <td height="40">内容模版：</td>
    <td>
      <input name="ntemp" type="text" id="ntemp" value="<%=ct.rs("ntemp")%>" class="kuang" onBlur="this.className='kuang'" onFocus="this.className='kuang1'" />
    </td>
  </tr>
  <tr>
    <td width="10%" height="40">分类关键字：</td>
    <td><input name="keyword" type="text" class="kuang" id="keyword" onFocus="this.className='kuang1'" onBlur="this.className='kuang'" value="<%=ct.rs("keyword")%>" /></td>
  </tr>
  <tr bgcolor="#F1F4F7">
    <td width="10%" height="140">分类简介：</td>
    <td><textarea name="info" cols="30" rows="6" id="info" class="kuangx" onBlur="this.className='kuangx'" onFocus="this.className='kuangx1'"><%=ct.rs("info")%></textarea></td>
  </tr></table>
</td></tr>
<tr>
      <td colspan="2" align="center" height="40"><input type="submit" value=" 保 存 " style="border:1px #000000 solid;vertical-align:middle;height:25px"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name="" type="reset" value=" 重 置 " style="border:1px #000000 solid;vertical-align:middle;height:25px"></td>
    </tr>
</table>
</form>
</div>
<%
end if
%>
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
caozuo=request.querystring("do")
select case caozuo
case "add"
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
cn.cimg=request.Form("img")
cn.ctemp=request.Form("ctemp")
cn.ntemp=request.Form("ntemp")
if cn.ctemp="" then
cn.ctemp="templist/newslist.html"
else
cn.ctemp=cn.ctemp
end if
if cn.ntemp="" then
cn.ntemp="templist/view.html"
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
cn.pcid=int(request.Form("pcid"))
if cn.pcid="" then cn.pcid=0
cn.insertcategory()
set cx=new category
rs=cx.get_one_class()
kk=class_arclist(cx.rs("cid"))
call cx.update_curl(kk,cx.rs("cid"))
xs_err("<font color='red'>添加分类失败！</font>")
jump=zx_url("添加分类成功!","Channel.asp?id=1")
case "edit"
cid=int(request.Form("cid"))
pcid=int(request.Form("pcid"))
if cid<>pcid then
set ct=new category
ct.pcid=int(request.Form("pcid"))
ct.cname=request.Form("cname")
ct.linkture=request.Form("linkture")
if ct.linkture="" then ct.linkture=0
ct.link=request.Form("link")
ct.keyword=request.Form("keyword")
ct.info=request.Form("info")
ct.types=request.Form("types")
ct.cimg=request.Form("img")
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

set cx=new category
rs=cx.getcategoryinfo(cid)
kk=class_arclist(cx.rs("cid"))
call cx.update_curl(kk,cx.rs("cid"))

xs_err("<font color='red'>修改分类失败！</font>")
jump=zx_url("修改分类成功!","Channel.asp?id=1")
else
response.Write "<script language='javascript'>alert('上级分类不能是自己！');javascript:window.history.go(-1);</script>"
end if
case "del"
cid=request.querystring("cid")
set ns=new news
If Not ns.HaveCategory(cid) Then
Set ct = New Category
ct.DeleteCategory(cid)
xs_err("<font color='red'>删除失败！</font>")
jump=zx_url("删除分类成功!","Channel.asp?id=1")
else
jump=zx_url("分类下有新闻，不能删除!","Channel.asp?id=1")
end if
end select
%>

