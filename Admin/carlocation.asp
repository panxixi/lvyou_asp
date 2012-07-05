<!--#include file="../inc/utf.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/HuoDong.asp"-->
<!--#include file="../inc/sys.asp"-->
<!--#include file="../inc/carlocation.asp"-->
<!--#include file="../inc/category.asp"-->
<!--#include file="../inc/tfunction.asp"-->
<!--#include file="isadmin.asp"-->
<!--#include file="HuoDongAutoTag.asp" -->

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
		Case "EditIsShow"
			Call EditZhuangtai()
		Case Else
			Call ListView()
	End Select

	Public Sub EditAct()
		set hd=new CarLocation
		hd.id = cInt(Replace(Request.QueryString("id"),"'","''"))
		hd.location = Replace(Request.Form("location"),"'","''")
		hd.isShow = Replace(Request.Form("isShow"),"'","''")
		hd.updateLocation()
		set hd = nothing
		Tourl("?Action=ListView")
	End Sub
	
	Public Sub AddAct()
		set hd=new CarLocation
		hd.location = Replace(Request.Form("location"),"'","''")
		hd.isShow = Replace(Request.Form("isShow"),"'","''")
		hd.addCarLocation()
		set hd = nothing
		Tourl("?Action=ListView")
	End Sub
	
	Public Sub DelAct()
		set hd=new CarLocation
		id = Cint(Request.QueryString("id"))
		hd.delcarLocation(id)
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
		set hd = New CarLocation
		rs = hd.getAllCarLocation()
		if hd.rs.eof then die "暂时没有任何数据"
		hd.rs.PageSize = 15
		Page = CLng(Request("Page"))
	    If Page < 1 Then Page = 1
	    If Page > hd.rs.PageCount Then Page = hd.rs.PageCount
	    hd.rs.AbsolutePage = Page
	%>
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
}
body,td,th {
	font-size: 12px;
}-->
</style>
<body>
  <table border="0" align="center" width="500" cellspacing="1" id="dd">
    <caption>
    <h3>活动乘车地点 <a href="?Action=AddView">添加乘车</a></h3>
    </caption>
    <colgroup>
	    <col />
	    <col class="name" />
	    <col />
	    <col />
    </colgroup>
    <thead>
      <tr>
        <th>ID</th>
        <th>乘车地点</th>
        <th>是否显示</th>
        <th>操作</th>
      </tr>
    </thead>


  <%
  for i=1 to hd.rs.PageSize
  if hd.rs.EOF then Exit For
  %>
      <tbody class="tb1">
      <tr>
        <td><%=hd.rs("ID")%></td>
        <td><%=hd.rs("Location")%>&nbsp;</td>
        <td><%=hd.rs("isShow")%>&nbsp;</td>
        <td><a href="?Action=EditView&id=<%=hd.rs("ID")%>">编辑</a> | <a href="?Action=DelAct&id=<%=hd.rs("ID")%>">删除</a></td>
      </tr>
    </tbody>

  <%
  hd.rs.movenext()
  next
  %>
    </table>
    <center>
    <%
		if page=1 then
			echo "首页 上一页 第"&page&"页 <a href='carLocation.asp?page="&page+1&"'>下一页</a>"
			zuihouyiye = hd.rs.pagecount
			echo "<a href='carLocation.asp?page="&zuihouyiye&"'>尾页</a>"
		elseif page=hd.rs.pagecount then
			echo "<a href='carLocation.asp?page=1'>首页</a> <a href='carLocation.asp?page="&page-1&"'>上一页</a> 第"&page&"页 下一页 尾页"
		else
		zuihouyiye = hd.rs.pagecount
			echo "<a href='carLocation.asp?page=1'>首页</a>&nbsp;<a href='carLocation.asp?page="&page-1&"'>上一页</a> 第"&page&"页 <a href='carLocation.asp?page="&page+1&"'>下一页</a> <a href='carLocation.asp?page="&zuihouyiye&"'>尾页</a>"
		end if
			%></center>
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
			<!--修改活动-->
			<center><h1>修改乘车地点</h1></center>
		<form action="?Action=AddAct" method="post" name="form1" id="form1">
	
  <table border="0" align="center" cellspacing="1" id="dd">
    <caption>
    <h3>活动乘车地点</h3>
    </caption>


    <tbody class="tb1">

      <tr>
        <td>地点</td>
        <td><input type="text" name="location" value=""></td>
      </tr>
      <tr>
        <td>是否显示</td>
        <td><input type="text" name="isShow" value="0"> 0为显示  1为不显示
      </tr>
    </tbody>


</table><center><input type="submit" value="添加"></center>
		</form>
	
	<%
	End Sub
	
	
	
	'#################################################################
	'修改活动的表单页面
	'#################################################################
	Public Sub EditView()
		id = Cint(Request.QueryString("id"))
		set hd=new CarLocation
		hd.getLoc(id)
	%>
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
}
body,td,th {
	font-size: 12px;
}-->
</style>
			<!--修改活动-->
			<center><h1>修改乘车地点</h1></center>
		<form action="?Action=EditAct&ID=<%=id%>" method="post" name="form1" id="form1">
	
  <table border="0" align="center" cellspacing="1" id="dd">
    <caption>
    <h3>活动乘车地点</h3>
    </caption>


    <tbody class="tb1">
      <tr>
        <td>ID</td>
        <td><%=hd.rs("id")%>&nbsp;</td>
      </tr>
      <tr>
        <td>地点</td>
        <td><input type="text" name="location" value="<%=hd.rs("location")%>"></td>
      </tr>
      <tr>
        <td>是否显示</td>
        <td><input type="text" name="isShow" value="<%=hd.rs("isShow")%>"> 0为显示  1为不显示
      </tr>
    </tbody>
  


</table>  <center><input type="submit" value="提交"></center>
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