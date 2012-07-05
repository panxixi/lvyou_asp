<%
Class category
Public cid '小类目编号
Public pcid '大类目名称
Public cname '类目名称
Public cimg '分类图片
Public curl '分类路径
Public linkture '是否外连
Public link '外连地址
Public keyword '关键字
Public info '简介
Public types '类型
Public ons '首页显示
Public px '排序
Public ctemp '栏目模版
Public ntemp '内容模版
Public rs       '结果集
Private strsql 'sql语句
'获取单个类目信息
Public Sub getcategoryinfo(cid)
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql="select * from category where cid="&cid
rs.open(sql)
End sub

'获取所有类目信息
Public  Sub getcategorylist()
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql = "SELECT * FROM Category where pcid is null order by px asc"
rs.open sql
End Sub

'统计分类总数
Public  function cate_count()
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql = "SELECT count(*) FROM Category"
rs.open sql
cate_count=rs(0)
End function

'获取所有分类
Public  Sub getlistdir()
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql = "select * from category where linkture=0"
rs.open sql
End Sub

'获取不是外连也不是单页的分类
Public  Sub no_diy_types()
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql = "select * from category where linkture=0 order by px asc,cid asc"
rs.open sql
End Sub

'获取指定类目信息
Public  Sub gettypes(tp)
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql = "SELECT * FROM Category where types="&tp&" and linkture=0 Order By px"
rs.open sql
End sub

'获取分类信息（模版分离）
Public Sub zdfl(row)
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql = "SELECT top "&row&" * FROM Category where ons=1 and pcid=0 Order By px,cid asc"
rs.open sql
End sub
'子分类
Public Sub zzdfl(mrow,cid)
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql = "SELECT top "&mrow&" * FROM Category where pcid="&cid&" and ons=1 Order By px,cid asc"
rs.open sql
End sub

'插入大类目信息
Public Sub insertcategory()
sql="insert into category (pcid,cname,linkture,link,keyword,info,types,ons,px,ctemp,ntemp,cimg) values('"&pcid&"','"&cname&"','"&linkture&"','"&link&"','"&keyword&"','"&info&"','"&types&"','"&ons&"','"&px&"','"&ctemp&"','"&ntemp&"','"&cimg&"')"
conn.execute(sql)
End Sub
'修改类目信息
Public Sub updatecategory(cid)
strsql="update category set pcid='"&pcid&"',cname='"&cname&"',linkture='"&linkture&"',link='"&link&"',ons='"&ons&"',keyword='"&keyword&"',info='"&info&"',px='"&px&"',ctemp='"&ctemp&"',ntemp='"&ntemp&"',types='"&types&"',cimg='"&cimg&"' where cid="&cid
conn.execute(strsql)
End Sub
'修改CURL，在添加分类的时候需要
Public Sub update_curl(c_url,cid)
strsql="update category set curl='"&c_url&"' where cid="&cid
conn.execute(strsql)
End Sub
'获取最新一条信息
Public  Sub get_one_class()
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql = "SELECT top 1 * FROM Category order by cid desc"
rs.open sql
End sub

'删除类目信息
Public Sub deletecategory(cids)
strsql="delete from category where cid in("&cids&")"
conn.execute(strsql)
End Sub
'判断指定类目是否存在
Public Function havecategory(cname)
strsql="select * from category where cname ='"&cname&"'"
Set rs=server.CreateObject("adodb.recordset")
rs.open strsql,conn,1,3
If rs.eof Then
exist=False
Else
exist=True
End If
havecategory=exist
End Function

'修复路径
Public Sub update_url()
strsql="update category set curl='"&curl&"' where cid="&cid
conn.execute(strsql)
End Sub

'判断指定分类中是否存在子分类，返回true or false
public function have_scate(cid)
sql="select * from category where pcid in("&cid&")"
set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,3
if rs.eof then
exist=false
else
exist=true
end if
have_scate=exist
end function

End class
%>