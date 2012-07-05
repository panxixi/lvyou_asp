<%
Class review
Public did
Public posttime
Public poster
Public dcontent
Public newsid
Public rs
Private strsql

'获取所有回复
Public Sub did_all()
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql="select discuss.*,news.* from discuss left join news on discuss.newsid=news.newsid"
rs.open(sql)
End Sub

'获取指定新闻的回复
Public Sub did_get(id)
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql="select * from discuss where newsid="&id&" order by did desc"
rs.open(sql)
End Sub

'删除回复
Public Sub did_del(id)
strsql="delete from discuss where did in("&id&")"
conn.execute(strsql)
End Sub

'插入新记录
public sub did_add()
sql="insert into discuss (poster,dcontent,posttime,newsid)values('"&poster&"','"&dcontent&"','"&posttime&"','"&newsid&"')"
conn.execute(sql)
end Sub

Public  function review_id_count(id)
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql = "SELECT count(*) FROM discuss where newsid="&id&""
rs.open sql
review_id_count=rs(0)
End Function

End class
%>