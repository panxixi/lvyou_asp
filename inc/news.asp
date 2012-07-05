<%
'本类用于保存对表category的数据库操作
class news
'表的每个字段对应类的一个成员变量
public newsid
public ntitle
public nkeyword
public ninfo
public ncontent
public posttime
public tuijian
Public pinyin
public types
public cid
public attpic
public readcount
public img
public rs '结果集
private strsql 'SQL语句
'获取单个新闻内容
public sub getnewsinfo(nid)
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql="select news.*,category.* from news left join category on news.cid=category.cid where news.newsid="&nid
rs.open(sql)
end Sub

        Public Sub getSlideNewsList(cid, lb, xx)
        	if cid = "" then
        		Set rs = server.CreateObject("adodb.recordset")
	            Set rs.activeconnection = conn
	            rs.cursortype = 3
	            sql = "select top "&lb&" * from news where "&xx
	            rs.Open(sql)
        		exit sub
        	end if
        	set rs = nothing
            Set rsq = server.CreateObject("adodb.recordset")
            Set rsq.activeconnection = conn
            rsq.cursortype = 3
            sql = "select * from category where pcid="&cid
            rsq.Open(sql)
            Do While Not rsq.EOF
                jj = jj + (rsq("cid")&",")
                rsq.movenext
            Loop
            rsq.Close
            Set rsq = Nothing
            Set rs = server.CreateObject("adodb.recordset")
            Set rs.activeconnection = conn
            rs.cursortype = 3
            sql = "select top "&lb&" news.*,category.* from news,category where news.cid=category.cid and news.cid in ("&jj&cid&") and news.attpic=1 and news.huandeng=1 order by news.newsid desc"
            rs.Open(sql)
        End Sub

'获取新闻列表
'public sub getnewslist(cid,lb,atp,xx)
'jj=get_all_stype(cid)
'set rs=server.CreateObject("adodb.recordset")
'set rs.activeconnection=conn
'rs.cursortype=3
'if atp=1 then
'sql="select top "&lb&" news.*,category.* from news,category where news.cid=category.cid and news.cid in ("&jj&cid&") and attpic=1 "&xx
'else
'sql="select top "&lb&" news.*,category.* from news,category where news.cid=category.cid and news.cid in ("&jj&cid&") "&xx
'end if
'rs.open(sql)
'end sub

        Public Sub getnewslist(cid, lb, atp, xx)
        	if cid = "" then
        		Set rs = server.CreateObject("adodb.recordset")
	            Set rs.activeconnection = conn
	            rs.cursortype = 3
	            sql = "select top "&lb&" * from news "&xx
	            rs.Open(sql)
        		exit sub
        	end if
        	set rs = nothing
            Set rsq = server.CreateObject("adodb.recordset")
            Set rsq.activeconnection = conn
            rsq.cursortype = 3
            sql = "select * from category where pcid="&cid
            rsq.Open(sql)
            Do While Not rsq.EOF
                jj = jj + (rsq("cid")&",")
                rsq.movenext
            Loop
            rsq.Close
            Set rsq = Nothing
            Set rs = server.CreateObject("adodb.recordset")
            Set rs.activeconnection = conn
            rs.cursortype = 3
            If atp = 1 Then
                sql = "select top "&lb&" news.*,category.* from news,category where news.cid=category.cid and news.cid in ("&jj&cid&") and attpic=1 "&xx
            Else
                sql = "select top "&lb&" news.*,category.* from news,category where news.cid=category.cid and news.cid in ("&jj&cid&") "&xx
            End If
            'die sql
            rs.Open(sql)
        End Sub


'获取相关文章列表
public sub getlikelist(keyword,lb,newsid)
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql="select top "&lb&" * from news where newsid<>"&newsid&" and ("&keyword&")"
rs.open(sql)
end sub


'获取所有新闻
public sub allnews()
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql="select news.*,category.* from news left join category on news.cid=category.cid order by news.newsid desc"
rs.open(sql)
end sub

'获取频道列表
public sub pagelist(cid)
jj=get_all_stype(cid)
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
if cid="" then
sql="select news.*,category.* from news left join category on news.cid=category.cid where news.cid in ("&jj&"0) order by posttime desc"
Else
sql="select news.*,category.* from news left join category on news.cid=category.cid where news.cid in ("&jj&cid&") order by posttime desc"
End if
rs.open(sql)
end sub






'搜索出新闻
public sub getnewssql(strsql)
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.curtype=3
sql="select * from news where ntitle="&strsql
rs.open sql
end Sub

'后台搜索新闻
public sub news_search(skey)
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql="select news.*,category.* from news,category where news.cid=category.cid and news.ntitle like '%"&skey&"%'"
rs.open(sql)
end Sub

'快速搜索新闻（批量替换用）
public sub news_search_q(skey)
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql="select newsid,ntitle from news where ntitle like '%"&skey&"%'"
rs.open(sql)
end Sub

'快速搜索新闻内容（批量替换用）
public sub news_search_qc(skey)
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql="select newsid,ncontent from news where ncontent like '%"&skey&"%'"
rs.open(sql)
end sub

'删除指定新闻记录
public sub deletenews(nid)
strsql="delete from news where newsid in("&nid&")"
conn.execute(strsql)
end sub
'判断指定分类中是否存在新闻，返回true or false
public function havecategory(cid)
sql="select * from news where cid in("&cid&")"
set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,3
if rs.eof then
exist=false
else
exist=true
end if
havecategory=exist
end function
'插入新记录
public sub insertnews()
sql="insert into news (cid,ntitle,nkeyword,ninfo,posttime,img,ncontent,attpic)values('"&cid&"','"&ntitle&"','"&nkeyword&"','"&ninfo&"','"&posttime&"','"&img&"','"&ncontent&"','"&attpic&"')"
conn.execute(sql)
end sub
'修改新闻
public sub updatenews(nid)
sql="update news set cid='"&cid&"',ntitle='"&ntitle&"',posttime='"&posttime&"',img='"&img&"',ncontent='"&ncontent&"',attpic='"&attpic&"',nkeyword='"&nkeyword&"',ninfo='"&ninfo&"' where newsid="&nid
conn.execute(sql)
end sub
'阅读次数加1
public sub updatecount(id)
strsql="update news set readcount=readcount+1 where newsid="&id
conn.execute(strsql)
end sub

'推荐文章
public sub updatecom(id)
strsql="update news set tuijian=1 where newsid="&id
conn.execute(strsql)
end Sub

'取消推荐文章
public sub updateqx(id)
strsql="update news set tuijian=0 where newsid="&id
conn.execute(strsql)
end sub

'推荐文章
public sub loo_sql(dm)
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql=""&dm&""
rs.open(sql)
end Sub

'批量移动文章
public sub news_move(id)
strsql="update news set cid='"&cid&"' where newsid in ("&id&")"
conn.execute(strsql)
end Sub

'批量设置文章
public sub news_edit_all(id)
strsql="update news set tuijian='"&tuijian&"' where newsid in ("&id&")"
conn.execute(strsql)
end sub

public function liebiao(a,b)
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql="select * from news where cid=2"
rs.open sql
end function

'RSS总获取新闻列表
public sub getallnewslist(lb)
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql="select top "&lb&" news.*,category.* from news,category where news.cid=category.cid order by posttime desc"
rs.open(sql)
end Sub

'统计新闻总数
Public  function news_count()
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql = "SELECT count(*) FROM news"
rs.open sql
news_count=rs(0)
End Function


        '推荐文章

        Public Sub updatehuandeng(id)
            strsql = "update news set huandeng=1 where newsid="&id
            conn.Execute(strsql)
        End Sub

        '取消推荐文章

        Public Sub updateqxhuandeng(id)
            strsql = "update news set huandeng=0 where newsid="&id
            conn.Execute(strsql)
        End Sub

'统计分类新闻
Public  function news_cid_count(cid)
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql = "SELECT count(*) FROM news where cid="&cid&""
rs.open sql
news_cid_count=rs(0)
End Function

'获取最新一条信息
Public  Sub get_one_news()
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql = "SELECT top 1 news.*,category.* FROM news,category where news.cid=category.cid order by news.newsid desc"
rs.open sql
End Sub



        Public Sub getprenewsinfo(nid)
            Set rs = server.CreateObject("adodb.recordset")
            Set rs.activeconnection = conn
            rs.cursortype = 3
            sql = "select top 1 * from news where newsid<"&nid&" order by newsid desc;"
            rs.Open sql
        End Sub

        Public Sub getnxtnewsinfo(nid)
            Set rs = server.CreateObject("adodb.recordset")
            Set rs.activeconnection = conn
            rs.cursortype = 3
            sql = "select top 1 * from news where newsid>"&nid&" order by newsid asc;"
            rs.Open sql
        End Sub

'修复路径
Public Sub update_url_news(url,id)
strsql="update news set pinyin='"&url&"' where newsid="&id
conn.execute(strsql)
End Sub

'批量替换标题
public sub update_title(aa,id)
strsql="update news set ntitle='"&aa&"' where newsid="&id
conn.execute(strsql)
end Sub

'批量替换内容
public sub update_content(aa,id)
strsql="update news set ncontent='"&aa&"' where newsid="&id
conn.execute(strsql)
end Sub

end class
%>