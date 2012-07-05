<!--#include file="utf.asp"-->
<!--#include file="conn.asp"-->
<!--#include file="s.asp"-->
<!--#include file="news.asp"-->
<!--#include file="category.asp"-->
<!--#include file="sys.asp"-->
<!--#include file="tfunction.asp"-->
<%
js_id=request.QueryString("id")
gg=dm2js(js_temp(js_id))

Set myCache = new Cache
myCache.name="index" '定义缓存名称 
if myCache.valid then '如果缓存有效
 qcms=myCache.value '读取缓存内容
else
 qcms="document.writeln('"&gg&"');" '大量内容，可以是非常耗时大量数据库查询记录集
 myCache.add qcms,dateadd("n",huancun,now) '将内容赋值给缓存，并设置缓存有效期是当前时间＋1000分钟
end if
Response.Write(qcms)
'myCache.makeEmpty()

set clsCache=nothing '释放对象
%>


