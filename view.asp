<!--#include file="inc/utf.asp"-->
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/s.asp"-->
<!--#include file="inc/news.asp"-->
<!--#include file="inc/category.asp"-->
<!--#include file="inc/sys.asp"-->
<!--#include file="inc/review.asp"-->
<!--#include file="inc/baoming.asp"-->
<!--#include file="inc/huodong.asp"-->
<!--#include file="inc/users.asp"-->
<!--#include file="inc/tfunction.asp"-->
<%
newsid=request.QueryString("newsid")
Set myCache = new Cache
myCache.name="view"&newsid '定义缓存名称 
if myCache.valid then '如果缓存有效
 qcms=myCache.value '读取缓存内容
else
 qcms=view_temp(newsid) '大量内容，可以是非常耗时大量数据库查询记录集
 myCache.add qcms,dateadd("n",huancun,now) '将内容赋值给缓存，并设置缓存有效期是当前时间＋1000分钟
end if
Response.Write(qcms)
If c_open=0 Then myCache.makeEmpty()

set clsCache=nothing '释放对象
%>
