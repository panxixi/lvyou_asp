<%
class config
public webname
public index
public itemp
public email
public keywords
public info
public copyright
public id
public admin_name
public admin_password
Public link_name
Public link_url
public js_name
public js_code
public title
public name
public names
public bq
public diy
public ip
public tel
public qq
Public hf
Public username
public userpwd
Public truename
Public address
public sex
public content
public times
Public plus_name
Public plus_manage
Public plus_setup
Public plus_on
public rs '结果集
private strsql 'SQL语句

'获取系统设置
public sub getconfig()
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
sql="select * from system where id=1"
rs.open sql
end sub

'修改系统设置
public sub upconfig()
sql="update system set webname='"&webname&"',[index]='"&index&"',itemp='"&itemp&"',email='"&email&"',keywords='"&keywords&"',info='"&info&"',copyright='"&copyright&"' where id=1"
conn.execute(sql)
end sub

'获取留言列表
public sub guest()
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql="select * from guest order by times desc"
rs.open(sql)
end sub

'删除留言
public sub delguest(id)
sql="delete from guest where id in ("&id&")"
conn.execute(sql)
end sub

'添加留言
public sub addguest()
sql="insert into guest(title,name,tel,email,qq,content,times,ip) values ('"&title&"','"&name&"','"&tel&"','"&email&"','"&qq&"','"&content&"','"&times&"','"&ip&"')"
conn.execute(sql)
end sub

'回复留言
public sub guesthf(id)
sql="update guest set hf='"&ghf&"' where id="&id&""
conn.execute(sql)
end sub

'修改留言
public sub guest_edit(id)
sql="update guest set title='"&title&"',name='"&name&"',tel='"&tel&"',email='"&email&"',qq='"&qq&"',content='"&content&"',ip='"&ip&"',times='"&times&"' where id="&id
conn.execute(sql)
end sub

'获取指定留言
public sub getguest(id)
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql="select * from guest where id="&id
rs.open(sql)
end Sub

'获取插件列表
public sub plus_list()
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql="select * from plus order by id desc"
rs.open(sql)
end Sub

'获取可用插件列表
public sub plus_list_on()
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql="select * from plus where plus_on=1 order by id desc"
rs.open(sql)
end Sub

'添加插件
public sub plus_add()
sql="insert into plus(plus_name,plus_manage,plus_setup,plus_on) values ('"&plus_name&"','"&plus_manage&"','"&plus_setup&"',"&plus_on&")"
conn.execute(sql)
end Sub

'获取指定插件
public sub plus_get(id)
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql="select * from plus where id="&id
rs.open(sql)
end Sub

'修改插件
public sub plus_edit(id)
sql="update plus set plus_name='"&plus_name&"',plus_manage='"&plus_manage&"',plus_setup='"&plus_setup&"',plus_on="&plus_on&" where id="&id
conn.execute(sql)
end Sub

'设插件为启用
public sub plus_Enable(id)
sql="update plus set plus_on=1 where id="&id
conn.execute(sql)
end Sub

'设插件为禁用
public sub plus_Disable(id)
sql="update plus set plus_on=0 where id="&id
conn.execute(sql)
end Sub

'删除插件
public sub plus_del(id)
sql="delete from plus where id in ("&id&")"
conn.execute(sql)
end sub

'获取自定义标签列表
public sub biaoqian()
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql="select * from bq"
rs.open(sql)
end sub

'获取指定自定义标签
public sub getbiaoqian(id)
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql="select * from bq where id="&id
rs.open(sql)
end sub

'修改自定义标签
public sub editbiaoqian(id)
sql="update bq set bq='"&bq&"',[names]='"&names&"' where id="&id
conn.execute(sql)
end sub

'删除自定义标签
public sub delbiaoqian(id)
sql="delete from bq where id in ("&id&")"
conn.execute(sql)
end sub

'添加自定义标签
public sub addbiaoqian()
sql="insert into bq(bq,[names])values('"&bq&"','"&names&"')"
conn.execute(sql)
end sub

'获得自定义页面列表
public sub diylist()
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql="select * from diy"
rs.open(sql)
end sub

'获得指定自定义页面
public sub getdiy(id)
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql="select * from diy where id="&id
rs.open(sql)
end sub

'修改自定义页面
public sub editdiy(id)
sql="update diy set diy='"&diy&"',[names]='"&names&"' where id="&id
conn.execute(sql)
end sub

'删除自定义页面
public sub deldiy(id)
sql="delete from diy where id in ("&id&")"
conn.execute(sql)
end sub

'添加自定义页面
public sub addiy()
sql="insert into diy (diy,[names]) values ('"&diy&"','"&names&"')"
conn.execute(sql)
end Sub

'获得指定站内连接
public sub getolink(id)
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql="select * from olink where id="&id
rs.open(sql)
end sub

'获得站内连接列表
public sub olink_list()
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql="select * from olink"
rs.open(sql)
end sub

'添加站内连接
public sub olink_add()
sql="insert into olink (link_name,link_url)values('"&link_name&"','"&link_url&"')"
conn.execute(sql)
end Sub

'修改站内连接
public sub olink_edit(id)
sql="update olink set link_name='"&link_name&"',link_url='"&link_url&"' where id="&id
conn.execute(sql)
end Sub

'删除站内连接
public sub olink_del(id)
sql="delete from olink where id in ("&id&")"
conn.execute(sql)
end sub

'获得指定JS
public sub getjss(id)
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql="select * from jss where id="&id
rs.open(sql)
end sub

'获得JS列表
public sub jss_list()
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql="select * from jss"
rs.open(sql)
end sub

'添加站内连接
public sub jss_add()
sql="insert into jss (js_name,js_code)values('"&js_name&"','"&js_code&"')"
conn.execute(sql)
end Sub

'修改站内连接
public sub jss_edit(id)
sql="update jss set js_name='"&js_name&"',js_code='"&js_code&"' where id="&id
conn.execute(sql)
end Sub

'删除站内连接
public sub jss_del(id)
sql="delete from jss where id in ("&id&")"
conn.execute(sql)
end sub

'获得管理员列表
public sub user()
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql="select * from admin"
rs.open(sql)
end sub

'获得用户列表
public sub user_normal()
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql="select * from users"
rs.open(sql)
end Sub

'获得用户列表
public sub user_search(searchkey)
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql="select * from users where username like '%"&searchkey&"%' or Realname like '%"&searchkey&"%' or phone like '"&searchkey&"'"
rs.open(sql)
end Sub

'添加用户
public sub add_users()
sql="insert into users (username,userpwd,sex,email,qq,tel,truename,address) values ('"&username&"','"&userpwd&"','"&sex&"','"&email&"','"&qq&"','"&tel&"','"&truename&"','"&address&"')"
conn.execute(sql)
end sub
'修改用户信息（不改密码）
public sub uinfo_nopass(id)
sql="update users set sex='"&sex&"',email='"&email&"',qq='"&qq&"',tel='"&tel&"',truename='"&truename&"',address='"&address&"' where id="&id
conn.execute(sql)
end Sub

'修改用户信息（改密码）
public sub uinfo_pass(id)
sql="update users set userpwd='"&userpwd&"',sex='"&sex&"',email='"&email&"',qq='"&qq&"',tel='"&tel&"',truename='"&truename&"',address='"&address&"' where id="&id
conn.execute(sql)
end Sub

'删除用户
public sub deluser_n(id)
sql="delete from users where id in ("&id&")"
conn.execute(sql)
end sub

'获得用户信息
public sub users_info(id)
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql="select * from users where id="&id
rs.open(sql)
end sub

'获得管理员信息
public sub admin_info(id)
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql="select * from admin where id="&id
rs.open(sql)
end sub

'判断用户是否存在
public function haveuser(user)
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql="select * from admin where admin_name='"&admin_name&"'"
rs.open sql
if rs.eof then haveuser=true
end function

'添加管理员
public sub adduser()
sql="insert into admin (admin_name,admin_password) values ('"&admin_name&"','"&admin_password&"')"
conn.execute(sql)
end sub

'删除管理员
public sub deluser(id)
sql="delete from admin where id in ("&id&")"
conn.execute(sql)
end sub

'修改管理员密码
public sub edituser(id)
sql="update admin set admin_password='"&admin_password&"' where id="&id
conn.execute(sql)
end Sub

'统计用户总数
Public  function user_count()
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql = "SELECT count(*) FROM users"
rs.open sql
user_count=rs(0)
End Function

'统计评论总数
Public  function pinglun_count()
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql = "SELECT count(*) FROM discuss"
rs.open sql
pinglun_count=rs(0)
End Function

'统计留言总数
Public  function guest_count()
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql = "SELECT count(*) FROM guest"
rs.open sql
guest_count=rs(0)
End Function

'统计标签总数
Public  function bq_count()
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql = "SELECT count(*) FROM bq"
rs.open sql
bq_count=rs(0)
End Function

'统计tag总数
Public  function tag_count()
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql = "SELECT count(*) FROM tags"
rs.open sql
tag_count=rs(0)
End Function

end class
%>