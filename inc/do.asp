<%
'获取指定分类下的所有分类
function get_all_stype(cid)
set rsq=server.CreateObject("adodb.recordset")
set rsq.activeconnection=conn
rsq.cursortype=3
if cid="" then
sql="select * from category where linkture=0 and types=0"
rsq.open(sql)
do while not rsq.eof
kk=kk+cstr(rsq("cid"))&","
kk=kk
rsq.movenext
Loop
else
sql="select * from category where pcid="&cstr(cid)&" and linkture=0 and types=0"
rsq.open(sql)
do while not rsq.eof
kk=kk+cstr(rsq("cid"))&","
kk=kk+get_all_stype(cstr(rsq("cid")))
rsq.movenext
Loop
end if
rsq.close
set rsq=Nothing
get_all_stype=kk
End function

'获得上一篇
function pre_news(id,cid)
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql = "SELECT top 1 * FROM [news] where newsid<"&id&" and cid="&cid&" order by newsid desc"
rs.open sql
If Not rs.eof then
pre_news=rs("newsid")
Else
pre_news="xxxx"
End if
End Function
'获取下一篇
function next_news(id,cid)
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql = "SELECT top 1 * FROM news where newsid>"&id&" and cid="&cid&" order by newsid asc"
rs.open sql
If Not rs.eof then
next_news=rs("newsid")
Else
next_news="xxxx"
End if
End Function
'*************************************************
'调用函数
'*************************************************
'数组替换函数
function replaces(str,aa,bb) 
Dim zz(20)
zz(0)=str
for j=0 to ubound(aa)
zz(j+1)=replace(zz(j),aa(j),bb(j))
replaces=zz(j+1)
Next

end Function

'bq内循环匹配
function nbq(xx,yy)
set reg=new regexp
reg.IgnoreCase = False 
reg.Global = True 
reg.pattern=""&yy&"=([0-9]+)"
set matches=reg.execute(xx)
for each match in matches
aa=match.submatches(0)
next
nbq=aa
end Function

'截取字符串长度函数
Function cutStr(ntlen,strlen)
If ntlen="" Then
cutStr=""
else
dim l,t,c 
l=len(ntlen) 
t=0 
for i=1 to l 
c=Abs(Asc(Mid(ntlen,i,1))) 
if c>255 then 
t=t+2 
else 
t=t+1 
end if 
if t>=strlen then 
cutStr=left(ntlen,i) 
exit for 
else 
cutStr=ntlen 
end if 
next 
cutStr=replace(cutStr,chr(10),"")
End if
end Function

'生成多级目录
Function MakeNewsDir(strPath) ' As Boolean
        On Error Resume Next

        Dim astrPath, ulngPath, i, strTmpPath
        Dim objFSO

        If InStr(strPath, "\") <=0 Or InStr(strPath, ":") <= 0 Then
                AutoCreateFolder = False
                Exit Function
        End If
        Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
        If objFSO.FolderExists(strPath) Then
                AutoCreateFolder = True
                Exit Function
        End If
        astrPath = Split(strPath, "\")
        ulngPath = UBound(astrPath)
        strTmpPath = ""
        For i = 0 To ulngPath
                strTmpPath = strTmpPath & astrPath(i) & "\"
                If Not objFSO.FolderExists(strTmpPath) Then
                        ' 创建
                        objFSO.CreateFolder(strTmpPath)
                End If
        Next
        Set objFSO = Nothing
        If Err = 0 Then
                MakeNewsDir = True
        Else
                MakeNewsDir = False
        End If
End Function 



'地址转向函数
function zx_url(smg,url)
response.write "<script>alert('"&smg&"');window.location.href='"&url&"';</script>"
response.end
end Function

'切源码
function cut(n,s)
ss=len(s)
xx=mid(s,n,ss)
cut=xx
end function

'过滤大括号{}
function dkh(d)
d=replace(d,"{","")
d=replace(d,"}","")
dkh=d
end function

'html转JS
function html2js(dm)
dm1=replace(dm,"""","[""]")
dm2=replace(dm1,"/","[/]")
html2js=dm2
end function

'JS转HTML
function js2html(dm)
dm1=replace(dm,"[""]","""")
dm2=replace(dm1,"[/]","/")
js2html=dm2
end function

'DM2JS
function dm2js(dm)
dm=replace(dm,"[""]","\""")
dm=replace(dm,"[/]","\/")
dm= Replace(dm, CHR(13),"") 
dm= Replace(dm, CHR(10),"")
dm= Replace(dm, CHR(9),"")
dm2js=dm
end Function

'普通转JS
function dm_to_js(dm)
dm=replace(dm,"""","\""")
dm=replace(dm,"/","\/")
dm=replace(dm,"'","\'")
dm= Replace(dm, CHR(13),"") 
dm= Replace(dm, CHR(10),"")
dm= Replace(dm, CHR(9),"")
dm_to_js=dm
end function

'万能标签正则过滤
Function wnbq(all) 
    Set ra = New RegExp 
    ra.IgnoreCase = True 
    ra.Global = True 
    ra.Pattern = "{sql:(.*?)}" 
    wnbq= ra.replace(all,"$1") 
End Function

'截取SQL语句需要
function jiesql(xx,yy)
set reg=new regexp
reg.IgnoreCase = False 
reg.Global = True 
reg.Pattern = ""&yy&"=""([\s\S.]*?)"""
set aa=reg.execute(xx)
for each match in aa
aa=match.submatches(0)
next
jiesql=aa
end Function

'过滤大括号{}
function dkh(d)
d=replace(d,"{","")
d=replace(d,"}","")
dkh=d
end Function

'替换单引号
function glyh(d)
d=replace(d,chr(39),"&#39;")
d=replace(d,chr(34),"&#34;")
glyh=d
end Function

'显示错误函数
sub xs_err(cz)
  If  Err.Number  <>  0  Then            
        Response.Write  cz&"<br>错误原因："&Err.Description
        Response.End        
  End  If
end Sub

'获取模板列表
function mblist()
set fso = Server.CreateObject("Scripting.FileSystemObject")
set folderFso = fso.GetFolder(server.MapPath(install&"templist/"&temp_url&"/"))
set filesFso = folderFso.files
for each file in filesFso
if right(UCase(file.name),4)="HTML" then
fl=fl&file.name&","
end if
next
mblist=fl
end Function

'数组排序
Function Sort(ary)
KeepChecking = TRUE
Do Until KeepChecking = FALSE
KeepChecking = FALSE
For I = 0 to UBound(ary)
If I = UBound(ary) Then Exit For
If ary(I) > ary(I+1) Then
FirstValue = ary(I)
SecondValue = ary(I+1)
ary(I) = SecondValue
ary(I+1) = FirstValue
KeepChecking = TRUE
End If
Next
Loop
Sort = ary
End Function

'添加关键字
function keyword_add(key)
keyword=split(key,",")
for i=0 to ubound(keyword)
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql="select * from tags where tag_name='"&trim(keyword(i))&"'"
rs.open sql
if rs.eof then
sql="insert into tags (tag_name,tag_num) values ('"&trim(keyword(i))&"','1')"
conn.execute(sql)
set rsc=server.CreateObject("adodb.recordset")
set rsc.activeconnection=conn
rsc.cursortype=3
sql="select * from tags order by id desc"
rsc.open(sql)
k=rsc("id")
else
set rsq=server.CreateObject("adodb.recordset")
set rsq.activeconnection=conn
rsq.cursortype=3
sql="select * from tags where tag_name='"&trim(keyword(i))&"'"
rsq.open(sql)
sql="update tags set tag_num='"&rsq("tag_num")+1&"' where id="&rsq("id")
conn.execute(sql)
k=rsq("id")
end if
bb=bb+",{"&k&"}"
next
keyword_add=cut(2,bb)
end function

'修改关键字
function keyword_edit(key)
keyword=split(key,",")
for i=0 to ubound(keyword)
set rs=server.CreateObject("adodb.recordset")
set rs.activeconnection=conn
rs.cursortype=3
sql="select * from tags where tag_name='"&trim(keyword(i))&"'"
rs.open sql
if rs.eof then
sql="insert into tags (tag_name,tag_num) values ('"&trim(keyword(i))&"','1')"
conn.execute(sql)
set rsc=server.CreateObject("adodb.recordset")
set rsc.activeconnection=conn
rsc.cursortype=3
sql="select * from tags order by id desc"
rsc.open(sql)
k=rsc("id")
else
set rsq=server.CreateObject("adodb.recordset")
set rsq.activeconnection=conn
rsq.cursortype=3
sql="select * from tags where tag_name='"&trim(keyword(i))&"'"
rsq.open(sql)
sql="update tags set tag_num='"&rsq("tag_num")&"' where id="&rsq("id")
conn.execute(sql)
k=rsq("id")
end if
bb=bb+",{"&k&"}"
next
keyword_edit=cut(2,bb)
end function

'添加关键字
function nkeyword_add(nkey)
  keyword=split(nkey,",")
  for i=0 to ubound(keyword)
    set rs=server.CreateObject("adodb.recordset")
    set rs.activeconnection=conn
    rs.cursortype=3
    sql="select * from tags where id="&dkh(keyword(i))
    rs.open sql
    kk=kk+","&rs("tag_name")
  next
  nkeyword_add=cut(2,kk)
end Function

'读取文件
Function read_html(url,charset)
set stm=Server.CreateObject("adodb.stream") 
stm.Type=1 'adTypeBinary，按二进制数据读入 
stm.Mode=3 'adModeReadWrite ,这里只能用3用其他会出错 
stm.Open 
stm.LoadFromFile url
stm.Position=0 '把指针移回起点 
stm.Type=2 '文本数据 
stm.Charset=charset
read_html = stm.ReadText
stm.Close 
set stm=nothing 
End Function

'检查模板名称是否可用
function chkname(mbnames)
chkname=False
kk=split(mblist(),",")
for i=0 to ubound(kk)
if UCase(kk(i))=UCase(mbnames) then
chkname=True
exit for
end if
next
end function

'生成文件
sub do_html(t_code,url,charset)
set stm=server.CreateObject("adodb.stream")
stm.Type=2
stm.mode=3
stm.charset=charset
stm.open
stm.WriteText t_code
stm.SaveToFile url,2
stm.flush
stm.Close
set stm=nothing
end Sub

'删除文件
Function delfile(t_url)
     dim fso
     Set fso = Server.CreateObject("Scripting.FileSystemObject")
     if fso.FileExists(t_url) then
     fso.DeleteFile(t_url)             
     end if
     Set fso = nothing
End Function

'禁止站外提交
public sub siteout()
server_v1=Cstr(Request.ServerVariables("HTTP_REFERER")) 
server_v2=Cstr(Request.ServerVariables("SERVER_NAME")) 
if mid(server_v1,8,len(server_v2))<>server_v2 then 
response.write "<br><br><center>" 
response.write " " 
response.write "你提交的路径有误，禁止从站点外部提交数据请不要乱改参数！" 
response.end 
end if 
end sub

'分类管理——无限分类
        Sub MainFl()
            Dim Rs
			types=0
			set rs=server.CreateObject("adodb.recordset")
            set rs.activeconnection=conn
            rs.cursortype=3
            sql="select * from category where pcid=0 order by px asc,cid asc"
			rs.open(sql)
            If Not Rs.Eof Then
                Do While Not Rs.Eof
                    %>
                    <tr bgcolor="#F1F4F7" height="30">
    <td width="5%"><%=rs("cid")%></td>
    <td width="52%"><b><%=rs("cname")%></b>&nbsp;&nbsp;&nbsp;
	<%
	if rs("linkture")=1 then
	response.Write "&nbsp;<font color='red'>外</font>"
	Else
	response.Write "&nbsp;<font color='#cccccc'>外</font>"
	end if
	if rs("types")=1 then
	response.Write "&nbsp;<font color='#0000FF'>单</font>"
	Else
	response.Write "&nbsp;<font color='#cccccc'>单</font>"
	end if
	if rs("ons")=1 then
	response.Write "&nbsp;<font color='#009900'>显</font>"
	Else
	response.Write "&nbsp;<font color='#cccccc'>显</font>"
	end if
	%>
	</td>
    <td width="30%">
<%
if rs("linkture")=1 then
%>
	<font color='cccccc'>添加子类 | 添加内容</font> | <a href="Channel.asp?id=2&cz=edit&pcid=<%=rs("pcid")%>&cid=<%=rs("cid")%>" target="_self">设置</a> | <a href="Channel.asp?do=del&cid=<%=rs("cid")%>">删除</a>
	<%else%>
<a href="Channel.asp?id=2&cz=add&cid=<%=rs("cid")%>">添加子类</a> | <a href="content.asp?active=add&cid=<%=rs("cid")%>">添加内容</a> | <a href="Channel.asp?id=2&cz=edit&pcid=<%=rs("pcid")%>&cid=<%=rs("cid")%>" target="_self">设置</a> | <a href="Channel.asp?do=del&cid=<%=rs("cid")%>">删除</a>
	<%End if%>
	</td>
    <td width="8%"><%=rs("px")%></td><td align="center" width="5"><input name="Cate" type="checkbox" id="<%=rs("cid")%>" /></td>
</tr>
<%
Call Subfl(Rs("cid"),"&nbsp;&nbsp;├") '循环子级分类
                    
                Rs.MoveNext
                Loop
            End If
            Set Rs = Nothing
        End Sub
        '定义子级分类
        Sub SubFl(FID,StrDis)
			set rs1=server.CreateObject("adodb.recordset")
            set rs1.activeconnection=conn
            rs1.cursortype=3
            Set Rs1 = Conn.Execute("SELECT * FROM category WHERE pcid = " & FID &" order by px asc,cid asc")
            If Not Rs1.Eof Then
                Do While Not Rs1.Eof
                    %>
                    <tr>
    <td width="5%"><%=rs1("cid")%></td>
    <td width="52%">&nbsp;<%=StrDis%>&nbsp;<%=rs1("cname")%>
	
	<%
	if rs1("linkture")=1 then
	response.Write "&nbsp;<font color='red'>外</font>"
	Else
	response.Write "&nbsp;<font color='#cccccc'>外</font>"
	end if
	if rs1("types")=1 then
	response.Write "&nbsp;<font color='#0000FF'>单</font>"
	Else
	response.Write "&nbsp;<font color='#cccccc'>单</font>"
	end if
	if rs1("ons")=1 then
	response.Write "&nbsp;<font color='#009900'>显</font>"
	Else
	response.Write "&nbsp;<font color='#cccccc'>显</font>"
	end if
	%>
	
	</td>
    <td width="30%"><a href="Channel.asp?id=2&cz=add&cid=<%=rs1("cid")%>">添加子类</a> | <a href="content.asp?active=add&cid=<%=rs1("cid")%>">添加内容</a> | <a href="Channel.asp?id=2&cz=edit&pcid=<%=rs1("pcid")%>&cid=<%=rs1("cid")%>" target="_self">设置</a> | <a href="Channel.asp?do=del&cid=<%=rs1("cid")%>">删除</a> </td>
    <td width="8%"><%=rs1("px")%></td><td align="center" width="5"><input name="Cate" type="checkbox" id="<%=rs1("cid")%>" /></td>
  </tr>
<%
                Call SubFl(Trim(Rs1("cid")),"&nbsp;&nbsp; &nbsp;&nbsp;" & Strdis) '递归子级分类
                Rs1.Movenext:Loop
                If Rs1.Eof Then
                    Rs1.Close
                    Exit Sub
                End If
            End If
            Set Rs1 = Nothing
        End Sub

'获取分类
function pagenum(code,ccid)
set fso=server.createobject("scripting.filesystemobject")
'die code
set mytext=fso.OpenTextFile(server.mappath(code)) 
str=mytext.readall
mytext.close
set reg=new regexp
reg.IgnoreCase = False 
reg.Global = True 
reg.Pattern = "{pagelist([\s\S.]*?)}([\s\S.]*?){pageend}"
set aa=reg.execute(str)
for each match in aa
bq=match.submatches(0)
size=nbq(bq,"size")
if size="" then
size=20
else
size=size
end if
next
set ns=new news
rs=ns.pagelist(ccid)
if not ns.rs.eof then
ns.rs.pagesize=size
pagenum=ns.rs.pagecount
else
pagenum=1
end if
end Function
'**************************************************************
'添加分类——无限分类
        Sub addfl(qesy)
			set rs=server.CreateObject("adodb.recordset")
            set rs.activeconnection=conn
            rs.cursortype=3
            sql="select * from category where pcid=0 and linkture=0 and types=0 order by px asc,cid asc"
			rs.open(sql)
            If Not Rs.Eof Then
                Do While Not Rs.Eof
					%>
                    <%
					if rs("cid")=qesy then
					%>
                    <option value="<%=rs("cid")%>" selected="selected"><%=rs("cname")%></option>
                    <%else%>
                    <option value="<%=rs("cid")%>"><%=rs("cname")%></option>
                    <%
					end if
					%>
                   
<%
 '循环子级分类
                Call addSubfl(Rs("cid"),"&nbsp;&nbsp;├",qesy)    
                rs.MoveNext()
                Loop
            End If
            Set Rs = Nothing
        End Sub
        '定义子级分类
        Sub addSubfl(FID,StrDis,qesy)
			set rs1=server.CreateObject("adodb.recordset")
            set rs1.activeconnection=conn
            rs1.cursortype=3
            Set Rs1 = Conn.Execute("SELECT * FROM category WHERE pcid = " & FID &" order by px asc,cid asc")
            If Not Rs1.Eof Then
                Do While Not Rs1.Eof
				if rs1("cid")=qesy then
                    %>
                    <option value="<%=rs1("cid")%>" selected="selected"><%=StrDis&rs1("cname")%></option>
                    <%else%>
                    <option value="<%=rs1("cid")%>"><%=StrDis&rs1("cname")%></option>
                    <%end if%>
                    
<%
                Call addSubfl(Trim(Rs1("cid")),"&nbsp;&nbsp; &nbsp;&nbsp;" & Strdis,qesy) '递归子级分类
                Rs1.Movenext:Loop
                If Rs1.Eof Then
                    Rs1.Close
                    Exit Sub
                End If
            End If
            Set Rs1 = Nothing
        End Sub
'**************************************************************
'跳转分类——无限分类
        Sub jumpfl()
			set rs=server.CreateObject("adodb.recordset")
            set rs.activeconnection=conn
            rs.cursortype=3
            sql="select * from category where pcid=0 and linkture=0 and types=0 order by px asc,cid asc"
			rs.open(sql)
            If Not Rs.Eof Then
                Do While Not Rs.Eof
					%>
					<option value="content.asp?active=list&cid=<%=rs("cid")%>"><%=rs("cname")%></option>
<%
 '循环子级分类
                Call jumpSubfl(Rs("cid"),"&nbsp;&nbsp;├")    
                rs.MoveNext()
                Loop
            End If
            Set Rs = Nothing
        End Sub
        '定义子级分类
        Sub jumpSubfl(FID,StrDis)
			set rs1=server.CreateObject("adodb.recordset")
            set rs1.activeconnection=conn
            rs1.cursortype=3
            Set Rs1 = Conn.Execute("SELECT * FROM category WHERE pcid = " & FID &" order by px asc,cid asc")
            If Not Rs1.Eof Then
                Do While Not Rs1.Eof
                    %>
					<option value="content.asp?active=list&cid=<%=rs1("cid")%>"><%=StrDis&rs1("cname")%></option>
                    
<%
                Call jumpSubfl(Trim(Rs1("cid")),"&nbsp;&nbsp; &nbsp;&nbsp;" & Strdis) '递归子级分类
                Rs1.Movenext:Loop
                If Rs1.Eof Then
                    Rs1.Close
                    Exit Sub
                End If
            End If
            Set Rs1 = Nothing
        End Sub
'***************************************************************************************
'函数调用需要
'***************************************************************************************
'检查FSO
Function IsObjInstalled(strClassString)
 On Error Resume Next
 IsObjInstalled = False
 Err = 0
 Dim xTestObj
 Set xTestObj = Server.CreateObject(strClassString)
 If 0 = Err Then IsObjInstalled = True
 Set xTestObj = Nothing
 Err = 0
End Function
'压缩数据库
JET_3X = 4
Function CompactDB(kkpath, boolIs97)
Dim fso, Engine, strkkpath
strkkpath = left(kkpath,instrrev(kkpath,""))
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(kkpath) Then
Set Engine = CreateObject("JRO.JetEngine")
If boolIs97 = "True" Then
Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & kkpath, _
"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strkkpath & "temp.mdb;" _
& "Jet OLEDB:Engine Type=" & JET_3X
Else
Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & kkpath, _
"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strkkpath & "temp.mdb"
End If
fso.CopyFile strkkpath & "temp.mdb",kkpath
fso.DeleteFile(strkkpath & "temp.mdb")
Set fso = nothing
Set Engine = nothing
CompactDB = "你的数据库, " & kkpath & ", 已经被压缩" & vbCrLf
Else
CompactDB = "你输入的数据库路径或名称未找到，请重试" & vbCrLf
End If
End Function

'********************************************************
'路径（动态，静态，伪静）
'********************************************************
'首页HERE路径
Function index_here()
If html=0 Then
index_here="<a href='"&install&"index.asp'>主页</a> >"
Else
index_here="<a href='"&install&"index.html'>主页</a> >"
End if
End Function

'留言页路径
Function guest_here()
If html=0 Then
guest_here="<a href='"&install&"index.asp'>首页</a>&nbsp;>&nbsp;<a href='"&install&"guest.asp'>留言反馈</a>&nbsp;>&nbsp;"
Else
guest_here="<a href='"&install&"index.html'>首页</a>&nbsp;>&nbsp;<a href='"&install&"guest.asp'>留言反馈</a>&nbsp;>&nbsp;"
End if
End Function

'留言页路径
Function search_here()
If html=0 Then
guest_here="<a href='"&install&"index.asp'>首页</a>&nbsp;>&nbsp;<a href='"&install&"search.asp'>搜索</a>&nbsp;>&nbsp;"
Else
guest_here="<a href='"&install&"index.html'>首页</a>&nbsp;>&nbsp;<a href='"&install&"search.asp'>搜索</a>&nbsp;>&nbsp;"
End if
End function

'分类页HERE路径
Function here(cid)
Set ct=new category
rs=ct.getcategoryinfo(cid)
If html=0 then
If ct.rs("pcid")<>0 Then
here=here(ct.rs("pcid"))&"<a href='"&install&"list.asp?classid="&cid&"'>"&ct.rs("cname")&"</a>&nbsp;>&nbsp;"
Else
here="<a href='"&install&"index.asp'>首页</a>&nbsp;>&nbsp;<a href='"&install&"list.asp?classid="&ct.rs("cid")&"'>"&ct.rs("cname")&"</a>&nbsp;>&nbsp;"
End If
Elseif html=1 Then
If ct.rs("pcid")<>0 Then
here=here(ct.rs("pcid"))&"<a href='"&ct.rs("curl")&"index.html'>"&ct.rs("cname")&"</a>&nbsp;>&nbsp;"
Else
here="<a href='"&install&"index.html'>首页</a>&nbsp;>&nbsp;<a href='"&ct.rs("curl")&"index.html'>"&ct.rs("cname")&"</a>&nbsp;>&nbsp;"
End If
ElseIf html=2 Then
If ct.rs("pcid")<>0 Then
here=here(ct.rs("pcid"))&"<a href='"&ct.rs("curl")&"'>"&ct.rs("cname")&"</a>&nbsp;>&nbsp;"
Else
here="<a href='"&install&"index.html'>首页</a>&nbsp;>&nbsp;<a href='"&ct.rs("curl")&"'>"&ct.rs("cname")&"</a>&nbsp;>&nbsp;"
End If
End if
End Function

'arclist内容路径函数
Function cont_arclist(id)
Set ns=new news
rs=ns.getnewsinfo(id)
If html=0 Then
cont_arclist=install&"view.asp?newsid="&id
ElseIf html=1 then
cont_arclist=ns.rs("curl")&ns.rs("pinyin")&".html"
ElseIf html=2 Then
cont_arclist=install&ns.rs("pinyin")&".html"
End if
End Function

'arclist分类路径函数
Function cont_class(url)
If html=1 Then
cont_class=url&"index.html"
else
cont_class=url
End if
End Function

'读取模版
Function read_temp(url)
set stm=Server.CreateObject("adodb.stream") 
stm.Type=1 'adTypeBinary，按二进制数据读入 
stm.Mode=3 'adModeReadWrite ,这里只能用3用其他会出错 
stm.Open 
stm.LoadFromFile url
stm.Position=0 '把指针移回起点 
stm.Type=2 '文本数据 
stm.Charset=q_Charset
read_temp = re_index(stm.ReadText)
stm.Close 
set stm=nothing 
End Function

'频道分页
function pages(size)
set ns=new news
rs=ns.pagelist(classid)
if ns.rs.eof then
pages="首页&nbsp;上一页&nbsp下一页&nbsp;尾页"
else
ns.rs.pagesize=size
if html=1 then
page=html_page
else
page=clng(request("page"))
end if
if page<1 then page=1
if page>ns.rs.pagecount then page=ns.rs.pagecount
ns.rs.absolutepage=page
If html=0 Then
if page=1 then
pages="首页&nbsp;上一页&nbsp;第"&page&"页&nbsp;<a href='list.asp?classid="&classid&"&page="&page+1&"'>下一页</a>&nbsp;<a href='list.asp?classid="&classid&"&page="&ns.rs.pagecount&"'>尾页</a>"
elseif page=ns.rs.pagecount then
pages="<a href='list.asp?classid="&classid&"&page=1'>首页</a>&nbsp;<a href='list.asp?classid="&classid&"&page="&page-1&"'>上一页</a>&nbsp;第"&page&"页&nbsp;下一页&nbsp;尾页"
else pages="<a href='list.asp?classid="&classid&"&page=1'>首页</a>&nbsp;<a href='list.asp?classid="&classid&"&page="&page-1&"'>上一页</a>&nbsp;第"&page&"页&nbsp;<a href='list.asp?classid="&classid&"&page="&page+1&"'>下一页</a>&nbsp;<a href='list.asp?classid="&classid&"&page="&ns.rs.pagecount&"'>尾页</a>"
end If
ElseIf html=1 Then
if page=1 then
pages="首页&nbsp;上一页&nbsp;第"&page&"页&nbsp;<a href='"&ns.rs("curl")&"list-"&page+1&".html'>下一页</a>&nbsp;<a href='"&ns.rs("curl")&"list-"&ns.rs.pagecount&".html'>尾页</a>"
elseif page=ns.rs.pagecount then
pages="<a href='"&ns.rs("curl")&"index.html'>首页</a>&nbsp;<a href='"&ns.rs("curl")&"list-"&page-1&".html'>上一页</a>&nbsp;第"&page&"页&nbsp;下一页&nbsp;尾页"
Else
pages="<a href='"&ns.rs("curl")&"index.html'>首页</a>&nbsp;<a href='"&ns.rs("curl")&"list-"&page-1&".html'>上一页</a>&nbsp;第"&page&"页&nbsp;<a href='"&ns.rs("curl")&"list-"&page+1&".html'>下一页</a>&nbsp;<a href='"&ns.rs("curl")&"list-"&ns.rs.pagecount&".html'>尾页</a>"
end If
ElseIf html=2 Then
if page=1 then
pages="首页&nbsp;上一页&nbsp;第"&page&"页&nbsp;<a href='"&cutc(ns.rs("curl"))&"-"&page+1&".html'>下一页</a>&nbsp;<a href='"&cutc(ns.rs("curl"))&"-"&ns.rs.pagecount&".html'>尾页</a>"
elseif page=ns.rs.pagecount then
pages="<a href='"&cutc(ns.rs("curl"))&".html'>首页</a>&nbsp;<a href='"&cutc(ns.rs("curl"))&"-"&page-1&".html'>上一页</a>&nbsp;第"&page&"页&nbsp;下一页&nbsp;尾页"
else pages="<a href='"&cutc(ns.rs("curl"))&".html'>首页</a>&nbsp;<a href='"&cutc(ns.rs("curl"))&"-"&page-1&".html'>上一页</a>&nbsp;第"&page&"页&nbsp;<a href='"&cutc(ns.rs("curl"))&"-"&page+1&".html'>下一页</a>&nbsp;<a href='"&cutc(ns.rs("curl"))&"-"&ns.rs.pagecount&".html'>尾页</a>"
end If
End if
end if
end Function

'***************************************************************************
'后台路径相关操作
'***************************************************************************
'arclist分类路径函数
Function class_arclist(cid)
set ct1=new category
rs=ct1.getcategoryinfo(cid)
If html=0 Then
class_arclist=install&"list.asp?classid="&cid
elseif html=1 then
If ct1.rs("pcid")=0 Then
class_arclist=install&class_html_url(cid)&"/"
Else
class_arclist=class_help(ct1.rs("pcid"))&class_html_url(cid)&"/"
End If
ElseIf html=2 Then
class_arclist=install&class_html_url(cid)&".html"
End if
End Function

'分类路径辅助
Function class_help(cid)
set ct2=new category
rs=ct2.getcategoryinfo(cid)
If ct2.rs("pcid")=0 Then
class_help=install&class_html_url(ct2.rs("cid"))&"/"
Else
class_help=install&class_help(ct2.rs("pcid"))&class_html_url(cid)&"/"
End If
End Function

'分类替换
Function class_html_url(cid)
Set ct3=new category
rs=ct3.getcategoryinfo(cid)
kk=class_url
kk=Replace(kk,"{pinyin}",pinyin(ct3.rs("cname")))
kk=Replace(kk,"{pin_yin}",pinyin_line(ct3.rs("cname")))
kk=Replace(kk,"{py}",py(ct3.rs("cname")))
kk=Replace(kk,"{p_y}",py_line(ct3.rs("cname")))
kk=Replace(kk,"{cid}",ct3.rs("cid"))
class_html_url=kk
End Function

'内容路径替换
Function cont_html_url(id)
Set ns3=new news
rs=ns3.getnewsinfo(id)
kk=content_url
kk=Replace(kk,"{pinyin}",pinyin(ns3.rs("ntitle")))
kk=Replace(kk,"{pin_yin}",pinyin_line(ns3.rs("ntitle")))
kk=Replace(kk,"{py}",py(ns3.rs("ntitle")))
kk=Replace(kk,"{p_y}",py_line(ns3.rs("ntitle")))
kk=Replace(kk,"{id}",ns3.rs("newsid"))
cont_html_url=kk
End Function

'路径需要
function pinyin(l1)
    dim I1,I2,l2,l3,l4,i,rs,temp,strsql
    'on error resume next
    set I2=server.createobject("adodb.connection")
    I2.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&server.mappath(install&"inc/pinyin.qcms")
    if err.number<>0 then response.Write("error")
    l4=true
    for i=1 to len(l1)
        l3=l4
        l2=mid(l1,i,1)
        if len(trim(l2))=1 then'长度为1
            StrSql="select top 1 pinyin,hanzi from Hzpy where hanzi ='"&l2&"' and PutOut=true"
            set rs=server.CreateObject("adodb.recordset")
            rs.open Strsql,I2,1,2
                if not rs.eof and not rs.bof then'中文
                    l2=rs(0)
                      If len(l2)>0 then
                           l2=UCase(Left(l2,1)) & right(l2,len(l2)-1)
                      End If  
                    l4=true'当前为中文，即true
                else '没有找到就自动添加，PutOut设置为False,并用"_"替换
                    rs.addnew
                       rs("hanzi")=l2
                    rs.update
                    l2="_"            
                    l4=false'当前为英文，false
                end if
                rs.close
            set rs=nothing
        else
            l2=""'换行替换为空格
        end if
        if l3=l4 then' l2=l2&" "
            I1=I1&l2
        else
            I1=I1&""&l2
        end if
    next
    I2.close
    set I2=nothing
    pinyin=trim(I1)
end Function

'全拼加下划线
function pinyin_line(l1)
    dim I1,I2,l2,l3,l4,i,rs,temp,strsql
    'on error resume next
    set I2=server.createobject("adodb.connection")
    I2.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&server.mappath(install&"inc/pinyin.qcms")
    if err.number<>0 then response.Write("error")
    l4=true
    for i=1 to len(l1)
        l3=l4
        l2=mid(l1,i,1)
        if len(trim(l2))=1 then'长度为1
            StrSql="select top 1 pinyin,hanzi from Hzpy where hanzi ='"&l2&"' and PutOut=true"
            set rs=server.CreateObject("adodb.recordset")
            rs.open Strsql,I2,1,2
                if not rs.eof and not rs.bof then'中文
                    l2=rs(0)
                      If len(l2)>0 then
                           l2=UCase(Left(l2,1)) & right(l2,len(l2)-1)&"_"
                      End If  
                    l4=true'当前为中文，即true
                else '没有找到就自动添加，PutOut设置为False,并用"_"替换
                    rs.addnew
                       rs("hanzi")=l2
                    rs.update
                    l2="_"            
                    l4=false'当前为英文，false
                end if
                rs.close
            set rs=nothing
        else
            l2=""'换行替换为空格
        end if
        if l3=l4 then' l2=l2&" "
            I1=I1&l2
        else
            I1=I1&""&l2
        end if
    next
    I2.close
    set I2=nothing
    pinyin_line=cutp(trim(I1))
end Function

'拼音首字母
function py(l1)
    dim I1,I2,l2,l3,l4,i,rs,temp,strsql
    'on error resume next
    set I2=server.createobject("adodb.connection")
    I2.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&server.mappath(install&"inc/pinyin.qcms")
    if err.number<>0 then response.Write("error")
    l4=true
    for i=1 to len(l1)
        l3=l4
        l2=mid(l1,i,1)
        if len(trim(l2))=1 then'长度为1
            StrSql="select top 1 pinyin,hanzi from Hzpy where hanzi ='"&l2&"' and PutOut=true"
            set rs=server.CreateObject("adodb.recordset")
            rs.open Strsql,I2,1,2
                if not rs.eof and not rs.bof then'中文
                    l2=rs(0)
                      If len(l2)>0 then
                           l2=Left(l2,1)
                      End If  
                    l4=true'当前为中文，即true
                else '没有找到就自动添加，PutOut设置为False,并用"_"替换
                    rs.addnew
                       rs("hanzi")=l2
                    rs.update
                    l2="_"            
                    l4=false'当前为英文，false
                end if
                rs.close
            set rs=nothing
        else
            l2=""'换行替换为空格
        end if
        if l3=l4 then' l2=l2&" "
            I1=I1&l2
        else
            I1=I1&""&l2
        end if
    next
    I2.close
    set I2=nothing
    py=trim(I1)
end Function

'拼音首字母加下划线
function py_line(l1)
    dim I1,I2,l2,l3,l4,i,rs,temp,strsql
    'on error resume next
    set I2=server.createobject("adodb.connection")
    I2.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&server.mappath(install&"inc/pinyin.qcms")
    if err.number<>0 then response.Write("error")
    l4=true
    for i=1 to len(l1)
        l3=l4
        l2=mid(l1,i,1)
        if len(trim(l2))=1 then'长度为1
            StrSql="select top 1 pinyin,hanzi from Hzpy where hanzi ='"&l2&"' and PutOut=true"
            set rs=server.CreateObject("adodb.recordset")
            rs.open Strsql,I2,1,2
                if not rs.eof and not rs.bof then'中文
                    l2=rs(0)
                      If len(l2)>0 then
                           l2=Left(l2,1)&"_"
                      End If  
                    l4=true'当前为中文，即true
                else '没有找到就自动添加，PutOut设置为False,并用"_"替换
                    rs.addnew
                       rs("hanzi")=l2
                    rs.update
                    l2="_"            
                    l4=false'当前为英文，false
                end if
                rs.close
            set rs=nothing
        else
            l2=""'换行替换为空格
        end if
        if l3=l4 then' l2=l2&" "
            I1=I1&l2
        else
            I1=I1&""&l2
        end if
    next
    I2.close
    set I2=nothing
    py_line=cutp(trim(I1))
end Function

'切源码
function cutp(n)
ss=len(n)
xx=mid(n,1,ss-1)
cutp=xx
end Function

'切源码
function cutc(n)
ss=len(n)
xx=mid(n,1,ss-5)
cutc=xx
end Function

'过滤HTML并替换
function nohtml(str)  
dim re  
Set re=new RegExp  
re.IgnoreCase =true  
re.Global=True  
str = replace(str,chr(13),"{br}") 
re.Pattern="<.*?>"  
str=re.replace(str,"")  
str = replace(str,"{br}","<br>") 
nohtml=str  
set re=nothing  
end function 

'********************************************** 
'vbs Cache类
' 属性valid，是否可用，取值前判断 
' 属性name，cache名，新建对象后赋值 
' 方法add(值,到期时间)，设置cache内容 
' 属性value，返回cache内容 
' 属性blempty，是否未设置值 
' 方法makeEmpty，释放内存，测试用 
' 方法equal(变量1)，判断cache值是否和变量1相同 
' 方法expires(time)，修改过期时间为time 
' 木鸟写的缓存类
'********************************************** 

class Cache 
private obj 'cache内容 
private expireTime '过期时间 
private expireTimeName '过期时间application名 
private cacheName 'cache内容application名 
private path 'uri 

private sub class_initialize() 

  path=request.servervariables("url") 

  path=left(path,instrRev(path,"/")) 
end sub 

private sub class_terminate() 
end sub 

public property get blEmpty 

  '是否为空 

  if isempty(obj) then 

      blEmpty=true 

  else 

      blEmpty=false 

  end if 
end property 

public property get valid 

  '是否可用(过期) 

  if isempty(obj) or not isDate(expireTime) then 

      valid=false 

  elseif CDate(expireTime)<now then 

      valid=false 

  else 

      valid=true 

  end if 
end property 

public property let name(str) 

  '设置cache名 

  cacheName=str & path 

  obj=application(cacheName) 

  expireTimeName=str & "expires" & path 

  expireTime=application(expireTimeName) 
end property 

public property let expires(tm) 

  '重设置过期时间 

  expireTime=tm 

  application.lock 

  application(expireTimeName)=expireTime 

  application.unlock 
end property 

public sub add(var,expire) 

  '赋值 

  if isempty(var) or not isDate(expire) then 

      exit sub 

  end if 

  obj=var 

  expireTime=expire 

  application.lock 

  application(cacheName)=obj 

  application(expireTimeName)=expireTime 

  application.unlock 
end sub 

public property get value 

  '取值 

  if isempty(obj) or not isDate(expireTime) then 

      value=null 

  elseif CDate(expireTime)<now then 

      value=null 

  else 

      value=obj 

  end if 
end property 

public sub makeEmpty() 

  '释放application 

  application.lock 

  application(cacheName)=empty 

  application(expireTimeName)=empty 

  application.unlock 

  obj=empty 

  expireTime=empty 
end sub 

public function equal(var2) 

  '比较 

  if typename(obj)<>typename(var2) then 

      equal=false 

  elseif typename(obj)="Object" then 

      if obj is var2 then 

          equal=true 
    
  else 
    
      equal=false 
    
  end if 

  elseif typename(obj)="Variant()" then 

      if join(obj,"^")=join(var2,"^") then 

          equal=true 
    
  else 
    
      equal=false 
    
  end if 

  else

      if obj=var2 then 

          equal=true 
    
  else 
    
      equal=false 
    
  end if 

  end if 
end function 

end class 
%> 