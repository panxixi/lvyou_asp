<!--#include file="do.asp"-->
<%
'获取首页模版路径
Function getAdoStreamText(byval templatePath)
    Set stm = Server.CreateObject("adodb.stream")
    stm.Type = 1 'adTypeBinary，按二进制数据读入
    stm.Mode = 3 'adModeReadWrite ,这里只能用3用其他会出错
    stm.Open
    
    stm.LoadFromFile server.mappath(install&templatePath)
    stm.Position = 0 '把指针移回起点
    stm.Type = 2 '文本数据
    stm.Charset = "utf-8"
    getAdoStreamText = stm.ReadText
    stm.Close
    Set stm = Nothing
end Function


Function index_temp()
    Set cn = New config
    rs = cn.getconfig()
    index_temp = re_index(getAdoStreamText(cn.rs("itemp")))
    cn.rs.Close()
    Set cn = Nothing
End Function

'获取分类模版

Function list_temp(id)
    If Not IsNumeric(id) Then response.Redirect(install&"fail.asp")
    Set ct = New category
    rs = ct.getcategoryinfo(id)
    If ct.rs.EOF Then
        response.Redirect(install&"fail.asp")
    Else
        list_temp = re_list(getAdoStreamText(ct.rs("ctemp")))
    End If
    ct.rs.Close()
    Set ct = Nothing
End Function

'获取内容模版

Function view_temp(id)
    If Not IsNumeric(id) Then response.Redirect(install&"fail.asp")
    Set ns = New news
    rs = ns.getnewsinfo(id)

    If ns.rs.EOF Then
        response.Redirect(install&"fail.asp")
    Else
        view_temp = re_view(getAdoStreamText(ns.rs("ntemp")))
    End If
    ns.rs.Close()
    Set ns = Nothing
End Function

function toBackurl()
	if request.servervariables("http_referer")<>"" then 
		toUrl(request.servervariables("http_referer"))
	end if 
End Function
'获取留言模版

Function guest_temp()
    guest_temp = re_index(getAdoStreamText("templist/guest.html"))
End Function

Function huodongList_temp()
    huodongList_temp = re_index(getAdoStreamText("templist/huodonglist.html"))
End Function


Function huoDong_temp()
    huoDong_temp = re_HuoDong(getAdoStreamText("templist/huoDongView.html"))
End Function

Function huoDong_temp2()
    huoDong_temp2 = re_baoming2(getAdoStreamText("templist/huoDongView.html"))
End Function

Function Baoming_temp(ByVal tempStr)
    Baoming_temp = re_HuoDong(getAdoStreamText("templist/Baoming"&tempStr&".html"))
End Function

'获取会员注册模板
Function userReg_temp()
	Set cn = New config
    cn.getconfig()
    userTemplates = split(cn.rs("usertemplates"),"|")
    userReg_temp = re_index(getAdoStreamText(userTemplates(0)))
    cn.rs.close
    set cn.rs = nothing
End Function

Function userQQReg_temp()
    userQQReg_temp = re_index(getAdoStreamText("templist/QQReg.html"))
End Function

'获取会员登录模板
Function userLogin_temp()
	Set cn = New config
    cn.getconfig()
    userTemplates = split(cn.rs("usertemplates"),"|")
    userLogin_temp = re_index(getAdoStreamText(userTemplates(1)))
    cn.rs.close
    set cn.rs = nothing
End Function

'获取会员首页模板
Function userIndex_temp()
	Set cn = New config
    cn.getconfig()
    userTemplates = split(cn.rs("usertemplates"),"|")
    userIndex_temp = re_UserIndex(getAdoStreamText(userTemplates(2)))
    cn.rs.close
    set cn.rs = nothing
End Function

'获取会员首页模板
Function userIndex_str(byval Tpath)
    userIndex_str = re_UserIndex2(getAdoStreamText(Tpath))
End Function

Function userIndex_Baomingstr(byval Tpath)
    userIndex_Baomingstr = re_baoMingStr(getAdoStreamText(Tpath))
End Function

Function userTp_str(byval Tpath)
	userTp_str = re_index(getAdoStreamText(Tpath))
end Function

Function userBaoming_str(byval Tpath)
	userBaoming_str = re_baoMing_str(getAdoStreamText(Tpath))
End Function

'获取会员报名模板
Function userBaoming_temp()
	Set cn = New config
    cn.getconfig()
    userTemplates = split(cn.rs("usertemplates"),"|")
    userBaoming_temp = re_UserIndex(getAdoStreamText(userTemplates(3)))
    cn.rs.close
    set cn.rs = nothing
End Function

'获取会员报名成功模板
Function userBaomingSuccess_temp()
	Set cn = New config
    cn.getconfig()
    userTemplates = split(cn.rs("usertemplates"),"|")
    userBaoming_temp = re_baomingsuccess(getAdoStreamText(userTemplates(4)))
    cn.rs.close
    set cn.rs = nothing
End Function

'获取自定义页面模版

Function diy_temp(id)
    If Not IsNumeric(id) Then response.Redirect(install&"fail.asp")
    Set cn = New config
    cn.getdiy(id)
    If cn.rs.EOF Then
        response.Redirect(install&"fail.asp")
    Else
        diy_nr = cn.rs("diy")
        diy_temp = re_index(diy_nr)
    End If
    cn.rs.Close()
    Set cn = Nothing
End Function

'获取JS站外调用模版

Function js_temp(id)
    If Not IsNumeric(id) Then response.Redirect(install&"fail.asp")
    Set cn = New config
    cn.getjss(id)
    If cn.rs.EOF Then
        response.Redirect(install&"fail.asp")
    Else
        js_nr = cn.rs("js_code")
        js_temp = re_index(js_nr, 6, id)
    End If
    cn.rs.Close()
    Set cn = Nothing
End Function

Function re_index(mb)'首页替换

    yuke = baohan(mb)
	
    yuke = otherLabel(yuke)
        'die yuke
    yuke = userViewlogin(yuke)
    yuke = reqlabel(yuke)
  
    yuke = qbq(yuke)
    yuke = th_index(yuke)

    yuke = memu(yuke)

    yuke = smemu(yuke)
    yuke = arclist(yuke)
 
    yuke = loop_sql(yuke)
         
    yuke = FriendLinklist(yuke)
    yuke = singleContent(yuke)

    yuke = slide(yuke)

    yuke = userList(yuke)

    yuke = huoDongList(yuke)
    yuke = huoDongZJList(yuke)
    yuke = iFThenElse(yuke)
    yuke = getLen(yuke)
    re_index = yuke
End Function

function re_baoMing_str(mb)
    yuke = baohan(mb)
    yuke = otherLabel(yuke)
    yuke = reqlabel(yuke)
    yuke = qbq(yuke)
    yuke = huoDongView(yuke)
    yuke = memu(yuke)
    yuke = smemu(yuke)
    yuke = arclist(yuke)
    yuke = loop_sql(yuke)
    yuke = FriendLinklist(yuke)
    yuke = singleContent(yuke)
    yuke = slide(yuke)
    yuke = userList(yuke)
    yuke = userView(yuke)
    yuke = baomingList(yuke)
    yuke = baomingShowList(yuke)
    yuke = baomingView(yuke)
    yuke = huoDongList(yuke)
    yuke = iFThenElse(yuke)
    yuke = getLen(yuke)
	'yuke = baomingSingleView(yuke)
	re_baoMing_str = yuke
end function

function includeRe(mb)
    yuke = qbq(mb)
    yuke = otherLabel(yuke)
    yuke = reqlabel(yuke)
    yuke = th_index(yuke)
    yuke = memu(yuke)
    yuke = smemu(yuke)
    yuke = arclist(yuke)
    yuke = loop_sql(yuke)
    yuke = FriendLinklist(yuke)
    yuke = singleContent(yuke)
    yuke = slide(yuke)
    yuke = userList(yuke)
    yuke = huoDongList(yuke)
    yuke = iFThenElse(yuke)
    yuke = getLen(yuke)
    includeRe = yuke
End Function

Function re_list(mb)'分类替换
    yuke = baohan(mb)
    yuke = otherLabel(yuke)

    yuke = reqlabel(yuke)
    yuke = qbq(yuke)
    yuke = th_list(yuke)
    yuke = memu(yuke)

    yuke = classlist(yuke)

    yuke = smemu(yuke)

    yuke = arclist(yuke)

    yuke = pth(yuke)
    yuke = loop_sql(yuke)
    yuke = FriendLinklist(yuke)

    yuke = singleContent(yuke)
    yuke = slide(yuke)
    yuke = userList(yuke)
    yuke = huoDongList(yuke)
    yuke = iFThenElse(yuke)
    yuke = getLen(yuke)
    re_list = yuke
End Function

Function re_view(mb)'内容替换
    yuke = baohan(mb)
    yuke = otherLabel(yuke)

    yuke = reqlabel(yuke)
    yuke = qbq(yuke)
    yuke = th_view(yuke)
    yuke = memu(yuke)

    yuke = smemu(yuke)

    yuke = arclist(yuke)

    yuke = content(yuke)
    yuke = like_list(yuke)
    yuke = loop_sql(yuke)
    yuke = FriendLinklist(yuke)
    yuke = singleContent(yuke)
    yuke = slide(yuke)

    yuke = userList(yuke)
    yuke = huoDongList(yuke)
    yuke = iFThenElse(yuke)
	yuke = getLen(yuke)
    re_view = yuke
End Function


Function re_HuoDong(mb)'内容替换
    yuke = baohan(mb)
    yuke = otherLabel(yuke)
    yuke = reqlabel(yuke)
    yuke = qbq(yuke)
    yuke = huoDongView(yuke)
    yuke = memu(yuke)
    yuke = smemu(yuke)
    yuke = arclist(yuke)
    yuke = loop_sql(yuke)
    yuke = FriendLinklist(yuke)
    yuke = singleContent(yuke)
    yuke = slide(yuke)
    yuke = userList(yuke)
    yuke = userView(yuke)
    yuke = baomingList(yuke)
    yuke = baomingShowList(yuke)
    'yuke = baomingView2(yuke)
    yuke = baomingView(yuke)
    'yuke = baomingViewSuccess(yuke)
    yuke = huoDongList(yuke)
    yuke = iFThenElse(yuke)
    yuke = getLen(yuke)
    yuke = replaceLocation(yuke)
    re_HuoDong = yuke
End Function

Function re_Baoming2(mb)'内容替换
    yuke = baohan(mb)
    yuke = otherLabel(yuke)
    yuke = userViewlogin(yuke)
    yuke = reqlabel(yuke)
    yuke = qbq(yuke)
    yuke = huoDongView(yuke)
    yuke = memu(yuke)
    yuke = smemu(yuke)
    yuke = arclist(yuke)
    yuke = loop_sql(yuke)
    yuke = FriendLinklist(yuke)
    yuke = singleContent(yuke)
    yuke = slide(yuke)
    yuke = userList(yuke)
    'yuke = userView(yuke)
    yuke = baomingList(yuke)
    yuke = baomingShowList(yuke)
    'yuke = baomingView2(yuke)
    yuke = baomingView(yuke)
    'yuke = baomingViewSuccess(yuke)
    yuke = huoDongList(yuke)
    yuke = iFThenElse(yuke)
    yuke = getLen(yuke)
    yuke = replaceLocation(yuke)
    re_Baoming2 = yuke
End Function

Function re_BaoMing(mb)'内容替换
    yuke = baohan(mb)
    yuke = otherLabel(yuke)
    yuke = reqlabel(yuke)
    yuke = qbq(yuke)
    yuke = huoDongView(yuke)
    yuke = memu(yuke)
    yuke = smemu(yuke)
    yuke = arclist(yuke)
    yuke = loop_sql(yuke)
    yuke = FriendLinklist(yuke)
    yuke = singleContent(yuke)
    yuke = slide(yuke)
    yuke = userList(yuke)
    yuke = userView(yuke)
    yuke = baomingList(yuke)
    yuke = baomingShowList(yuke)
    yuke = huoDongList(yuke)
    yuke = iFThenElse(yuke)
    yuke = getLen(yuke)
    re_BaoMing = yuke
End Function


Function re_UserIndex(mb)
    yuke = baohan(mb)
    yuke = otherLabel(yuke)
    yuke = reqlabel(yuke)
    yuke = qbq(yuke)
    yuke = th_index(yuke)
    yuke = memu(yuke)
    yuke = smemu(yuke)
    yuke = arclist(yuke)
    yuke = loop_sql(yuke)
    yuke = FriendLinklist(yuke)
    yuke = singleContent(yuke)
    yuke = slide(yuke)
    yuke = userList(yuke)
    yuke = userView(yuke)
    yuke = baomingList(yuke)
    yuke = baomingShowList(yuke)
    yuke = huoDongList(yuke)
    yuke = iFThenElse(yuke)
    yuke = getLen(yuke)
    re_UserIndex = yuke
end Function

Function re_UserIndex2(mb)
    yuke = baohan(mb)
    yuke = otherLabel(yuke)
    yuke = reqlabel(yuke)
    yuke = qbq(yuke)
    yuke = th_index(yuke)
    yuke = memu(yuke)
    yuke = smemu(yuke)
    yuke = arclist(yuke)
    yuke = loop_sql(yuke)
    yuke = FriendLinklist(yuke)
    yuke = singleContent(yuke)
    yuke = slide(yuke)
    yuke = userList(yuke)
    yuke = userView(yuke)
    'yuke = baomingView(yuke)
    yuke = baomingList(yuke)
    yuke = baomingShowList(yuke)
    yuke = huoDongList(yuke)
    yuke = iFThenElse(yuke)
    yuke = getLen(yuke)
    re_UserIndex2 = yuke
end Function

Function re_baoMingStr(mb)
    yuke = baohan(mb)
    yuke = otherLabel(yuke)
    yuke = reqlabel(yuke)
    yuke = qbq(yuke)
    yuke = th_index(yuke)
    yuke = memu(yuke)
    yuke = smemu(yuke)
    yuke = arclist(yuke)
    yuke = loop_sql(yuke)
    yuke = FriendLinklist(yuke)
    yuke = singleContent(yuke)
    yuke = slide(yuke)
    yuke = userList(yuke)
    yuke = userView(yuke)
    yuke = baomingView(yuke)
    yuke = baomingList(yuke)
    yuke = baomingShowList(yuke)
    yuke = huoDongList(yuke)
    yuke = iFThenElse(yuke)
    yuke = getLen(yuke)
    re_UserIndex2 = yuke
end Function


function re_baomingsuccess(mb)

end function
'匹配单标签




Function qbq(Str)
    Set reg = New regexp
    reg.IgnoreCase = False
    reg.Global = True
    reg.Pattern = "{旅游:自定义标签:([0-9]+)?}"
    Set aa = reg.Execute(Str)
    For Each match in aa
        dbq = match.submatches(0)
        ybq = match
        Set cn = New config
        cn.getbiaoqian(dbq)
        If Not cn.rs.EOF Then
            Str = Replace(Str, ybq, cn.rs("bq"))
        End If
    Next
    qbq = Str
    Set reg = Nothing
    Set aa = Nothing
    Set cn = Nothing
End Function


Function reqlabel(Str)
    Set reg = New regexp
    reg.IgnoreCase = False
    reg.Global = True
    reg.Pattern = "{旅游:地址栏请求标签:([\s\S.]*?)}"
    Set aa = reg.Execute(Str)
    For Each match in aa
        dbq = match.submatches(0)
        ybq = match
        aaa = Request.QueryString(dbq)
        Str = Replace(Str, ybq, aaa)
    Next
    reqlabel = Str
    Set reg = Nothing
    Set aa = Nothing
End Function



'匹配调用单条内容标签

Function singleContent(Str)
    Set reg = New regexp
    reg.IgnoreCase = False
    reg.Global = True
    reg.Pattern = "{旅游:获取单条内容:([\s\S.]*?)}"
    Set aa = reg.Execute(Str)
    For Each match in aa
        dbq = match.submatches(0)
        ybq = match
        Set cn = New config
        cn.getbiaoqianbyName(dbq)
        If Not cn.rs.EOF Then
            Str = Replace(Str, ybq, cn.rs("bq"))
        End If
    Next
    singleContent = Str
    Set reg = Nothing
    Set aa = Nothing
    Set cn = Nothing
End Function



'包含标签替换

Function baohan(qzy)
    Set reg = New regexp
    reg.IgnoreCase = False
    reg.Global = True
    reg.Pattern = "{旅游_包含:(.*?)}"
    Set aa = reg.Execute(qzy)
    For Each match in aa
        url = match.submatches(0)
        Set stm = Server.CreateObject("adodb.stream")
        stm.Type = 1 'adTypeBinary，按二进制数据读入
        stm.Mode = 3 'adModeReadWrite ,这里只能用3用其他会出错
        stm.Open
        stm.LoadFromFile server.mappath(install&"templist/"&match.submatches(0))
        stm.Position = 0 '把指针移回起点
        stm.Type = 2 '文本数据
        stm.Charset = "utf-8"
        cont = stm.ReadText
        stm.Close
        Set stm = Nothing
        kk = Replace(qzy, match, cont)
        qzy = Replace(kk, match, cont)
    Next
    baohan = qzy
    Set aa = Nothing
    Set reg = Nothing
End Function

'首页单标签替换

Function th_index(qzy)
    Set cn = New config
    cn.getconfig()
    aa = Array("{旅游:网站名称}", "{旅游:版权信息}", "{旅游:安装目录}", "{旅游:留言表单}", "{旅游:留言本}", "&#39;", "&#34;", "{旅游:主页}")
    bb = Array(cn.rs("webname"), cn.rs("copyright"), install, guest, lyb(lyts), Chr(39), Chr(34), "<a href='"&install&"/index.asp'>主页</a> >")
    For i = 0 To UBound(bb)
        zz = Replace(qzy, aa(i), bb(i))
        qzy = Replace(zz, aa(i), bb(i))
    Next
    th_index = qzy
    cn.rs.Close()
    Set cn = Nothing
End Function

Function userIsLogin()
	userName = vbsUnEscape(Replace(Request.Cookies(CookieName)("UserName"),"'","''"))
	password = Replace(Request.Cookies(CookieName)("password"),"'","''")
	if userName = "" or password = "" then userIsLogin = 0 : Exit Function
	if not checkLoginBool(userName,password) then
		userIsLogin = 0
	else
		userIsLogin = 1
	end if
End Function

Function isBindQQ()
	dim flag,userName
	userName = vbsUnEscape(Replace(Request.Cookies(CookieName)("UserName"),"'","''"))
	set u=new Users
	u.getuser(userName)
	if not u.rs.eof then 
		if u.rs("QQHASH") <> "" or u.rs("QQHASH") = session("qqId") then
			flag = true
		else
			flag = false
		end if
	end if
	isBindQQ = flag
	u.rs.close
End Function

Function otherLabel(qzy)
	aa = Array("{旅游:检测登陆}")
    bb = Array(userIsLogin)
    For i = 0 To UBound(bb)
        'zz = Replace(qzy, aa(i), bb(i))
        qzy = Replace(qzy, aa(i), bb(i))
    Next
    'die qzy
    otherLabel = qzy
End Function

'返回长度
Function getLen(qzy)
	Set reg = New regexp
    reg.IgnoreCase = False
    reg.Global = True
    reg.Pattern = "{长度:([\s\S.]*?)}"
    Set aa = reg.Execute(qzy)
    For Each match in aa
        lenNum = clng(len(match.submatches(0)))
        kk = Replace(qzy, match, cont)
        qzy = Replace(kk, match, cont)
    Next
    getLen = qzy
End Function

Function replaceLocation(yuke)
str = replace(yuke,"{旅游:获得乘车地点}",GetCarLocationView())
replaceLocation = str
End Function

'列表单标签替换

Function th_list(qzy)
    Set cn = New config
    cn.getconfig()
    Set ct = New category
    rs = ct.getcategoryinfo(classid)
    aa = Array("{旅游:网站名称}", "{旅游:版权信息}", "{旅游:安装目录}", "{旅游:留言表单}", "{旅游:留言本}", "&#39;", "&#34;", "{旅游:主页}")
    bb = Array(cn.rs("webname"), cn.rs("copyright"), install, guest, lyb(lyts), Chr(39), Chr(34), "<a href='"&install&"'>主页</a> > <a href='"&install&"list.asp?classid="&classid&"'>"&ct.rs("cname")&"</a> > 列表 ")
    For i = 0 To UBound(bb)
        zz = Replace(qzy, aa(i), bb(i))
        qzy = Replace(zz, aa(i), bb(i))
    Next
    th_list = qzy
    cn.rs.Close()
    ct.rs.Close()
    Set cn = Nothing
    Set ct = Nothing
End Function

'内容单标签替换

Function th_view(qzy)
    Set cn = New config
    cn.getconfig()
    Set ns = New news
    call ns.getnewsinfo(newsid)
    aa = Array("{旅游:网站名称}", "{旅游:版权信息}", "{旅游:安装目录}", "{旅游:留言表单}", "{旅游:留言本}", "&#39;", "&#34;", "{旅游:主页}")
    bb = Array(cn.rs("webname"), cn.rs("copyright"), install, guest, lyb(lyts), Chr(39), Chr(34), "<a href='"&install&"'>主页</a> > <a href='"&install&"list.asp?classid="&ns.rs("news.cid")&"'>"&ns.rs("cname")&"</a> > ")
    For i = 0 To UBound(bb)
        zz = Replace(qzy, aa(i), bb(i))
        qzy = Replace(zz, aa(i), bb(i))
    Next
    th_view = qzy
    cn.rs.Close()
    ns.rs.Close()
    Set cn = Nothing
    Set ns = Nothing
End Function

Function userView(qzy)
	set u = new Users
	u.getuser(userName)
	if not u.rs.eof then
		'#####################
		aa = Array("{旅游:用户编号}", "{旅游:用户名称}", "{旅游:用户密码}", "{旅游:用户真名}", "{旅游:用户性别}", "{旅游:用户电话}", "{旅游:身份证号}", "{旅游:加入时间}", "{旅游:当前积分}", "{旅游:论坛用户ID}", "{旅游:获取头像编辑器}")
	    bb = Array(u.rs("userId"), u.rs("userName"), u.rs("password"), u.rs("RealName"), u.rs("sex"), u.rs("phone"),u.rs("idcard"),formatdatetime(u.rs("joinTime"), 2),u.rs("jifen"),Request.cookies(CookieName)("Uid"), "无")'uc_avatar(Request.cookies(CookieName)("Uid"),"virtual","1")
	    For i = 0 To UBound(bb)
	        zz = Replace(qzy, aa(i), bb(i))
	        qzy = Replace(zz, aa(i), bb(i))
	    Next
    end if
    userView = qzy
    u.rs.Close()
    Set u = Nothing
End Function

Function userViewlogin(qzy)

	set u = new Users
	u.getuser(userName)
	password = Replace(Request.Cookies(CookieName)("password"),"'","''")
	if not u.rs.eof then
		'#####################
		if u.rs("password") = password then
			aa = Array("{旅游:用户编号}", "{旅游:用户名称}", "{旅游:用户密码}", "{旅游:用户真名}", "{旅游:用户性别}", "{旅游:用户电话}", "{旅游:身份证号}", "{旅游:加入时间}", "{旅游:当前积分}", "{旅游:论坛用户ID}", "{旅游:获取头像编辑器}")
		    bb = Array(u.rs("userId"), u.rs("userName"), u.rs("password"), u.rs("RealName"), u.rs("sex"), u.rs("phone"),u.rs("idcard"),formatdatetime(u.rs("joinTime"), 2),u.rs("jifen"),Request.cookies(CookieName)("Uid"),"无")
		    For i = 0 To UBound(bb)
		        zz = Replace(qzy, aa(i), bb(i))
		        qzy = Replace(zz, aa(i), bb(i))
		    Next
		end if
    end if
    userViewlogin = qzy
    u.rs.Close()
    Set u = Nothing
End Function


Function isBindQQBackContent(Content)
	if not isBindQQ() then
		isBindQQBackContent = replace(Content,"{旅游:提示语句}","<a href='/txapi/QQbind.asp?apiType=qq'><img src='/images/qq_bind_small.gif' /></a>")
	else
		isBindQQBackContent = replace(Content,"{旅游:提示语句}","")
	end if
End Function


Function huoDongView(qzy)
	set h = new HuoDong
	h.gethuoDong(huoDongId)
	if not h.rs.eof then
		aa = Array("{旅游:HDYM活动编号}", "{旅游:HDYM论坛链接}", "{旅游:HDYM活动名称}", "{旅游:HDYM活动人数}", "{旅游:HDYM活动剩余人数}", "{旅游:HDYM活动行程}", "{旅游:HDYM活动协议}", "{旅游:HDYM活动领队}", "{旅游:HDYM活动积分}","{旅游:HDYM活动状态}","{旅游:HDYM活动状态文字版}","{旅游:HDYM活动出发时间}","{旅游:HDYM活动返回时间}","{旅游:HDYM活动添加时间}","{旅游:HDYM活动关键字}","{旅游:HDYM活动介绍}","{旅游:HDYM活动已报人数}","{旅游:HDYM活动图片地址}")
	    bb = Array(h.rs("huoDongId"), h.rs("bbsLink"), h.rs("huoDongName"), h.rs("huoDongRenShu"), h.rs("huoDongShengyu"), cstr(h.rs("huoDongXingcheng")), cstr(h.rs("huoDongxieyi")),h.rs("huoDongLingdui"),h.rs("huoDongJifen"),huoDongZtBack(h.rs("huoDongZhuangtai")),huoDongZtBackText(h.rs("huoDongZhuangtai")),formatdatetime(h.rs("huoDongChufaTime"), 2),formatdatetime(h.rs("huoDongFanhuiTime"), 2),h.rs("huoDongAddTime"),h.rs("nkeyword"),cstr(h.rs("huoDongInfo")),h.rs("huoDongRenShu")-h.rs("huoDongShengyu"),toppic(h.rs("huoDongXingcheng"),"html","/images/cl_ad.jpg"))
	    For i = 0 To UBound(bb)
	        'zz = Replace(qzy, aa(i), bb(i))
	        qzy = Replace(qzy, aa(i), bb(i))
	    Next
    end if
    huoDongView = qzy
    h.rs.Close()
    set h.rs = nothing
    Set h = Nothing
End Function

Function baomingSingleView(qzy)
	set b = new Baoming
	call b.getUserBaomingSingle(userName,baoMingId,huoDongId)
	if not b.rs.eof then
		aa = Array("{旅游:报名编号}", "{旅游:活动名称}", "{旅游:活动编号}", "{旅游:报名用户昵称}", "{旅游:报名真实姓名}", "{旅游:报名人数}", "{旅游:报名人电话}", "{旅游:报名人身份证}","{旅游:报名状态}","{旅游:报名人应得积分}","{旅游:报名添加时间}","{旅游:报名是否使用积分}","{旅游:活动总人数}","{旅游:活动行程}","{旅游:活动协议}","{旅游:活动领队}","{旅游:活动积分}","{旅游:活动状态}","{旅游:活动状态文字版}","{旅游:活动出发时间}","{旅游:活动返回时间}","{旅游:活动添加时间}","{旅游:活动介绍}","{旅游:论坛链接}")
	    bb = Array(b.rs("baoMingId"), b.rs("baoMingHuoDong"), b.rs("baoMingHuoDongId"), b.rs("baoMingUserName"), cstr(b.rs("baoMingRealName")), cstr(b.rs("baoMingRenshu")),b.rs("baoMingPhone"),b.rs("baoMingIdcard"),b.rs("baoMingZhuangtai"),huoDongZtBackText(h.rs("huoDongZhuangtai")),b.rs("baoMingJifen"),formatdatetime(b.rs("baoMingAddTime"), 2),b.rs("baoMingUseJifen"),b.rs("huoDongrenshu"),b.rs("huoDongxingcheng"),b.rs("huoDongxieyi"),b.rs("huoDonglingdui"),b.rs("huoDongjifen"),b.rs("huoDongZhuangtai"),b.rs("huoDongChufaTime"),b.rs("huoDongFanhuiTime"),b.rs("huoDongAddtime"),b.rs("huoDongInfo"),b.rs("bbsLink"))
	    For i = 0 To UBound(bb)
	        'zz = Replace(qzy, aa(i), bb(i))
	        qzy = Replace(qzy, aa(i), bb(i))
	    Next
    end if
    baomingSingleView = qzy
    b.rs.Close()
    set b.rs = nothing
    Set b = Nothing
end Function

Function GetCarLocationView()
	set carl = new CarLocation
	call carl.allCarLocation()
	dim str
	Do While Not carl.rs.EOF
	if str<>"" then
		str = str & "<input id="""&carl.rs("location")&"_"&carl.rs("id")&""" type=""radio"" name=""location"" value="""&carl.rs("location")&"""> &nbsp;<label for="""&carl.rs("location")&"_"&carl.rs("id")&""">"&carl.rs("location")&"</label>&nbsp;&nbsp;&nbsp;"
	else
		str = str & "<input id="""&carl.rs("location")&"_"&carl.rs("id")&""" type=""radio"" name=""location"" value="""&carl.rs("location")&""" checked> &nbsp;<label for="""&carl.rs("location")&"_"&carl.rs("id")&""">"&carl.rs("location")&"</label>&nbsp;&nbsp;&nbsp;"
	end if
		carl.rs.movenext
	Loop
	GetCarLocationView = str
End Function

Function baomingView(qzy)
	set b = new Baoming
	call b.getUserbaoMingList(userName,huoDongId,"")
	if not b.rs.eof then
		aa = Array("{旅游:报名编号}", "{旅游:活动名称}", "{旅游:活动编号}", "{旅游:报名用户昵称}", "{旅游:报名真实姓名}", "{旅游:报名人数}", "{旅游:报名人电话}", "{旅游:报名人身份证}","{旅游:报名状态}","{旅游:报名人应得积分}","{旅游:报名添加时间}","{旅游:报名是否使用积分}","{旅游:活动总人数}","{旅游:活动行程}","{旅游:活动协议}","{旅游:活动领队}","{旅游:活动积分}","{旅游:活动状态}","{旅游:活动状态文字版}","{旅游:活动出发时间}","{旅游:活动返回时间}","{旅游:活动添加时间}","{旅游:活动介绍}","{旅游:论坛链接}")
	    bb = Array(b.rs("baoMingId"), b.rs("baoMingHuoDong"), b.rs("baoMingHuoDongId"), b.rs("baoMingUserName"), cstr(b.rs("baoMingRealName")), cstr(b.rs("baoMingRenshu")),b.rs("baoMingPhone"),b.rs("baoMingIdcard"),b.rs("baoMingZhuangtai"),huoDongZtBackText(b.rs("huoDongZhuangtai")),b.rs("baoMingJifen"),formatdatetime(b.rs("baoMingAddTime"), 2),b.rs("baoMingUseJifen"),b.rs("huoDongrenshu"),b.rs("huoDongxingcheng"),b.rs("huoDongxieyi"),b.rs("huoDonglingdui"),b.rs("huoDongjifen"),b.rs("huoDongZhuangtai"),b.rs("huoDongChufaTime"),b.rs("huoDongFanhuiTime"),b.rs("huoDongAddtime"),b.rs("huoDongInfo"),b.rs("bbsLink"))
	    For i = 0 To UBound(bb)
	        'zz = Replace(qzy, aa(i), bb(i))
	        qzy = Replace(qzy, aa(i), bb(i))
	    Next
    end if
    baomingView = qzy
    b.rs.Close()
    set b.rs = nothing
    Set b = Nothing
End Function

Function baomingView2(qzy)
	set b = new Baoming
	call b.getUserBaomingSingle(userName,baomingId,huoDongId)
	if not b.rs.eof then
		aa = Array("{旅游:报名编号}", "{旅游:活动名称}", "{旅游:活动编号}", "{旅游:报名用户昵称}", "{旅游:报名真实姓名}", "{旅游:报名人数}", "{旅游:报名人电话}", "{旅游:报名人身份证}","{旅游:报名状态}","{旅游:报名人应得积分}","{旅游:报名添加时间}","{旅游:报名是否使用积分}","{旅游:活动总人数}","{旅游:活动行程}","{旅游:活动协议}","{旅游:活动领队}","{旅游:活动积分}","{旅游:活动状态}","{旅游:活动状态文字版}","{旅游:活动出发时间}","{旅游:活动返回时间}","{旅游:活动添加时间}","{旅游:活动介绍}","{旅游:论坛链接}")
	    bb = Array(b.rs("baoMingId"), b.rs("baoMingHuoDong"), b.rs("baoMingHuoDongId"), b.rs("baoMingUserName"), cstr(b.rs("baoMingRealName")), cstr(b.rs("baoMingRenshu")),b.rs("baoMingPhone"),b.rs("baoMingIdcard"),b.rs("baoMingZhuangtai"),huoDongZtBackText(b.rs("huoDongZhuangtai")),b.rs("baoMingJifen"),formatdatetime(b.rs("baoMingAddTime"), 2),b.rs("baoMingUseJifen"),b.rs("huoDongrenshu"),b.rs("huoDongxingcheng"),b.rs("huoDongxieyi"),b.rs("huoDonglingdui"),b.rs("huoDongjifen"),b.rs("huoDongZhuangtai"),b.rs("huoDongChufaTime"),b.rs("huoDongFanhuiTime"),b.rs("huoDongAddtime"),b.rs("huoDongInfo"),b.rs("bbsLink"))
	    For i = 0 To UBound(bb)
	        'zz = Replace(qzy, aa(i), bb(i))
	        qzy = Replace(qzy, aa(i), bb(i))
	    Next
    end if
    baomingView2 = qzy
    b.rs.Close()
    set b.rs = nothing
    Set b = Nothing
End Function

Function baomingViewSuccess(qzy)
	set b = new Baoming
	call b.getUserbaoMingSuccess(userName,huoDongId)
	if not b.rs.eof then
		aa = Array("{旅游:报名编号}", "{旅游:活动名称}", "{旅游:活动编号}", "{旅游:报名用户昵称}", "{旅游:报名真实姓名}", "{旅游:报名人数}", "{旅游:报名人电话}", "{旅游:报名人身份证}","{旅游:报名状态}","{旅游:报名人应得积分}","{旅游:报名添加时间}","{旅游:报名是否使用积分}","{旅游:活动总人数}","{旅游:活动行程}","{旅游:活动协议}","{旅游:活动领队}","{旅游:活动积分}","{旅游:活动状态}","{旅游:活动出发时间}","{旅游:活动返回时间}","{旅游:活动添加时间}","{旅游:活动介绍}","{旅游:论坛链接}")
	    bb = Array(b.rs("baoMingId"), b.rs("baoMingHuoDong"), b.rs("baoMingHuoDongId"), b.rs("baoMingUserName"), cstr(b.rs("baoMingRealName")), cstr(b.rs("baoMingRenshu")),b.rs("baoMingPhone"),b.rs("baoMingIdcard"),b.rs("baoMingZhuangtai"),b.rs("baoMingJifen"),formatdatetime(b.rs("baoMingAddTime"), 2),b.rs("baoMingUseJifen"),b.rs("huoDongrenshu"),b.rs("huoDongxingcheng"),b.rs("huoDongxieyi"),b.rs("huoDonglingdui"),b.rs("huoDongjifen"),b.rs("huoDongZhuangtai"),b.rs("huoDongChufaTime"),b.rs("huoDongFanhuiTime"),b.rs("huoDongAddtime"),b.rs("huoDongInfo"),b.rs("bbsLink"))
	    For i = 0 To UBound(bb)
	        'zz = Replace(qzy, aa(i), bb(i))
	        qzy = Replace(qzy, aa(i), bb(i))
	    Next
    end if
    baomingViewSuccess = qzy
    b.rs.Close()
    set b.rs = nothing
    Set b = Nothing
End Function

function huoDongZtBack(byval str)
	strT = ""
	select case str
		case "0"
			'strT = "有空缺"
			strT = "<img src=""/images/weiman.gif"" alt='活动报名未满,还有空缺!' />"
		case "1"
			'strT = "报名已满"
			strT = "<img src=""/images/yiman.gif"" alt='活动报名人数已满,请联系相关人员!' />"
		case "2"
			'strT = "活动出发"
			strT = "<img src=""/images/chufa.gif"" alt='活动已经出发咯!' />"
		case "3"
			'strT = "活动结束"
			strT = "<img src=""/images/jieshu.gif"" alt='人数已满!' />"
	end select
	huoDongZtBack = strT
end function

function huoDongZtBackText(byval str)
	strT = ""
	select case str
		case "0"
			strT = "有空缺"
			'strT = "<img src=""/images/weiman.gif"" alt='活动报名未满,还有空缺!' />"
		case "1"
			strT = "报名已满"
			'strT = "<img src=""/images/yiman.gif"" alt='活动报名人数已满,请联系相关人员!' />"
		case "2"
			strT = "活动出发"
			'strT = "<img src=""/images/chufa.gif"" alt='活动已经出发咯!' />"
		case "3"
			strT = "活动结束"
			'strT = "<img src=""/images/jieshu.gif"" alt='人数已满!' />"
	end select
	huoDongZtBackText = strT
end function

'匹配列表循环

Function arclist(Str)
    Set reg = New regexp
    reg.IgnoreCase = False
    reg.Global = True
    reg.Pattern = "{旅游:文章列表开始([\s\S.]*?)}([\s\S.]*?){旅游:文章列表结束}"
    Set aa = reg.Execute(Str)
    For Each match in aa
        bq = match.submatches(0)
        nr = match.submatches(1)
        row = nbq(bq, "行")
        cid = nbq(bq, "类")
        tlen = nbq(bq, "长")
        tim = nbq(bq, "期")
        att = nbq(bq, "图")
        ord = nbq(bq, "排")
        If ord = 1 Then
            xx = "order by readcount desc"
        ElseIf ord = 2 Then
            xx = "and tuijian=1 order by newsid desc"
        ElseIf ord = 3 Then
            xx = "order by posttime desc"
        Else
            xx = "order by newsid desc"
        End If
        If tim = "" Or tim<0 Or tim>4 Then
            tim = 0
        Else
            tim = tim
        End If
        If tlen = "" Then
            tlen = 20
        Else
            tlen = CInt(tlen)
        End If
        info = nbq(bq, "截取")
        If info = "" Then
            info = 40
        Else
            info = CInt(info)
        End If
        Str = Replace(Str, match, lb(row, nr, cid, tlen, info, tim, att, xx))
    Next
    arclist = Str
    Set reg = Nothing
    Set aa = Nothing
End Function

'列表

Function lb(liebiao, qzy, lid, btcd, nrcd, ntt, atp, tj)
    Set ns = New news
    rs = ns.getnewslist(lid, liebiao, atp, tj)
    If ns.rs.EOF Then
        lb = "暂时没有内容"
    Else
        i = 0
        Dim zz(12)
        zz(0) = qzy
        Do While Not ns.rs.EOF
            i = i + 1
            if lid = "" then
            	aa = Array("{旅游:文章编号}", "{旅游:文章标题}", "{旅游:文章内容}", "{旅游:文章发表时间}", "{旅游:文章点击量}", "{旅游:文章封面图片}", "{旅游:文章内容地址}", "{旅游:文章循环数}", "{旅游:文章导读}","{旅游:文章图片地址}")
	            bb = Array(ns.rs("newsid"), cutStr(ns.rs("ntitle"), btcd), cutStr(RemoveHTML(ns.rs("ncontent")), nrcd), formatdatetime(ns.rs("posttime"), ntt), ns.rs("readcount"), ns.rs("img"), "view.asp?newsid="&ns.rs("newsid"), i, ns.rs("ninfo"),toppic(ns.rs("ncontent"),"html","/images/nopic.gif"))
            else
	            aa = Array("{旅游:文章编号}", "{旅游:文章标题}", "{旅游:文章内容}", "{旅游:文章发表时间}", "{旅游:文章点击量}", "{旅游:文章所属分类}", "{旅游:文章封面图片}", "{旅游:文章分类地址}", "{旅游:文章内容地址}", "{旅游:文章循环数}", "{旅游:文章导读}", "{旅游:文章图片地址}")
	            
	            bb = Array(ns.rs("newsid"), cutStr(ns.rs("ntitle"), btcd), cutStr(RemoveHTML(ns.rs("ncontent")), nrcd), formatdatetime(ns.rs("posttime"), ntt), ns.rs("readcount"), ns.rs("cname"), ns.rs("img"), "list.asp?classid="&ns.rs("news.cid"), "view.asp?newsid="&ns.rs("newsid"), i, ns.rs("ninfo"),toppic(ns.rs("ncontent"),"html","/images/nopic.gif"))
			end if
            For j = 0 To UBound(aa)
                zz(j + 1) = Replace(zz(j), aa(j), bb(j))
            Next
            cc = cc + zz(j)
            ns.rs.movenext
        Loop
        lb = cc
    End If
    ns.rs.Close()
    Set ns = Nothing
End Function


Function iFThenElse(Str)
    Set reg = New regexp
    reg.IgnoreCase = False
    reg.Global = True
    reg.Pattern = "{旅游:判断开始(([\s\S.]*?)([><=])([\s\S.]*?)?)}([\s\S.]*?){旅游:否则}([\s\S.]*?){旅游:判断结束}"
    Set aa = reg.Execute(Str)
    For Each match in aa
        aab = match.submatches(1)
        ccd = match.submatches(2)
        bbb = match.submatches(3)
        thenContent = match.submatches(4)
        elseContent = match.submatches(5)
        select case ccd
        	case ">"
        		if aab > bbb then
        			Stra = thenContent
        		else
        			Stra = elseContent
        		end if
        	case "<"
        		if aab < bbb then
        			Stra = thenContent
        		else
        			Stra = elseContent
        		end if
        	case "="
        		if aab = bbb then
        			Stra = thenContent
        		else
        			Stra = elseContent
        		end if
        end select
        Str = Replace(Str, match, Stra)
    Next
    iFThenElse = Str
    Set reg = Nothing
    Set aa = Nothing
End Function


'匹配列表循环

Function userList(Str)
    Set reg = New regexp
    reg.IgnoreCase = False
    reg.Global = True
    reg.Pattern = "{旅游:用户列表开始([\s\S.]*?)}([\s\S.]*?){旅游:用户列表结束}"
    Set aa = reg.Execute(Str)
    For Each match in aa
        bq = match.submatches(0)
        nr = match.submatches(1)
        row = nbq(bq, "行")
        ord = nbq(bq, "排")
        If ord = 1 Then
            xx = "order by userId desc"
        ElseIf ord = 2 Then
            xx = "order by joinTime desc"
        End If
        Str = Replace(Str, match, userListNei(row,nr,xx))
    Next
    userList = Str
    Set reg = Nothing
    Set aa = Nothing
End Function

'列表

Function userListNei(row,qzy,orderby)
    Set u = New Users
    rs = u.getUsers(row,orderby)
    If u.rs.EOF Then
        userListNei = "暂时没有内容"
    Else
        i = 0
        Dim zz(11)
        zz(0) = qzy
        Do While Not u.rs.EOF
            i = i + 1
            	aa = Array("{旅游:用户编号}", "{旅游:用户名称}", "{旅游:用户密码}", "{旅游:用户真名}", "{旅游:用户性别}", "{旅游:用户电话}", "{旅游:身份证号}", "{旅游:加入时间}", "{旅游:用户循环数}", "{旅游:论坛用户ID}")
	            bb = Array(u.rs("userId"), u.rs("userName"), u.rs("password"), u.rs("RealName"), u.rs("sex"), u.rs("phone"),u.rs("idcard"),formatdatetime(u.rs("joinTime"), 2), i, Request.cookies(CookieName)("Uid"))
            For j = 0 To UBound(aa)
                zz(j + 1) = Replace(zz(j), aa(j), bb(j))
            Next
            cc = cc + zz(j)
            u.rs.movenext
        Loop
        userListNei = cc
    End If
    u.rs.Close()
    set u.rs = nothing
    Set u = Nothing
End Function



Function huoDongList(Str)
    Set reg = New regexp
    reg.IgnoreCase = False
    reg.Global = True
    reg.Pattern = "{旅游:活动列表开始([\s\S.]*?)}([\s\S.]*?){旅游:活动列表结束}"
    Set aa = reg.Execute(Str)
    For Each match in aa
        bq = match.submatches(0)
        nr = match.submatches(1)
        row = nbq(bq, "行")
        ord = nbq(bq, "排")
        If ord = 1 Then
            xx = "order by huoDongId desc"
        ElseIf ord = 2 Then
            xx = "order by huoDongAddTime desc"
        End If
        info = nbq(bq, "截取")
        If info = "" Then
            info = 40
        Else
            info = CInt(info)
        End If
        Str = Replace(Str, match, huoDongListNei(row,nr,xx,info))
    Next
    huoDongList = Str
    Set reg = Nothing
    Set aa = Nothing
End Function


Function huoDongZJList(Str)
    Set reg = New regexp
    reg.IgnoreCase = False
    reg.Global = True
    reg.Pattern = "{旅游:活动终极列表开始([\s\S.]*?)}([\s\S.]*?){旅游:活动终极列表结束}"
    Set aa = reg.Execute(Str)
    For Each match in aa
        bq = match.submatches(0)
        nr = match.submatches(1)
        row = nbq(bq, "行")
        ord = nbq(bq, "排")
        If ord = 1 Then
            xx = "order by huoDongId desc"
        ElseIf ord = 2 Then
            xx = "order by huoDongAddTime desc"
        End If
        info = nbq(bq, "截取")
        If info = "" Then
            info = 40
        Else
            info = CInt(info)
        End If
        Str = Replace(Str, match, huoDongListNei(row,nr,xx,info))
        Str = Replace(Str, "{旅游:活动终极列表导航}", pagesHD(row,xx))
    Next
    huoDongZJList = Str
    Set reg = Nothing
    Set aa = Nothing
End Function


'列表

Function huoDongListNei(row,qzy,orderby,info)

    Set h = New HuoDong
	Call h.getAutoZhuangtaiList()
	Do While Not h.rs2.eof
		If h.rs2("huoDongZhuangtai") = 1 Then'活动已经满员的状态！
			If DateDiff("d",h.rs2("huoDongChufaTime"),date())>=0 Then
				Call h.updateHuodongzhuangtai("2",h.rs2("huoDongId"))
			End if
		End If
		If h.rs2("huoDongZhuangtai") = 2 Then'活动已经出发的状态！
			If DateDiff("d",h.rs2("huoDongFanhuiTime"),date())>=0 Then
				Call h.updateHuodongzhuangtai("3",h.rs2("huoDongId"))
			End if
		End If
		h.rs2.movenext
	loop
    rs = h.getHuoDongs(row,orderby)
    If h.rs.EOF Then
        huoDongListNei = "暂时没有内容"
    Else
        i = 0
        Dim zz(21)
        zz(0) = qzy
        Do While Not h.rs.EOF
            i = i + 1
			aa = Array("{旅游:活动编号}","{旅游:论坛链接}", "{旅游:活动名称}", "{旅游:活动人数}", "{旅游:活动剩余人数}", "{旅游:活动行程}", "{旅游:活动协议}", "{旅游:活动领队}", "{旅游:活动积分}","{旅游:活动状态}","{旅游:活动状态文字版}","{旅游:活动出发时间}","{旅游:活动返回时间}","{旅游:活动添加时间}","{旅游:活动循环数}","{旅游:活动关键字}","{旅游:活动介绍}","{旅游:活动已报人数}","{旅游:活动链接地址}","{旅游:活动图片地址}","{旅游:列表页活动名称}")
		    bb = Array(h.rs("huoDongId"),h.rs("bbsLink") ,cutStr(h.rs("huoDongName"), info), h.rs("huoDongRenShu"), h.rs("huoDongShengyu"), h.rs("huoDongXingcheng"), h.rs("huoDongxieyi"),h.rs("huoDongLingdui"),h.rs("huoDongJifen"),huoDongZtBack(h.rs("huoDongZhuangtai")),huoDongZtBackText(h.rs("huoDongZhuangtai")),formatdatetime(h.rs("huoDongChufaTime"), 2),formatdatetime(h.rs("huoDongFanhuiTime"), 2),h.rs("huoDongAddTime"),i,h.rs("nkeyword"),h.rs("huoDongInfo"),h.rs("huoDongRenShu")-h.rs("huoDongShengyu"),"HuoDong.asp?huoDongId="&h.rs("huoDongId"),toppic(h.rs("huoDongXingcheng"),"html","/images/nopic.gif"), cutStr(h.rs("huoDongName"), info))
            For j = 0 To UBound(aa)
                zz(j + 1) = Replace(zz(j), aa(j), bb(j))
            Next
            cc = cc + zz(j)
            h.rs.movenext
        Loop
        huoDongListNei = cc
    End If
    h.rs.Close()
    set h.rs = nothing
    Set h = Nothing
End Function

Function RemoveHTML(str)
	Dim RegEx
	Set RegEx = New RegExp
	RegEx.Pattern = "<[^>]*>"
	RegEx.Global = True
	RemoveHTML = RegEx.Replace(str, "")
End Function

Function toppic(code,leixing,nopic)
        set regex = new regexp
        regex.ignorecase = true
        regex.global = true
        if leixing = "html" then
                regex.pattern = "<img [^>]*src=""([^"">]+)""[^>]+>"
        else
                regex.pattern = "\[img\]([^\u005B]+)"
        end if
        set matches = regex.execute(code)
        if regex.test(code) then
              toppic = matches(0).submatches(0)
        else
                toppic = nopic
        end if
End Function



'baomingList(yuke)

Function baomingList(Str)
    Set reg = New regexp
    reg.IgnoreCase = False
    reg.Global = True
    reg.Pattern = "{旅游:报名终极列表开始([\s\S.]*?)}([\s\S.]*?){旅游:报名终极列表结束}"
    Set aa = reg.Execute(Str)
    For Each match in aa
        bq = match.submatches(0)
        nr = match.submatches(1)
        row = nbq(bq, "行")
        ord = nbq(bq, "排")
        uName = vbsUnEscape(Replace(Request.Cookies(CookieName)("UserName"),"'","''"))
        Action = Request.QueryString("Action")
        zt = nbq(bq, "状态")
        If ord = 1 Then
            xx = " order by baoMingId desc"
        ElseIf ord = 2 Then
            xx = " order by baoMingAddTime desc"
        End If
        If row = "" Then
            row = 20
        Else
            row = row
        End If
        Str = Replace(Str, match, baomingListNei(uName,row,nr,xx,zt))
        Str = Replace(Str, "{旅游:报名终极列表导航}", pagesBm(Action,uName,row,xx,zt))
    Next
    baomingList = Str
    Set reg = Nothing
    Set aa = Nothing
End Function

'列表

Function baomingListNei(uName,row,qzy,orderby,zhuangtai)
    Set b = New Baoming
	call b.getUserCanJiabaoMingList(uName,zhuangtai,orderby)
    If b.rs.EOF Then
        baomingListNei = "暂时没有内容"
    Else
        b.rs.pagesize = row
        page = CLng(request("page"))
        If page<1 Then page = 1
        If page>b.rs.pagecount Then page = b.rs.pagecount
        b.rs.absolutepage = page
        
        if zhuangtai <> "0" then
	        Dim zz(18)
	        zz(0) = qzy
	        For i = 1 To b.rs.pagesize
	            If b.rs.EOF Then Exit For
				aa = Array("{旅游:报名编号}", "{旅游:报名活动名称}", "{旅游:报名活动编号}", "{旅游:报名用户}", "{旅游:报名用户真名}", "{旅游:报名人数}", "{旅游:报名电话}", "{旅游:报名身份证}","{旅游:报名状态}","{旅游:报名积分}","{旅游:报名添加时间}","{旅游:报名是否使用积分}","{旅游:报名循环数}","{旅游:活动状态}","{旅游:活动状态文字版}","{旅游:活动链接地址}","{旅游:活动编号}","{旅游:论坛链接}")
			    bb = Array(b.rs("baoMingId"), b.rs("baoMingHuoDong"), b.rs("baoMingHuoDongId"), b.rs("baoMingUserName"), b.rs("baoMingRealName"), b.rs("baoMingRenshu"),b.rs("baoMingPhone"),b.rs("baoMingIdcard"),b.rs("baoMingZhuangtai"),b.rs("baoMingJifen"),formatdatetime(b.rs("baoMingAddTime"), 2),UseJifen(b.rs("baoMingUseJifen")),i,huoDongZtBack(b.rs("huoDongZhuangtai")),huoDongZtBackText(b.rs("huoDongZhuangtai")),"HuoDong.asp?huoDongId="&b.rs("huoDongId"),b.rs("huoDongId"),b.rs("bbsLink"))
	            For j = 0 To UBound(aa)
	                zz(j + 1) = Replace(zz(j), aa(j), bb(j))
	            Next
	            cc = cc + zz(j)
	            b.rs.movenext
	        Next
	    else
		    reDim zz(19)
	        zz(0) = qzy
	        Do While Not b.rs.EOF
	            i = i + 1
				aa = Array("{旅游:活动编号}","{旅游:论坛链接}", "{旅游:活动名称}", "{旅游:活动人数}", "{旅游:活动剩余人数}", "{旅游:活动行程}", "{旅游:活动协议}", "{旅游:活动领队}", "{旅游:活动积分}","{旅游:活动状态}","{旅游:活动状态文字版}","{旅游:活动出发时间}","{旅游:活动返回时间}","{旅游:活动添加时间}","{旅游:活动循环数}","{旅游:活动关键字}","{旅游:活动介绍}","{旅游:活动已报人数}","{旅游:活动链接地址}")
			    bb = Array(b.rs("huoDongId"),b.rs("bbsLink") ,b.rs("huoDongName"), b.rs("huoDongRenShu"), b.rs("huoDongShengyu"), b.rs("huoDongXingcheng"), b.rs("huoDongxieyi"),b.rs("huoDongLingdui"),b.rs("huoDongJifen"),huoDongZtBack(b.rs("huoDongZhuangtai")),huoDongZtBackText(b.rs("huoDongZhuangtai")),formatdatetime(b.rs("huoDongChufaTime"), 2),formatdatetime(b.rs("huoDongFanhuiTime"), 2),b.rs("huoDongAddTime"),i,b.rs("nkeyword"),b.rs("huoDongInfo"),b.rs("huoDongRenShu")-b.rs("huoDongShengyu"),"HuoDong.asp?huoDongId="&b.rs("huoDongId"))
	            For j = 0 To UBound(aa)
	                zz(j + 1) = Replace(zz(j), aa(j), bb(j))
	            Next
	            cc = cc + zz(j)
	            b.rs.movenext
	        Loop
	    end if
        baomingListNei = cc
    End If
    b.rs.Close()
    set b.rs = nothing
    Set b = Nothing
End Function


Function baomingShowList(Str)
    Set reg = New regexp
    reg.IgnoreCase = False
    reg.Global = True
    reg.Pattern = "{旅游:报名显示列表开始([\s\S.]*?)}([\s\S.]*?){旅游:报名显示列表结束}"
    Set aa = reg.Execute(Str)
    For Each match in aa
        bq = match.submatches(0)
        nr = match.submatches(1)
        row = nbq(bq, "行")
        ord = nbq(bq, "排")
        huoDongId = Request.QueryString("huoDongId")
        zt = nbq(bq, "状态")
        If ord = 1 Then
            xx = " order by baoMingId desc"
        ElseIf ord = 2 Then
            xx = " order by baoMingAddTime desc"
        End If
        If row = "" Then
            row = 20
        Else
            row = row
        End If
        Str = Replace(Str, match, baomingShowListNei(huoDongId,row,nr,xx,zt))
    Next
    baomingShowList = Str
    Set reg = Nothing
    Set aa = Nothing
End Function

'列表

Function baomingShowListNei(huoDongId,row,qzy,orderby,zhuangtai)
    Set b = New Baoming
	call b.BaoMingPageList(huoDongId)
    If b.rs.EOF Then
        baomingShowListNei = "暂时没有内容"
    Else
        b.rs.pagesize = row
        page = CLng(request("page"))
        If page<1 Then page = 1
        If page>b.rs.pagecount Then page = b.rs.pagecount
        b.rs.absolutepage = page
        Dim zz(20)
        zz(0) = qzy
        For i = 1 To b.rs.pagesize
            If b.rs.EOF Then Exit For
			aa = Array("{旅游:报名用户编号}","{旅游:报名编号}", "{旅游:报名活动名称}", "{旅游:报名活动编号}", "{旅游:报名用户}", "{旅游:报名用户真名}", "{旅游:报名人数}", "{旅游:报名电话}", "{旅游:报名身份证}","{旅游:报名状态}","{旅游:报名积分}","{旅游:报名添加时间}","{旅游:报名是否使用积分}","{旅游:报名循环数}","{旅游:活动状态}","{旅游:活动状态文字版}","{旅游:活动链接地址}","{旅游:活动编号}","{旅游:论坛链接}","{旅游:简短活动名称}")
		    bb = Array(b.getID(b.rs("baoMingren")), b.rs("baoMingId"), b.rs("baoMingHuoDong"), b.rs("baoMingHuoDongId"), b.rs("baoMingUserName"), b.rs("baoMingRealName"), b.rs("baoMingRenshu"),b.rs("baoMingPhone"),b.rs("baoMingIdcard"),b.rs("baoMingZhuangtai"),b.rs("baoMingJifen"),formatdatetime(b.rs("baoMingAddTime"), 2),UseJifen(b.rs("baoMingUseJifen")),i,huoDongZtBack(b.rs("huoDongZhuangtai")),huoDongZtBackText(b.rs("huoDongZhuangtai")),"HuoDong.asp?huoDongId="&b.rs("huoDongId"),b.rs("huoDongId"),b.rs("bbsLink"),cutStr(b.rs("baoMingHuoDong"),25))
            For j = 0 To UBound(aa)
                zz(j + 1) = Replace(zz(j), aa(j), bb(j))
            Next
            cc = cc + zz(j)
            b.rs.movenext
        Next
        baomingShowListNei = cc
    End If
    b.rs.Close()
    set b.rs = nothing
    Set b = Nothing
End Function

function UseJifen(byval jifen)
	if jifen="0" then
		UseJifen = "否"
	else
		UseJifen = "是"
	end if
end function

Function isBaoxian(byval baoxian)
	if baoxian<>"" and baoxian = "1" then
		isBaoxian = "YES"
	else
		isBaoxian = "NO"
	end if
End Function

Function isBaoxianCn(byval baoxian)
	if baoxian<>"" and baoxian = "1" then
		isBaoxianCn = "已申请"
	else
		isBaoxianCn = "未申请"
	end if
End Function


Function slide(Str)
    Set reg = New regexp
    reg.IgnoreCase = False
    reg.Global = True
    reg.Pattern = "{旅游:幻灯列表开始([\s\S.]*?)}([\s\S.]*?){旅游:幻灯列表结束}"
    Set aa = reg.Execute(Str)
    For Each match in aa
        bq = match.submatches(0)
        nr = match.submatches(1)
        row = nbq(bq, "行")
        cid = nbq(bq, "类")
        tlen = nbq(bq, "长")
        tim = nbq(bq, "期")
        ord = nbq(bq, "排")
        If ord = 1 Then
            xx = "attpic=1 and huandeng=1 order by readcount desc"
        ElseIf ord = 2 Then
            xx = "attpic=1 and huandeng=1 and tuijian=1 order by newsid desc"
        ElseIf ord = 3 Then
            xx = "attpic=1 and huandeng=1 order by posttime desc"
        Else
            xx = "attpic=1 and huandeng=1 order by newsid desc"
        End If
        If tim = "" Or tim<0 Or tim>4 Then
            tim = 0
        Else
            tim = tim
        End If
        If tlen = "" Then
            tlen = 20
        Else
            tlen = CInt(tlen)
        End If
        'Str = Replace(Str, match, slideContent(row, nr, cid, tlen, tim, xx))
        Str = Replace(Str, match, slideContent2(row, nr, cid, tlen, tim, xx))
    Next
    slide = Str
    Set reg = Nothing
    Set aa = Nothing
End Function

'列表

Function slideContent(liebiao, qzy, lid, btcd, ntt, tj)
    Set ns = New news
    rs = ns.getSlideNewsList(lid, liebiao, tj)
    If ns.rs.EOF Then
        slideContent = "暂时没有内容"
    Else
        i = 0
        Dim zz(11),titles,links,pics
        zz(0) = qzy
        Do While Not ns.rs.EOF
            i = i + 1
            if titles <>"" then
				titles = titles&"|"&CutStr(ns.rs("ntitle"),btcd)
				links = links&"|"&"view.asp?newsid="&ns.rs("newsid")
				pics = pics&"|"&ns.rs("img")
			else
				titles = CutStr(ns.rs("ntitle"),btcd)
				links = "view.asp?newsid="&ns.rs("newsid")
				pics = ns.rs("img")
			end if
            aa = Array("{旅游:幻灯标题序列}", "{旅游:幻灯链接序列}", "{旅游:幻灯图片序列}")
            bb = Array(titles, links, pics)
            ns.rs.movenext
        Loop
        For j = 0 To UBound(aa)
            zz(j + 1) = Replace(zz(j), aa(j), bb(j))
        Next
        cc = cc + zz(j)
        slideContent = cc
    End If
    ns.rs.Close()
    Set ns = Nothing
End Function


Function slideContent2(liebiao, qzy, lid, btcd, ntt, tj)
    Set ns = New news
    rs = ns.getSlideNewsList(lid, liebiao, tj)
    If ns.rs.EOF Then
        slideContent = "暂时没有内容"
    Else
        i = 0
        Dim zz(11),titles,links,pics
        zz(0) = qzy
        Do While Not ns.rs.EOF
            i = i + 1
			titles = CutStr(ns.rs("ntitle"),btcd)
			'links = Server.URLENCODE(ns.rs("ninfo"))
			links = "view.asp?newsid="&ns.rs("newsid")
			pics = ns.rs("img")
            aa = Array("{旅游:幻灯标题}", "{旅游:幻灯链接}", "{旅游:幻灯图片}")
            bb = Array(titles, links, pics)
            For j = 0 To UBound(aa)
	            zz(j + 1) = Replace(zz(j), aa(j), bb(j))
	        Next
	        cc = cc + zz(j)
            ns.rs.movenext
        Loop
        
        slideContent2 = cc
    End If
    ns.rs.Close()
    Set ns = Nothing
End Function


Function FriendLinklist(Str)
    Set reg = New regexp
    reg.IgnoreCase = False
    reg.Global = True
    reg.Pattern = "{旅游:友情链接开始([\s\S.]*?)}([\s\S.]*?){旅游:友情链接结束}"
    Set aa = reg.Execute(Str)
    For Each match in aa
        bq = match.submatches(0)
        nr = match.submatches(1)
        row = nbq(bq, "行")
        tlen = nbq(bq, "长")
        att = nbq(bq, "图")
        If tim = "" Or tim<0 Or tim>4 Then
            tim = 0
        Else
            tim = tim
        End If
        If tlen = "" Then
            tlen = 20
        Else
            tlen = CInt(tlen)
        End If
        info = nbq(bq, "截取")
        If info = "" Then
            info = 40
        Else
            info = CInt(info)
        End If
        Str = Replace(Str, match, FriendLink(row, nr, tlen, info, att))
    Next
    FriendLinklist = Str
    Set reg = Nothing
    Set aa = Nothing
End Function

'列表

Function FriendLink(liebiao, qzy, btcd, nrcd, atp)
    Set ns = New news
    rs = ns.getfriendlink(liebiao, atp)
    If ns.rs.EOF Then
        FriendLink = "暂时没有内容"
    Else
        i = 0
        Dim zz(11)
        zz(0) = qzy
        Do While Not ns.rs.EOF
            i = i + 1
            aa = Array("{旅游:友情链接编号}", "{旅游:友情链接标题}", "{旅游:友情链接循环数}", "{旅游:友情链接地址}")
            bb = Array(ns.rs("id"), cutStr(ns.rs("link_name"), btcd), i, ns.rs("link_url"))
            For j = 0 To UBound(aa)
                zz(j + 1) = Replace(zz(j), aa(j), bb(j))
            Next
            cc = cc + zz(j)
            ns.rs.movenext
        Loop
        FriendLink = cc
    End If
    ns.rs.Close()
    Set ns = Nothing
End Function





'匹配分类列表

Function classlist(Str)

    Set reg = New regexp
    reg.IgnoreCase = False
    reg.Global = True
    reg.Pattern = "{旅游:终极列表开始([\s\S.]*?)}([\s\S.]*?){旅游:终极列表结束}"
    Set aa = reg.Execute(Str)
    For Each match in aa
        bq = match.submatches(0)
        nr = match.submatches(1)
        Size = nbq(bq, "行")
        tlen = nbq(bq, "长")
        tim = nbq(bq, "期")
        If tim = "" Or tim<0 Or tim>4 Then
            tim = 0
        Else
            tim = tim
        End If
        If tlen = "" Then
            tlen = 20
        Else
            tlen = CInt(tlen)
        End If
        info = nbq(bq, "截取")
        If info = "" Then
            info = 40
        Else
            info = CInt(info)
        End If
        If Size = "" Then
            Size = 20
        Else
            Size = Size
        End If
        Str = Replace(Str, match, pagelb(nr, Size, tlen, info, tim))
        Str = Replace(Str, "{旅游:终极列表导航}", pages(Size))
    Next
    classlist = Str
    Set reg = Nothing
    Set aa = Nothing
End Function

'频道列表列表

Function pagelb(qzy, Size, btcd, nrcd, ntt)
    Set ns = New news
    lid = classid
    rs = ns.pagelist(lid)
    If ns.rs.EOF Then
        pagelb = "该分类下暂时没有内容！"
    Else
        ns.rs.pagesize = Size
        If html = 1 Then
            page = pagel
        Else
            page = CLng(request("page"))
        End If
        If page<1 Then page = 1
        If page>ns.rs.pagecount Then page = ns.rs.pagecount
        ns.rs.absolutepage = page
        Dim zz(12)
        zz(0) = qzy
        For i = 1 to ns.rs.pagesize
            If ns.rs.EOF Then Exit For
            aa = Array("{旅游:文章编号}", "{旅游:文章标题}", "{旅游:文章内容}", "{旅游:文章发布时间}", "{旅游:文章点击数}", "{旅游:文章所属分类}", "{旅游:文章封面图片}", "{旅游:文章分类地址}", "{旅游:文章内容地址}", "{旅游:文章循环数}", "{旅游:文章导读}","{旅游:文章图片地址}")
            bb = Array(ns.rs("newsid"), cutStr(ns.rs("ntitle"), btcd), cutStr(RemoveHTML(ns.rs("ncontent")), nrcd), formatdatetime(ns.rs("posttime"), ntt), ns.rs("readcount"), ns.rs("cname"), ns.rs("img"), "list.asp?classid="&ns.rs("news.cid"), "view.asp?newsid="&ns.rs("newsid"), i, ns.rs("ninfo"),toppic(ns.rs("ncontent"),"html","/images/nopic.gif"))
            For j = 0 To UBound(aa)

                zz(j + 1) = Replace(zz(j), aa(j), bb(j))
            Next
            cc = cc + zz(j)
            ns.rs.movenext
        Next

        pagelb = cc
    End If
    ns.rs.Close()
    Set ns = Nothing
End Function


'匹配memu

Function memu(Str)
    Set reg = New regexp
    reg.IgnoreCase = False
    reg.Global = True
    reg.Pattern = "{旅游:分类开始([\s\S.]*?)}([\s\S.]*?){旅游:分类结束}"
    Set aa = reg.Execute(Str)
    For Each match in aa
        bq = match.submatches(0)
        nr = match.submatches(1)
        row = nbq(bq, "行")
        Str = Replace(Str, match, mlb(row, nr))
    Next
    memu = Str
    Set reg = Nothing
    Set aa = Nothing
End Function

'列表

Function mlb(mrow, qzy)
    Set ct = New category
    rs = ct.zdfl(mrow)
    Dim zz(5)
    zz(0) = qzy
    Do While Not ct.rs.EOF
        aa = Array("{旅游:分类标题}", "{旅游:分类编号}", "{旅游:分类关键词}", "{旅游:分类简介}", "{旅游:分类链接地址}")
        If ct.rs("linkture") = 0 Then
            bb = Array(ct.rs("cname"), ct.rs("cid"), ct.rs("keyword"), ct.rs("info"), "list.asp?classid="&ct.rs("cid"))
        Else
            bb = Array(ct.rs("cname"), ct.rs("cid"), ct.rs("keyword"), ct.rs("info"), ct.rs("link"))
        End If
        For j = 0 To UBound(aa)
            zz(j + 1) = Replace(zz(j), aa(j), bb(j))
        Next
        cc = cc + zz(j)
        ct.rs.movenext
    Loop
    mlb = cc
    ct.rs.Close()
    Set ct = Nothing
End Function

'匹配小memu

Function smemu(Str)
    Set reg = New regexp
    reg.IgnoreCase = False
    reg.Global = True
    reg.Pattern = "{旅游:获取子类开始([\s\S.]*?)}([\s\S.]*?){旅游:获取子类结束}"
    Set aa = reg.Execute(Str)
    For Each match in aa
        bq = match.submatches(0)
        nr = match.submatches(1)
        row = nbq(bq, "行")
        cid = nbq(bq, "类")
        Str = Replace(Str, match, zmlb(row, nr, cid))
    Next
    smemu = Str
    Set reg = Nothing
    Set aa = Nothing
End Function

'子分类列表

Function zmlb(mrow, qzy, mcid)
    If mcid = "" Then mcid = classid
    Set ct = New category
    rs = ct.zzdfl(mcid)
    Dim zz(5)
    zz(0) = qzy
    Do While Not ct.rs.EOF
        aa = Array("{旅游:分类名称}", "{旅游:分类编号}", "{旅游:分类关键字}", "{旅游:分类介绍}", "{旅游:分类链接地址}")
        If ct.rs("linkture") = 0 Then
            bb = Array(ct.rs("cname"), ct.rs("cid"), ct.rs("keyword"), ct.rs("info"), "list.asp?classid="&ct.rs("cid"))
        Else
            bb = Array(ct.rs("cname"), ct.rs("cid"), ct.rs("keyword"), ct.rs("info"), ct.rs("link"))
        End If
        For j = 0 To UBound(aa)
            zz(j + 1) = Replace(zz(j), aa(j), bb(j))
        Next
        cc = cc + zz(j)
        ct.rs.movenext
    Loop
    zmlb = cc
    ct.rs.Close()
    Set ct = Nothing
End Function

'频道分页

Function pages(Size)
    Set ns = New news
    rs = ns.pagelist(classid)
    If ns.rs.EOF Then
        pages = "首页&nbsp;上一页&nbsp下一页&nbsp;尾页"
    Else
        ns.rs.pagesize = Size
        If html = 1 Then
            page = pagel
        Else
            page = CLng(request("page"))
        End If
        If page<1 Then page = 1
        If page>ns.rs.pagecount Then page = ns.rs.pagecount
        ns.rs.absolutepage = page
        If page = 1 Then
            pages = "首页&nbsp;上一页&nbsp;第"&page&"页&nbsp;<a href='list.asp?classid="&classid&"&page="&page + 1&"'>下一页</a>&nbsp;<a href='list.asp?classid="&classid&"&page="&ns.rs.pagecount&"'>尾页</a>"
        ElseIf page = ns.rs.pagecount Then
            pages = "<a href='list.asp?classid="&classid&"&page=1'>首页</a>&nbsp;<a href='list.asp?classid="&classid&"&page="&page -1&"'>上一页</a>&nbsp;第"&page&"页&nbsp;下一页&nbsp;尾页"
        Else pages = "<a href='list.asp?classid="&classid&"&page=1'>首页</a>&nbsp;<a href='list.asp?classid="&classid&"&page="&page -1&"'>上一页</a>&nbsp;第"&page&"页&nbsp;<a href='list.asp?classid="&classid&"&page="&page + 1&"'>下一页</a>&nbsp;<a href='list.asp?classid="&classid&"&page="&ns.rs.pagecount&"'>尾页</a>"
        End If
    End If
    ns.rs.Close()
    Set ns = Nothing
End Function


'频道分页

Function pagesBm(Action,uName,row,zt,xx)
    Set b = New Baoming
    rs = b.getUserCanJiabaoMingList(uName,zt,xx)
    If b.rs.EOF Then
        pagesBm = "首页&nbsp;上一页&nbsp;下一页&nbsp;尾页"
    Else
        b.rs.pagesize = row
        page = CLng(request("page"))
        If page<1 Then page = 1
        If page>b.rs.pagecount Then page = b.rs.pagecount
        b.rs.absolutepage = page
        If page = 1 Then
            pagesBm = "首页&nbsp;上一页&nbsp;第"&page&"页&nbsp;<a href='/Users/Huodong.asp?Action="&Action&"&page="&page + 1&"'>下一页</a>&nbsp;<a href='/Users/Huodong.asp?Action="&Action&"&page="&b.rs.pagecount&"'>尾页</a>"
        ElseIf page = b.rs.pagecount Then
            pagesBm = "<a href='/Users/Huodong.asp?Action="&Action&"&page=1'>首页</a>&nbsp;<a href='/Users/Huodong.asp?Action="&Action&"&page="&page -1&"'>上一页</a>&nbsp;第"&page&"页&nbsp;下一页&nbsp;尾页"
        Else pagesBm = "<a href='/Users/Huodong.asp?Action="&Action&"&page=1'>首页</a>&nbsp;<a href='/Users/Huodong.asp?Action="&Action&"&page="&page -1&"'>上一页</a>&nbsp;第"&page&"页&nbsp;<a href='/Users/Huodong.asp?Action="&Action&"&page="&page + 1&"'>下一页</a>&nbsp;<a href='/Users/Huodong.asp?Action="&Action&"&page="&b.rs.pagecount&"'>尾页</a>"
        End If
    End If
    b.rs.Close()
    Set b = Nothing
End Function



Function pagesHD(row,orderby)
    Set h = New Huodong
    rs = h.getHuoDongs(row,orderby)
    If h.rs.EOF Then
        pagesHD = "首页&nbsp;上一页&nbsp;下一页&nbsp;尾页"
    Else
        h.rs.pagesize = row
        page = CLng(request("page"))
        If page<1 Then page = 1
        If page>h.rs.pagecount Then

        	page = h.rs.pagecount
        	if page<0 then page = 0
        end if
        h.rs.absolutepage = page
        If page = 1 Then
            pagesHD = "首页&nbsp;上一页&nbsp;第"&page&"页&nbsp;<a href='/HDList.asp?page="&page + 1&"'>下一页</a>&nbsp;<a href='/HDList.asp?page="&h.rs.pagecount&"'>尾页</a>"
        ElseIf page = h.rs.pagecount Then
            pagesHD = "<a href='/HDList.asp?page=1'>首页</a>&nbsp;<a href='/HDList.asp?page="&page -1&"'>上一页</a>&nbsp;第"&page&"页&nbsp;下一页&nbsp;尾页"
        Else pagesHD = "<a href='/HDList.asp?page=1'>首页</a>&nbsp;<a href='/HDList.asp?page="&page -1&"'>上一页</a>&nbsp;第"&page&"页&nbsp;<a href='/HDList.asp?page="&page + 1&"'>下一页</a>&nbsp;<a href='/HDList.asp?page="&h.rs.pagecount&"'>尾页</a>"
        End If
    End If
    h.rs.Close()
    Set h = Nothing
End Function

'匹配并替换内容标签

Function content(qzy)

    If newsid = "" Then
        content = Str
    Else
        Set ns = New news
        ns.getnewsinfo(newsid)
        Set ps = New news
        ps.getprenewsinfo(newsid)
        Set nes = New news
        nes.getnxtnewsinfo(newsid)
        If ps.rs.EOF Then
            cc = "前面没有文章了"
        Else
            cc = "<a href='"&install&"view.asp?newsid="&ps.rs("newsid")&"'>"&ps.rs("ntitle")&"</a>"
        End If
        If nes.rs.EOF Then
            dd = "后面没有文章了"
        Else
            dd = "<a href='"&install&"view.asp?newsid="&nes.rs("newsid")&"'>"&nes.rs("ntitle")&"</a>"
        End If



        aa = Array("{旅游:文章内容标题}", "{旅游:文章主体内容}", "{旅游:文章发布时间}", "{旅游:文章分类ID}", "{旅游:文章总点击数}", "{旅游:文章封面图片地址}", "{旅游:文章分类名称}", "{旅游:文章关键词}", "{旅游:文章内容导读}", "{旅游:tag}", "{旅游:上一篇文章}", "{旅游:下一篇文章}", "{旅游:文章内容地址}")

        bb = Array(ns.rs("ntitle"), o_link(ns.rs("ncontent")), ns.rs("posttime"), ns.rs("news.cid"), "<script language=JavaScript src='"&install&"inc/hits.asp?newsid="&ns.rs("newsid")&"'></script>", ns.rs("img"), ns.rs("cname"), tag(ns.rs("nkeyword"), 0), ns.rs("ninfo"), tag(ns.rs("nkeyword"), 1), cc, dd, "<a href='"&install&"view.asp?newsid="&ns.rs("newsid")&"'>"&ns.rs("ntitle")&"</a>")

        For i = 0 To UBound(bb)
            zz = Replace(qzy, aa(i), bb(i))
            qzy = Replace(zz, aa(i), bb(i))
        Next
        content = qzy
    End If
End Function

'相关文章标签

Function like_list(Str)
    Set reg = New regexp
    reg.IgnoreCase = False
    reg.Global = True
    reg.Pattern = "{旅游:相关文章开始([\s\S.]*?)}([\s\S.]*?){旅游:相关文章结束}"
    Set aa = reg.Execute(Str)
    For Each match in aa
        bq = match.submatches(0)
        nr = match.submatches(1)
        row = nbq(bq, "行")
        tlen = nbq(bq, "长")
        tim = nbq(bq, "期")

        If tim = "" Or tim<0 Or tim>4 Then
            tim = 0
        Else
            tim = tim
        End If
        If tlen = "" Then
            tlen = 20
        Else
            tlen = CInt(tlen)
        End If
        info = nbq(bq, "截取")
        If info = "" Then
            info = 40
        Else
            info = CInt(info)
        End If

        Str = Replace(Str, match, likelb(row, tlen, tim, newsid, nr))
    Next
    like_list = Str
    Set reg = Nothing
    Set aa = Nothing
End Function

'列表

Function likelb(liebiao, btcd, ntt, nid, qzy)
    Set ns2 = New news
    ns2.getnewsinfo(nid)

    If Trim(ns2.rs("nkeyword")) = "" Then
        ss = "暂时没有相关内容"
    Else
        '内判断开始
        kk = Split(Trim(ns2.rs("nkeyword")), ",")
        Dim mm
        For i = 0 To UBound(kk)
            ll = " or nkeyword like '%"&dkh(kk(i))&"%'"
            mm = mm + ll
        Next

        nn = cut(4, mm)
        Set ns = New news
        ff = ns.getlikelist(nn, liebiao, nid)
        If ns.rs.EOF Then
            ss = "暂时没有相关内容"
        Else
            Dim zz(5)
            zz(0) = qzy
            Do While Not ns.rs.EOF
                aa = Array("{旅游:相关文章编号}", "{旅游:相关文章标题}", "{旅游:相关文章时间}", "{旅游:相关文章图片}", "{旅游:相关文章地址}")
                bb = Array(ns.rs("newsid"), cutStr(ns.rs("ntitle"), btcd), formatdatetime(ns.rs("posttime"), ntt), ns.rs("img"), "view.asp?newsid="&ns.rs("newsid"))
                For j = 0 To UBound(aa)
                    zz(j + 1) = Replace(zz(j), aa(j), bb(j))
                Next
                cc = cc + zz(j)
                ns.rs.movenext
            Loop
            ss = cc
        End If
        '内判断结束

    End If
    likelb = ss
    ns2.rs.Close()
    Set ns2 = Nothing
End Function

'bq内循环匹配

Function nbq(xx, yy)
    Set reg = New regexp
    reg.IgnoreCase = False
    reg.Global = True
    reg.Pattern = ""&yy&"=([0-9]+)"
    Set matches = reg.Execute(xx)
    For Each match in matches
        aa = match.submatches(0)
    Next
    nbq = aa
    Set reg = Nothing
    Set matches = Nothing
End Function

Function tiaojian(xx, yy)
	dim flag
	flag = false
    Set reg = New regexp
    reg.IgnoreCase = False
    reg.Global = True
    reg.Pattern = "([\s\S.]*?)"&yy&"([\s\S.]*?)"
    Set matches = reg.Execute(xx)
    For Each match in matches
        aa = match.submatches(0)
        bb = match.submatches(1)
    Next
    select Case yy
    	Case ">"
    		if aa > bb then
    			flag = true
    		else
    			flag = false
    		end if
    	Case "<"
    		if aa < bb then
    			flag = true
    		else
    			flag = false
    		end if
    	Case "="
    		if aa = bb then
    			flag = true
    		else
    			flag = false
    		end if
    	Case Else
    end select
    tiaojian = flag
    Set reg = Nothing
    Set matches = Nothing
End Function

'截取字符串长度函数

Function cutStr(ntlen, strlen)
    Dim l, t, c
    l = Len(ntlen)
    t = 0
    For i = 1 To l
        c = Abs(Asc(Mid(ntlen, i, 1)))
        If c>255 Then
            t = t + 2
        Else
            t = t + 1
        End If
        If t>= strlen Then
            cutStr = Left(ntlen, i)
            Exit For
        Else
            cutStr = ntlen
        End If
    Next
    cutStr = Replace(cutStr, Chr(10), "")
End Function

'替换分类单页标签

Function pth(qzy)
    If classid = "" Then
        pth = qzy
    Else
        Set ct = New category
        ct.getcategoryinfo(classid)
        aa = Array("{旅游:分类标题}", "{旅游:分类编号}", "{旅游:分类简介}", "{旅游:分类关键词}")
        bb = Array(ct.rs("cname"), ct.rs("cid"), ct.rs("info"), ct.rs("keyword"))
        For i = 0 To UBound(bb)
            zz = Replace(qzy, aa(i), bb(i))
            qzy = Replace(zz, aa(i), bb(i))
        Next
        pth = qzy
    End If
    ct.rs.Close
    Set ct = Nothing
End Function

'留言提交表单
guest = "<form action='guest.asp?active=do' method='post' name='form1' id='form1'><table width='100%' border='1' cellspacing='1' bordercolor='#FFFFFF' bgcolor='#999999'><tr><td width='20%' bgcolor='#FFFFFF'>留言标题：</td><td bgcolor='#FFFFFF'><input type='text' name='title' id='title'></td></tr><tr><td width='20%' bgcolor='#FFFFFF'>您的名字：</td><td bgcolor='#FFFFFF'><input type='text' name='name' id='name'></td></tr><tr><td width='20%' bgcolor='#FFFFFF'>联系电话：</td><td bgcolor='#FFFFFF'><input type='text' name='tel' id='tel'></td></tr><tr><td width='20%' bgcolor='#FFFFFF'>电子邮箱：</td><td bgcolor='#FFFFFF'><input type='text' name='email' id='email'></td></tr><tr><td width='20%' bgcolor='#FFFFFF'>联系QQ：</td><td bgcolor='#FFFFFF'><input type='text' name='qq' id='qq'></td></tr><tr><td width='20%' bgcolor='#FFFFFF'>留言内容：</td><td bgcolor='#FFFFFF'><textarea name='content' cols='50' rows='7' id='content'></textarea></td></tr><tr><td colspan='2' bgcolor='#FFFFFF'><table width='100%' border='0'><tr><td align='center'><input type='submit' value='提交留言' onClick='return FormCheck(form1)' /></td></tr></table></td></tr></table></form>"

'留言内容调用

Function lyb(num)
    Set cn = New config
    cn.guest()
    If cn.rs.EOF Then
        lyb = "暂时没有留言"
    Else
        If num = "" Then num = 10
        cn.rs.pagesize = num
        page = CLng(request.QueryString("page"))
        If Not IsNumeric(page) Then response.Redirect(install&"fail.asp")
        If page<1 Then page = 1
        If page>cn.rs.pagecount Then page = cn.rs.pagecount
        cn.rs.absolutepage = page
        For i = 1 To cn.rs.pagesize
            If cn.rs.EOF Then Exit For
            If IsNull(cn.rs("hf")) Then
                lylb = "<table width='100%' border='1' cellspacing='1' bordercolor='#FFFFFF' bgcolor='#999999'><tr><td colspan='2' bgcolor='#FFFFFF'><b>标题：</b>"&cn.rs("title")&"&nbsp;&nbsp;<b>姓名：</b>"&cn.rs("name")&"&nbsp;&nbsp;<b>电话：</b>"&cn.rs("tel")&"<b>&nbsp;&nbsp;时间：</b>"&cn.rs("times")&"</td></tr><tr><td width='10%' bgcolor='#FFFFFF'>内容：</td><td bgcolor='#FFFFFF'>"&cn.rs("content")&"</td></tr></table>"
            Else
                lylb = "<table width='100%' border='1' cellspacing='1' bordercolor='#FFFFFF' bgcolor='#999999'><tr><td colspan='2' bgcolor='#FFFFFF'><b>标题：</b>"&cn.rs("title")&"&nbsp;&nbsp;<b>姓名：</b>"&cn.rs("name")&"&nbsp;&nbsp;<b>电话：</b>"&cn.rs("tel")&"<b>&nbsp;&nbsp;时间：</b>"&cn.rs("times")&"</td></tr><tr><td width='10%' bgcolor='#FFFFFF'>内容：</td><td bgcolor='#FFFFFF'>"&cn.rs("content")&"</td></tr><tr><td width='10%' bgcolor='#FFFFFF'><font color='ff0000'>回复：</font></td><td bgcolor='#FFFFFF'>"&cn.rs("hf")&"</td></tr></table>"
            End If
            xx = xx + lylb
            cn.rs.movenext
        Next
        If page = 1 Then
            lypages = "<table width='100%' border='1' cellspacing='1' bordercolor='#FFFFFF' bgcolor='#999999'><tr><td colspan='2' bgcolor='#FFFFFF' align='center'>首页&nbsp;上一页&nbsp;第"&page&"页&nbsp;<a href='guest.asp?page="&page + 1&"'>下一页</a>&nbsp;<a href='guest.asp?page="&cn.rs.pagecount&"'>尾页</a></td></tr></table>"
        ElseIf page = cn.rs.pagecount Then
            lypages = "<table width='100%' border='1' cellspacing='1' bordercolor='#FFFFFF' bgcolor='#999999'><tr><td colspan='2' bgcolor='#FFFFFF' align='center'><a href='guest.asp?page=1'>首页</a>&nbsp;<a href='guest.asp?page="&page -1&"'>上一页</a>&nbsp;第"&page&"页&nbsp;下一页&nbsp;尾页</td></tr></table>"
        Else
            lypages = "<table width='100%' border='1' cellspacing='1' bordercolor='#FFFFFF' bgcolor='#999999'><tr><td colspan='2' bgcolor='#FFFFFF' align='center'><a href='guest.asp?page=1'>首页</a>&nbsp;<a href='guest.asp?page="&page -1&"'>上一页</a>&nbsp;第"&page&"页&nbsp;<a href='guest.asp?page="&page + 1&"'>下一页</a>&nbsp;<a href='guest.asp?page="&cn.rs.pagecount&"'>尾页</a></td></tr></table>"
        End If
        lyb = xx + lypages
    End If
    cn.rs.Close()
    Set cn = Nothing
End Function

'生成目录

Function MakeNewsDir(foldername)
    Dim fso, f
    On Error Resume Next
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    Set f = fso.CreateFolder(Server.MapPath(foldername))
    MakeNewsDir = True
    Set fso = Nothing
End Function

'站内连接替换

Function o_link(Str)
    Set cn = New config
    rs = cn.olink_list()
    Do While Not cn.rs.EOF
        Str = Replace(Str, cn.rs("link_name"), "<a href='"&cn.rs("link_url")&"'>"&cn.rs("link_name")&"</a>")
        cn.rs.movenext
    Loop
    o_link = Str
    cn.rs.Close()
    Set cn = Nothing
End Function

'地址转向函数

Function Tourl(url)
    response.redirect(url)
    response.End
End Function

Function zx_url(smg, url)
    response.Write "<script>alert('"&smg&"');window.location.href='"&url&"';</script>"
    response.End
End Function

Function zx_url2(smg, url, exec)
    response.Write "<script>alert('"&smg&"');window.location.href='"&url&"';"&exec&"</script>"
    response.End
End Function

'切源码

Function cut(n, s)
    ss = Len(s)
    xx = Mid(s, n, ss)
    cut = xx
End Function

'过滤大括号{}

Function dkh(d)
    d = Replace(d, "{", "")
    d = Replace(d, "}", "")
    dkh = d
End Function

'tag替换成连接

Function tag(t, n)
    If t = "" Then
        tag = ""
    Else
        Dim k
        k = dkh(t)

        Set rs = server.CreateObject("adodb.recordset")
        Set rs.activeconnection = conn
        rs.cursortype = 3
        sql = "select * from tags where id in ("&k&")"
        rs.Open sql
        If n = 1 Then
            Do While Not rs.EOF
                If html = 0 Then
                    l = "<a href='"&install&"tag.asp?id="&rs("id")&"'>"&rs("tag_name")&"</a>&nbsp;&nbsp;"
                Else
                    l = "<a href='"&install&"tag-"&rs("id")&".html'>"&rs("tag_name")&"</a>&nbsp;&nbsp;"
                End If
                j = j + l
                rs.movenext
            Loop
            tag = j
        Else
            Do While Not rs.EOF
                l = ","&rs("tag_name")
                j = j + l
                rs.movenext
            Loop
            tag = cut(2, j)
        End If
    End If
End Function

'html转JS

Function html2js(dm)
    dm1 = Replace(dm, """", "[""]")
    dm2 = Replace(dm1, "/", "[/]")
    html2js = dm2
End Function

'JS转HTML

Function js2html(dm)
    dm1 = Replace(dm, "[""]", """")
    dm2 = Replace(dm1, "[/]", "/")
    js2html = dm2
End Function

'DM2JS

Function dm2js(dm)
    dm = Replace(dm, "[""]", "\""")
    dm = Replace(dm, "[/]", "\/")
    dm = Replace(dm, Chr(13), "")
    dm = Replace(dm, Chr(10), "")
    dm = Replace(dm, Chr(9), "")
    dm2js = dm
End Function

'匹配SQL万能标签

Function loop_sql(Str)
    Set reg = New regexp
    reg.IgnoreCase = False
    reg.Global = True
    reg.Pattern = "{loop([\s\S.]*?)}([\s\S.]*?){loopend}"
    Set aa = reg.Execute(Str)
    For Each match in aa
        bq = match.submatches(0)
        nr = match.submatches(1)
        qsql = jiesql(bq, "sql")
        Str = Replace(Str, match, slb(qsql, nr))
    Next
    loop_sql = Str
End Function

'列表

Function slb(yj, mneirong)

    '正则替换
    Set reg = New regexp
    reg.IgnoreCase = False
    reg.Global = True
    reg.Pattern = "{fqesy:([a-zA-Z0-9]*?)}"
    Set aa = reg.Execute(mneirong)

    Dim kk()
    ReDim kk(aa.Count)
    '循环开始
    Set rs = server.CreateObject("adodb.recordset")
    Set rs.activeconnection = conn
    rs.cursortype = 3
    sql = ""&yj&""
    rs.Open(sql)

    kk(0) = mneirong
    Do While Not rs.EOF
        For i = 0 To aa.Count -1
            kk(i + 1) = Replace(kk(i), aa(i), rs(""&wnbq(aa(i))&""))
            y = y + kk(i)
        Next
        rs.movenext

    Loop
    '循环结束
    slb = y
End Function

'万能标签正则过滤

Function wnbq(all)
    Set ra = New RegExp
    ra.IgnoreCase = True
    ra.Global = True
    ra.Pattern = "{fqesy:(.*?)}"
    wnbq = ra.Replace(all, "$1")
End Function

'截取SQL语句需要

Function jiesql(xx, yy)
    Set reg = New regexp
    reg.IgnoreCase = False
    reg.Global = True
    reg.Pattern = ""&yy&"=""([\s\S.]*?)"""
    Set aa = reg.Execute(xx)
    For Each match in aa
        aa = match.submatches(0)
    Next
    jiesql = aa
End Function

'过滤大括号{}

Function dkh(d)
    d = Replace(d, "{", "")
    d = Replace(d, "}", "")
    dkh = d
End Function

'替换单引号

Function glyh(d)
    d = Replace(d, Chr(39), "&#39;")
    d = Replace(d, Chr(34), "&#34;")
    glyh = d
End Function

'添加关键字

Function keyword_add(Key)
    keyword = Split(Key, ",")
    For i = 0 To UBound(keyword)
        Set rs = server.CreateObject("adodb.recordset")
        Set rs.activeconnection = conn
        rs.cursortype = 3
        sql = "select * from tags where tag_name='"&Trim(keyword(i))&"'"
        rs.Open sql
        If rs.EOF Then
            sql = "insert into tags (tag_name,tag_num) values ('"&Trim(keyword(i))&"','1')"
            conn.Execute(sql)
            Set rsc = server.CreateObject("adodb.recordset")
            Set rsc.activeconnection = conn
            rsc.cursortype = 3
            sql = "select * from tags order by id desc"
            rsc.Open(sql)
            k = rsc("id")
        Else
            Set rsq = server.CreateObject("adodb.recordset")
            Set rsq.activeconnection = conn
            rsq.cursortype = 3
            sql = "select * from tags where tag_name='"&Trim(keyword(i))&"'"
            rsq.Open(sql)
            sql = "update tags set tag_num='"&rsq("tag_num") + 1&"' where id="&rsq("id")
            conn.Execute(sql)
            k = rsq("id")
        End If
        bb = bb + ",{"&k&"}"
    Next
    keyword_add = cut(2, bb)
End Function

'修改关键字

Function keyword_edit(Key)
    keyword = Split(Key, ",")
    For i = 0 To UBound(keyword)
        Set rs = server.CreateObject("adodb.recordset")
        Set rs.activeconnection = conn
        rs.cursortype = 3
        sql = "select * from tags where tag_name='"&Trim(keyword(i))&"'"
        rs.Open sql
        If rs.EOF Then
            sql = "insert into tags (tag_name,tag_num) values ('"&Trim(keyword(i))&"','1')"
            conn.Execute(sql)
            Set rsc = server.CreateObject("adodb.recordset")
            Set rsc.activeconnection = conn
            rsc.cursortype = 3
            sql = "select top 1 * from tags order by id desc"
            rsc.Open(sql)
            k = rsc("id")
        Else
            Set rsq = server.CreateObject("adodb.recordset")
            Set rsq.activeconnection = conn
            rsq.cursortype = 3
            sql = "select * from tags where tag_name='"&Trim(keyword(i))&"'"
            rsq.Open(sql)
            sql = "update tags set tag_num='"&rsq("tag_num")&"' where id="&rsq("id")
            conn.Execute(sql)
            k = rsq("id")
        End If
        bb = bb + ",{"&k&"}"
    Next
    keyword_edit = cut(2, bb)
End Function

sub checkUserlogin()
		if userName = "" then Tourl("./Login.asp")
		if password = "" then Tourl("./Login.asp")
		set u=new Users
		if userName <> "" or password <> "" then
			u.getuser(userName)
			if u.rs("password") <> password then die "密码错误！"
			if u.rs("userName")=userName and u.rs("password")=password then
				Response.cookies(CookieName)("UserName") = u.rs("userName")
				Response.cookies(CookieName)("password") = u.rs("password")
				Response.cookies(CookieName)("isUserLogin") = "True"
				isUserLogin = true
			end if
		else
			die "抱歉，您没有登录！<a href='javascript:history.go(-1);'>返回</a>"
		end If
end Sub
Function vbsEscape(str)    
    dim i,s,c,a    
    s=""    
    For i=1 to Len(str)    
        c=Mid(str,i,1)    
        a=ASCW(c)    
        If (a>=48 and a<=57) or (a>=65 and a<=90) or (a>=97 and a<=122) Then    
            s = s & c    
        ElseIf InStr("@*_+-./",c)>0 Then    
            s = s & c    
        ElseIf a>0 and a<16 Then    
            s = s & "%0" & Hex(a)    
        ElseIf a>=16 and a<256 Then    
            s = s & "%" & Hex(a)    
        Else    
            s = s & "%u" & Hex(a)    
        End If    
    Next    
    vbsEscape = s    
End Function  
  
Function vbsUnEscape(str)    
    dim i,s,c    
    s=""    
    For i=1 to Len(str)    
        c=Mid(str,i,1)    
        If Mid(str,i,2)="%u" and i<=Len(str)-5 Then    
            If IsNumeric("&H" & Mid(str,i+2,4)) Then    
                s = s & CHRW(CInt("&H" & Mid(str,i+2,4)))    
                i = i+5    
            Else    
                s = s & c    
            End If    
        ElseIf c="%" and i<=Len(str)-2 Then    
            If IsNumeric("&H" & Mid(str,i+1,2)) Then    
                s = s & CHRW(CInt("&H" & Mid(str,i+1,2)))    
                i = i+2    
            Else    
                s = s & c    
            End If    
        Else    
            s = s & c    
        End If    
    Next    
    vbsUnEscape = s    
End Function    
Function checkLoginBool(byval userName,byval password)
	Dim flag
	if userName = "" then flag = False
	if password = "" then flag = False
	set u=new Users
	u.getuser(userName)
	if not u.rs.eof then 
		'die u.rs("password")&"<br/>"&password
		If u.rs("password") = password And u.rs("userName")=userName Then 
			flag = True
		End If
	end if
	checkLoginBool = flag
	u.rs.close
	Set u.rs = Nothing
	Set u = Nothing
End Function

'显示错误函数

Sub xs_err(cz)
    If Err.Number <> 0 Then
        Response.Write cz&"<br>错误原因："&Err.Description
        Response.End
    End If
End Sub

Response.Charset=q_Charset
Sub echo(str)
	response.write(str)
	response.Flush()
End Sub

Sub die(str)
	if not isNul(str) then
		Response.write str
	end if	 
	response.End()
End Sub

Function isNul(str)
	if isnull(str) or str = ""  Then
	isNul = true 
	else 
	isNul = false 
	End if
End Function


%>