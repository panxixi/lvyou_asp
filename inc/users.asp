
<%
Class Users
    Public userId
    Public userName
    Public password
    Public RealName
    Public sex
    Public phone
    Public idcard
    Public joinTime
    Public jifen
    Public uid
    Public rs '结果集
    Public QQHASH
    Private strsql 'SQL语句

    '获取用户列表

    Public Sub getUsers(byval rows,byval orderby)
        Set rs = server.CreateObject("adodb.recordset")
        Set rs.activeconnection = conn
        sql = "select top "&rows&" * from users "&orderby
        rs.Open sql
    End Sub

    '修改用户

    Public Sub updateUser()
        sql="update users set RealName='"&RealName&"',password='"&password&"',sex='"&sex&"',phone='"&phone&"',idcard='"&idcard&"',jifen="&jifen&" where userId="&userId
        conn.Execute(sql)
    End Sub
    
    Public Sub updateUserHt()
        sql="update users set [userName]='"&userName&"',[password]='"&password&"',[RealName]='"&RealName&"',[phone]='"&phone&"',[idcard]='"&idcard&"',[jifen]="&jifen&" where [userId]="&userId
        conn.Execute(sql)
    End Sub
    
    '修改资料
    Public Sub updateUserOption()
        sql="update users set RealName='"&RealName&"',sex='"&sex&"',phone='"&phone&"',idcard='"&idcard&"' where UserName='"&UserName&"'"
        conn.Execute(sql)
    End Sub
    
    '删除用户

    Public Sub deluser(userId)
        sql = "delete from users where userId="&userId
        conn.Execute(sql)
    End Sub

	Public Sub delusers(userId)
        sql = "delete from users where userId in ("&userId&")"
        conn.Execute(sql)
    End Sub
    '添加用户

    Public Sub adduser()
        'sql = "insert into users (userName,RealName,password,sex,phone,idcard,jifen,joinTime) values ('"&userName&"','"&RealName&"','"&password&"','"&sex&"','"&phone&"','"&idcard&"','"&jifen&"','"&joinTime&"')"
        sql = "Insert into users ([userName],[RealName],[password],[sex],[phone],[idcard],[jifen],[joinTime])values('"&userName&"','"&RealName&"','"&password&"','"&sex&"','"&phone&"','"&idcard&"',"&jifen&",'"&joinTime&"')"
        conn.Execute(sql)
        
    End Sub
    
    Public Sub addQQuser()
        'sql = "insert into users (userName,RealName,password,sex,phone,idcard,jifen,joinTime) values ('"&userName&"','"&RealName&"','"&password&"','"&sex&"','"&phone&"','"&idcard&"','"&jifen&"','"&joinTime&"')"
        sql = "Insert into users ([userName],[RealName],[password],[sex],[phone],[idcard],[jifen],[joinTime],[QQHASH])values('"&userName&"','"&RealName&"','"&password&"','"&sex&"','"&phone&"','"&idcard&"',"&jifen&",'"&joinTime&"','"&QQHASH&"')"
        conn.Execute(sql)
        
    End Sub

	Public Sub addusersId()
		sql = "insert into users (userName,RealName,password,sex,phone,idcard,jifen,joinTime,uid) values ('"&userName&"','"&RealName&"','"&password&"','"&sex&"','"&phone&"','"&idcard&"',"&jifen&",'"&joinTime&"','"&uid&"')"
        conn.Execute(sql)
	End Sub
	
	Public Sub adduserForbbs()
        sql = "Insert into users ([userName],[RealName],[password],[jifen],[joinTime])values('"&username&"','"&username&"','"&password&"',1,'"&now()&"')"
        conn.Execute(sql)
	End Sub

	'加积分操作
	Public Sub Jiajifen(byval num,byval userId)
        sql = "update users set jifen=jifen+"&num&" where userId="&userId
        conn.Execute(sql)
	end Sub
	
	Public Sub JiajifenByuserName(byval num,byval userName)
		sql = "update Users set jifen=jifen+"&cint(num)&" where userName='"&userName&"'"
		conn.Execute(sql)
	end Sub
	
	Public Sub JianjifenByuserName(byval num,byval userName)
		sql = "update Users set jifen=jifen-"&cint(num)&" where userName='"&userName&"'"
		conn.Execute(sql)
	end Sub
	
	'减积分操作
	Public Sub Jianjifen(byval num,byval userId)
        sql = "update users set jifen=jifen-"&num&" where userId="&userId
        conn.Execute(sql)
	end Sub
	
	
    '查看单个用户

    Public Sub getuser(userName)
        sql = "select * from users where userName='"&userName&"'"
        set rs = conn.execute(sql)
    End Sub
    
    Public Sub getuserById(id)
        sql = "select * from users where userId="&id
        set rs = conn.execute(sql)
    End Sub
	Public Sub getuserbyQQHASH(qqhash)
		sql = "select * from users where QQHASH='"&qqhash&"'"
        set rs = conn.execute(sql)
	End Sub
    '判断用户是否存在

    Public Function haveuser(userName)
        Set rs = server.CreateObject("adodb.recordset")
        Set rs.activeconnection = conn
        rs.cursortype = 3
        sql = "select * from users where userName='"&userName&"'"
        rs.Open sql
        If rs.EOF Then haveuser = true
    End Function
    
    '修改用户密码

    Public Sub edituserPass(userName)
        'sql = "update users set password='"&password&"' where userName='"&userName&"'"
        sql = "Update Users set [Password] = '"&password&"' where [UserName] = '"&userName&"'"
        conn.Execute(sql)
    End Sub
    
    Public Sub bindQQ(userName,qqHash)
        'sql = "update users set password='"&password&"' where userName='"&userName&"'"
        sql = "Update Users set [QQHASH] = '"&qqHash&"' where [UserName] = '"&userName&"'"
        conn.Execute(sql)
    End Sub

    '禁止站外提交

    Public Sub siteout()
        server_v1 = CStr(Request.ServerVariables("HTTP_REFERER"))
        server_v2 = CStr(Request.ServerVariables("SERVER_NAME"))
        If Mid(server_v1, 8, Len(server_v2))<>server_v2 Then
            response.Write "<br><br><center>"
            response.Write " "
            response.Write "你提交的路径有误，禁止从站点外部提交数据请不要乱改参数！"
            response.End
        End If
    End Sub

End Class
%>