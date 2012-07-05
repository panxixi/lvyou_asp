
<%
Class CarLocation
    Public id
    Public location
    Public isShow
    Public rs '结果集
    Private strsql 'SQL语句

'获取所有报名列表
    Public Sub allCarLocation()
        Set rs = server.CreateObject("adodb.recordset")
        Set rs.activeconnection = conn
        rs.cursortype = 3
        sql = "select * from carlocation where isShow='0' order by id desc"
        rs.Open(sql)
    End Sub
    Public Sub getAllCarLocation()
        Set rs = server.CreateObject("adodb.recordset")
        Set rs.activeconnection = conn
        rs.cursortype = 3
        sql = "select * from carlocation order by id desc"
        rs.Open(sql)
    End Sub
    
	Public Sub getLoc(id)
        Set rs = server.CreateObject("adodb.recordset")
        Set rs.activeconnection = conn
        rs.cursortype = 3
        sql = "select * from carlocation where [id]="&id
        rs.Open(sql)
    End Sub
	Public Function getLocation(id)
		Set rsa = Conn.execute("select * from carlocation where [id]="&id)
		if not rsa.eof then
			location=rsa("location")
		else
			location = "无"
		end if
		rsa.close
		set rsa = nothing
		getLocation = location
	End Function


    Public Sub updateLocation()
        sql="update carlocation set location='"&location&"',isShow='"&isShow&"' where [id]="&id
        conn.Execute(sql)
    End Sub
    
    
	
    '删除用户

    Public Sub delcarLocation(id)
        sql = "delete from carLocation where [id]="&id
        conn.Execute(sql)
    End Sub

    Public Sub addCarLocation()
        sql = "insert into carlocation (location,isShow) values ('"&location&"','"&isShow&"')"
        conn.Execute(sql)
    End Sub


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