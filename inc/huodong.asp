
<%
Class HuoDong
    Public huoDongId
    Public bbsLink
    Public huoDongName
    Public huoDongRenShu
    Public huoDongShengyu'由程序修改
    Public huoDongXingcheng
    Public huoDongxieyi
    Public huoDongLingdui
    Public huoDongJifen
    Public huoDongZhuangtai '0 报名未满，1 报名已满，2 活动出发, 3 活动结束 |由程序修改
    Public huoDongChufaTime
    Public huoDongFanhuiTime
    Public huoDongAddTime
    Public nkeyword
    Public huoDongInfo
    Public rs '结果集
	Public rs2
    Private strsql 'SQL语句

    Public Sub getHuoDongs(byval rows,byval orderby)
        Set rs = server.CreateObject("adodb.recordset")
        Set rs.activeconnection = conn
        rs.cursortype = 3
        sql = "select top "&rows&" * from huoDong where huoDongZhuangtai='0' or huoDongZhuangtai='1' "&orderby
        rs.Open sql
    End Sub

    Public Sub getAllHuoDong()
        Set rs = server.CreateObject("adodb.recordset")
        Set rs.activeconnection = conn
        rs.cursortype = 3
        sql = "select * from huoDong order by huodongId desc"
        rs.Open sql
    End Sub

	Public Sub getAutoZhuangtaiList()
        sql = "select * from huoDong where huoDongZhuangtai='1' or huoDongZhuangtai='2' order by huoDongId desc"
        Set rs2 = conn.execute(sql)
    End Sub
    
    Public Sub getJiesuanList()
        sql = "select * from huoDong where huodongJiesuan=0 order by huoDongId desc"
        Set rs = conn.execute(sql)
    End Sub
	Public Sub getSelecthuodong()
		sql = "select * from huodong where huodongZhuangtai='0' or huodongzhuangtai='1' order by huodongId desc"
		set rs = conn.execute(sql)
	End sub

    '修改用户

    Public Sub updateHuodong()
        sql="update huoDong set  bbsLink='"&bbsLink&"',huoDongName='"&huoDongName&"',huoDongRenShu="&huoDongRenShu&",huoDongXingcheng='"&huoDongXingcheng&"',huoDongxieyi='"&huoDongxieyi&"',huoDongLingdui='"&huoDongLingdui&"',huoDongZhuangtai='"&huoDongZhuangtai&"',huoDongJifen="&huoDongJifen&",huoDongChufaTime='"&huoDongChufaTime&"',huoDongFanhuiTime='"&huoDongFanhuiTime&"',nkeyword='"&nkeyword&"',huoDongInfo='"&huoDongInfo&"' where huoDongId="&huoDongId
        conn.Execute(sql)
    End Sub
    
    '修改活动状态
    Public Sub updateHuodongzhuangtai(byval zhuangtai,byval huoDongId)
    	sql = "update huoDong set "
    	select case zhuangtai
    		case "0"
    			sql = sql &"huoDongZhuangtai=0"
			case "1"
				sql = sql &"huoDongZhuangtai=1"
			case "2"
				sql = sql &"huoDongZhuangtai=2"
			case "3"
				sql = sql &"huoDongZhuangtai=3"
    	end select
    	sql = sql & " where huoDongId="&huoDongId
    	conn.execute(sql)
    end sub
    
    '加剩余人数-便于后台取消活动时添加剩余数量
    Public sub jiaShengyu(byval num,byval huoDongId)
    	sql = "update huoDong set huoDongShengyu=huoDongShengyu+"&num&" where huoDongId="&huoDongId
    	conn.execute(sql)
    	sql = "select * from huoDong where huoDongId="&huoDongId
    	set rs = conn.execute(sql)
    	if not rs.eof then
    		if rs("huoDongShengyu") > 0 then
    			call updateHuodongzhuangtai("0",huoDongId)
    		end if
    	end if
    end sub
    
    '减剩余人数-便于前台报名用户
    Public sub JianShengyu(byval num,byval huoDongId)
    	sql = "update huoDong set huoDongShengyu=huoDongShengyu-"&num&" where huoDongId="&huoDongId
    	conn.execute(sql)
    	sql = "select * from huoDong where huoDongId="&huoDongId
    	set rs = conn.execute(sql)
    	if not rs.eof then
    		if rs("huoDongShengyu") = 0 then
    			call updateHuodongzhuangtai("1",huoDongId)
    		end if
    	end if
	end sub
	
	Public sub updateRenshu(byval num,byval huoDongId)
    	sql = "update huoDong set huoDongRenShu="&num&" where huoDongId="&huoDongId
    	conn.execute(sql)
	end sub
	
    '删除用户

	Public sub updateJiesuan(byval huodongId)
		sql = "update Huodong set huodongJiesuan=1 where huoDongId="&huodongId
		conn.Execute(sql)
	End Sub

    Public Sub delhuoDong(huoDongId)
        sql = "delete from huoDong where huoDongId="&huoDongId
        conn.Execute(sql)
    End Sub

    '添加用户

    Public Sub addhuoDong()
        sql = "insert into huoDong (bbsLink,huoDongName,huoDongRenShu,huoDongShengyu,huoDongXingcheng,huoDongxieyi,huoDongLingdui,huoDongZhuangtai,huoDongJifen,huoDongChufaTime,huoDongFanhuiTime,huoDongAddTime,nkeyword,huoDongInfo) values ('"&bbsLink&"','"&huoDongName&"','"&huoDongRenShu&"',"&huoDongShengyu&",'"&huoDongXingcheng&"','"&huoDongxieyi&"','"&huoDongLingdui&"','0','"&huoDongJifen&"','"&huoDongChufaTime&"','"&huoDongFanhuiTime&"','"&huoDongAddTime&"','"&nkeyword&"','"&huoDongInfo&"')"
        conn.Execute(sql)
    End Sub

    '查看单个活动

    Public Sub gethuoDong(huoDongId)
        Set rs = server.CreateObject("adodb.recordset")
        Set rs.activeconnection = conn
        sql = "select * from huoDong where huoDongId="&huoDongId
        rs.Open sql
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