
<%
Class Baoming
    Public baoMingId
    Public baoMingHuoDong
    Public baoMingHuoDongId
    Public baomingLiuyan
    Public baoMingUserName
    Public baoMingRealName
    Public baoMingRenshu
    Public baoMingPhone
    Public baoMingIdcard
    Public baoMingZhuangtai '0 报名未删，1 报名已删 |由程序修改
    Public baoMingJifen
    Public baoMingAddTime
    Public baoMingUseJifen
    Public baoMingRen
    Public baoxianList
    Public location
    Public isBaoxian
    Public rs '结果集
    Private strsql 'SQL语句

'获取所有报名列表
    Public Sub allBaoMing()
        Set rs = server.CreateObject("adodb.recordset")
        Set rs.activeconnection = conn
        rs.cursortype = 3
        sql = "select Baoming.*,HuoDong.* from Baoming,HuoDong where Baoming.baoMingHuoDongId=HuoDong.huoDongId and baomingRenshu<>0 and baomingZhuangtai='0' order by baomingId desc"
        rs.Open(sql)
    End Sub
    
    '获取所有报名列表
    Public Sub HuishouBaoMing()
        Set rs = server.CreateObject("adodb.recordset")
        Set rs.activeconnection = conn
        rs.cursortype = 3
        sql = "select Baoming.*,HuoDong.* from Baoming,HuoDong where Baoming.baoMingHuoDongId=HuoDong.huoDongId and baomingRenshu<>0 and baomingZhuangtai='1' order by baomingId desc"
        rs.Open(sql)
    End Sub

    '根据活动 获取报名列表

    Public Sub BaoMingPageList(huoDongId)
        Set rs = server.CreateObject("adodb.recordset")
        Set rs.activeconnection = conn
        rs.cursortype = 3
        sql = "select Baoming.*,HuoDong.* from Baoming,HuoDong where Baoming.baoMingHuoDongId=HuoDong.huoDongId and HuoDong.huoDongId="&huoDongId&" and Baoming.baomingRenshu <> 0 and Baoming.baomingZhuangtai='0' order by Baoming.baoMingAddTime desc"
        'sql = "select * from Baoming where baominghuoDongId=13 and baomingRenshu <> 0 and baomingZhuangtai='0'"
        rs.Open(sql)
    End Sub
    
    Public sub SearchBaoming(searchkey)
        Set rs = server.CreateObject("adodb.recordset")
        Set rs.activeconnection = conn
        rs.cursortype = 3
        sql = "select Baoming.*,HuoDong.* from Baoming,HuoDong where Baoming.baomingUserName like '%"&searchkey&"%' or Baoming.baomingRealName like '%"&searchkey&"%' or Baoming.baomingPhone like '%"&searchkey&"%' or Baoming.baomingIdcard like '%"&searchkey&"%'"
        rs.Open(sql)
    End Sub
	
	Public Function getID(userName)
		Set rsa = Conn.execute("select userId from users where username='"&userName&"'")
		if not rsa.eof then
			getID=rsa("userId")
		else
			getID = 0
		end if
		rsa.close
		set rsa = nothing
	End Function


    '获取报名列表
    Public Sub getBaomings(byval rows,byval orderby)
        Set rs = server.CreateObject("adodb.recordset")
        Set rs.activeconnection = conn
        sql = "select top "&rows&" * from baoMing "&orderby
        rs.Open sql
    End Sub
    
    '获取此用户已经参加的活动列表
    Public sub getUserCanJiabaoMingList(byval baoMingUserName,byval typea,byval orderby)
        Set rs = server.CreateObject("adodb.recordset")
        Set rs.activeconnection = conn
        rs.cursortype = 3
    	sql = "select Baoming.*,HuoDong.* from Baoming,HuoDong where Baoming.baoMingHuoDongId=HuoDong.huoDongId and Baoming.baomingZhuangtai='0'"
'    	sql = "select * from baoMing where "
    	select case typea
    		case "0"'未参加的活动并且错过
    			'sql = sql & " and Baoming.baoMingUserName<>'"&baoMingUserName&"' and HuoDong.huodongzhuangtai='3'"
    			sql = "select * from huodong where huoDongId not in (select DISTINCT BaoMingHuoDongId from baoming where baomingUserName = '"&baoMingUserName&"') and huodongzhuangtai='3'"
    		case "1"'已经参加过的活动列表
    			sql = sql & " and Baoming.baoMingUserName='"&baoMingUserName&"'"
    		case "12"'获取当前未报满，已报满，已出发活动列表！
    			sql = sql & " and Baoming.baoMingUserName='"&baoMingUserName&"' and (HuoDong.huodongzhuangtai='0' or HuoDong.huodongzhuangtai='1' or HuoDong.huodongzhuangtai='2')"
    	end select
    	if typea <> "0" then
    		sql = sql &" and Baoming.baoMingRenShu <>0 order by Baoming.baoMingId desc"
    	end if
    	'die sql
        rs.Open sql
    end sub
    
    
        '获取此本次活动的操作列表
    Public sub getUserbaoMingList(byval baoMingUserName,byval huoDongId,byval orderby)
    	sql = "select Baoming.*,HuoDong.* from Baoming,HuoDong where Baoming.baoMingHuoDongId=HuoDong.huoDongId and Baoming.baomingZhuangtai='0'"
    	sql = sql & " and Baoming.baoMingUserName='"&baoMingUserName&"' and HuoDong.huoDongId="&huoDongId
    	sql = sql & orderby
    	set rs = conn.execute(sql)
    end sub
    
    Public sub getUserbaoMingSuccess(byval baoMingUserName,byval huoDongId)
    	sql = "select max(baomingId) from Baoming"
    	baomingId = conn.execute(sql)(0)
    	sql = "select Baoming.*,HuoDong.* from Baoming,HuoDong where Baoming.baoMingHuoDongId=HuoDong.huoDongId and Baoming.baomingZhuangtai='0'"
    	sql = sql & " and Baoming.baoMingUserName='"&baoMingUserName&"' and HuoDong.huoDongId="&huoDongId&" and Baoming.baomingId="&baomingId
    	sql = sql & orderby
    	set rs = conn.execute(sql)
    end sub
    
    Public Sub getUserBaomingSingle(byval baoMingUserName,byval baoMingId,byval huodongId)
	    sql = "select Baoming.*,HuoDong.* from Baoming,HuoDong where Baoming.baoMingHuoDongId=HuoDong.huoDongId and Baoming.baomingZhuangtai='0'"
    	sql = sql & " and Baoming.baoMingUserName='"&baoMingUserName&"' and HuoDong.huoDongId="&huoDongId&" and Baoming.baomingId="&baoMingId
    	set rs = conn.execute(sql)
	End Sub
    '修改报名

    Public Sub updatebaoMing()
        sql="update baoMing set baoMingUserName='"&baoMingUserName&"',baoMingRealName='"&baoMingRealName&"',baoMingRenshu="&baoMingRenshu&",baoMingPhone="&baoMingPhone&",baoMingIdcard='"&baoMingIdcard&"' where baomingId="&baoMingId
        conn.Execute(sql)
    End Sub
    
    Public Sub updatejiesuanBaoming(byval baoMingRenshu,byval baoMingId)
    	sql = "Update Baoming set baoMingRenshu="&baoMingRenshu&" where baomingId="&baoMingId
    	conn.execute(sql)
    End Sub
    
    '修改报名状态
    Public sub updatebaoMingzt(byval baoMingZhuangtai,byval baoMingId)
    	sql = "update baoming set baoMingZhuangtai='"&baoMingZhuangtai&"' where baoMingId="&baoMingId
    	conn.execute(sql)
    end sub
    
    '给此次报名加人
    public sub updateJiaren(byval num,byval baoMingId)
    	sql = "update baoming set baoMingRenshu=baoMingRenshu+"&num&" where baoMingId="&baoMingId
    	conn.execute(sql)
    end sub
    
    public sub updateJifen(byval num,byval baoMingId)
    	sql = "update baoming set baoMingJifen="&num&" where baoMingId="&baoMingId
    	conn.execute(sql)
    end sub
    
    public sub updateJianJifen(byval num,byval baoMingId)
    	sql = "update baoming set baoMingJifen="&num&" where baoMingId="&baoMingId
    	conn.execute(sql)
    end sub
    
    '给此次报名减人
    Public sub updateJianren(byval num,byval baoMingId)
    	sql = "update baoming set baoMingRenshu=baoMingRenshu-"&num&" where baoMingId="&baoMingId
    	conn.execute(sql)
    end sub
    
	
    '删除用户

    Public Sub delbaoMing(baoMingId)
        sql = "delete from baoming where baoMingId="&baoMingId
        conn.Execute(sql)
    End Sub


	Public Sub updateHuodongZhuangtai()
		sql = "update Huodong set huoDongZhuangtai=1 where huoDongId="&baoMingHuoDongId
		conn.execute(sql)
	End Sub

    '添加用户

    Public Sub addbaoMing()
    	sql = "update huoDong set huoDongShengyu=huoDongShengyu-"&baoMingRenshu&" where huoDongId="&baoMingHuoDongId
    	conn.execute(sql)
    	if isBaoxian <> "" then
    	else
    		isBaoxian = "0"
    	end if
        sql = "insert into baoMing (baoMingHuoDong,baoMingHuoDongId,baoMingUserName,baoMingRealName,baoMingRenshu,baoMingPhone,baoMingIdcard,baoMingZhuangtai,baoMingJifen,baoMingAddTime,baoMingUseJifen,baomingLiuyan,baoMingRen,location,isBaoxian,baoxianList) values ('"&baoMingHuoDong&"',"&baoMingHuoDongId&",'"&baoMingUserName&"','"&baoMingRealName&"',"&baoMingRenshu&",'"&baoMingPhone&"','"&baoMingIdcard&"','"&baoMingZhuangtai&"','"&baoMingJifen&"','"&baoMingAddTime&"',"&baoMingUseJifen&",'"&baomingLiuyan&"','"&baoMingRen&"','"&location&"','"&isBaoxian&"','"&baoxianList&"')"
        conn.Execute(sql)
    End Sub

    '查看单个报名

    Public Sub getbaoMing(baoMingId)
        sql = "select * from baoMing where baoMingId="&baoMingId
        set rs = conn.execute(sql)
    End Sub
    
    Public Function getAlljifen()
    	Set rsa = Conn.execute("select sum(baoMingJifen) from baoMing where baomingZhuangtai='0'")
		if not rsa.eof then
			getAlljifen=rsa(0)
		else
			getAlljifen = 0
		end if
		rsa.close
		set rsa = nothing
    End Function
    
    Public Function getAllIdcard()'查询所有填写真实身份证的用户 返回一个Rs数据集试试
    	set rsa = Conn.execute("select baomingIdcard from baoming where (((len(baomingIdcard)=15) and (mid(baomingIdcard,15,1) in (1,3,5,7,9)))or ((len(baomingIdcard)=15) and (mid(baomingIdcard,15,1) in (2,4,6,8,0))))or (((len(baomingIdcard)=18) and (mid(baomingIdcard,17,1) in (1,3,5,7,9)))or ((len(baomingIdcard)=18) and (mid(baomingIdcard,17,1) in (2,4,6,8,0))))")
      dim idcard
      Do while not rsa.eof
		idcard = idcard & rsa("baomingIdcard") & chr(10)
		rsa.movenext
	  Loop
	  getAllIdcard = idcard
		rsa.close
		set rsa = nothing
    End Function
    
    Public Function getAllphone()
    	set rsa = Conn.execute("select baoMingPhone from baoming where baomingZhuangtai='0'")
      dim idcard
      Do while not rsa.eof
		idcard = idcard & rsa("baoMingPhone") & chr(10)
		rsa.movenext
	  Loop
	  getAllphone = idcard
		rsa.close
		set rsa = nothing
    End Function
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