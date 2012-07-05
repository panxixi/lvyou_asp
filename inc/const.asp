<!--#include file="php_md5.asp"-->
<%
'QCMS安装目录
Const install="/"

'是否开启缓存
Const c_open="0"

'缓存时间（以分为单位）
Const huancun="1000"

'运行模式,ASP=0,html=1,rewrite=2
Const html="0"

'模版路径
Const temp_url="default"

'数据库名称
Const dbpath="#db.mdb"

'后台登陆验证码
Const yzcode="qcms"

'QCMS版本
Const Version="Qcms 1.4"

Const Message="true" '留言验证，如果是ture，可以直接留言，false则必须登陆才能留言

Const q_Charset="utf-8" '网站编码

'*************************************************
'路径规则
'*************************************************
'分类命名规则（当静态时，定义的是文件夹的名字，当伪静态时，定义的是文件名）
Const class_url="{pinyin}"

'文章命名规则
Const content_url="{pinyin}"

'上传大小
Const m_MaxSize  = "500000"

'上传格式
Const m_FileType = "jpg/gif"

'*************************************************
'数据库设置
'*************************************************
Const data_type="access" '如果是access数据库请写"access",如果是mssql请填写"mssql" 

Const mssql_ip="127.0.0.1" '数据库地址，本地可以写 localhost ,否则填写IP

Const mssql_data_name="qcms" '数据库名称

Const mssql_user="qcms" '数据库帐号

Const mssql_pwd="qcms" '数据库密码

CookieName="lvwo"

isUserLogin = false

MeirenciJia = 2

MeirenciJian = 3

ChushiJifen = 1

userName = vbsUnEscape(Replace(Request.Cookies(CookieName)("UserName"),"'","''"))
password = Replace(Request.cookies(CookieName)("password"),"'","''")
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

Public Function URLDecode(enStr)
dim deStr,strSpecial
dim c,i,v
  deStr=""
  strSpecial="!""#$%&'()*+,.-_/:;<=>?@[\]^`{|}~%"
  for i=1 to len(enStr)
    c=Mid(enStr,i,1)
    if c="%" then
      v=eval("&h"+Mid(enStr,i+1,2))
      if inStr(strSpecial,chr(v))>0 then
        deStr=deStr&chr(v)
        i=i+2
      else
        v=eval("&h"+ Mid(enStr,i+1,2) + Mid(enStr,i+4,2))
        deStr=deStr & chr(v)
        i=i+5
      end if
    else
      if c="+" then
        deStr=deStr&" "
      else
        deStr=deStr&c
      end if
    end if
  next
  URLDecode=deStr
End Function  
%>