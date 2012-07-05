<!--#include file="../../inc/utf.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=<%=q_Charset%>" />
<style type="text/css">
<!--
body {
	margin-top: 0px;
	margin-bottom: 0px;
	background-color: #F1F4F7;
}
body,td,th {
	font-size: 12px;
	line-height: 30px;
}
-->
</style>
</head>
<body>
<%
'----------------------------------------------------------
'***************** 风声无组件上传类 2.1 *****************
'用法举例：快速应用[添加产品一]
'该例主要说明默认模式下的运用
'以常见的产品更新为例<br>
'该例以UTF-8字符集测试
'下面是上传程序(upload.asp)的代码和注释
'**********************************************************
'---------------------------------------------------------- 
Server.ScriptTimeOut=5000
%>
<!--#include file="../../inc/conn.asp"-->
<!--#include file="../isadmin.asp"-->
<!--#include file="../UpLoadClass.asp"-->
<%
'生成文件夹
MakeNewsDir(install&"upfile/")
dim request2

'建立上传对象
set request2=New UpLoadClass

	'设置字符集
	request2.Charset=q_Charset
	
	'保存路径
	request2.SavePath=install&"upfile/"

	'打开对象
	request2.Open()
	
	'显示目标文件路径与名称
	
	response.write "已上传成功！&nbsp;&nbsp;&nbsp;<a href='javascript:history.go(-1);'>重新上传</a>"
    Response.Write "<script>parent.uploadForm.url.value='"&request2.SavePath&request2.Form("strPhoto")&"'</script>"

'释放上传对象
set request2=nothing

'生成目录文件夹
Function MakeNewsDir(foldername)
     dim fso,f
	 On Error Resume Next
     Set fso = Server.CreateObject("Scripting.FileSystemObject")
     Set f = fso.CreateFolder(Server.MapPath(foldername))
     MakeNewsDir = True
     Set fso = nothing
End Function
%></body></html>
