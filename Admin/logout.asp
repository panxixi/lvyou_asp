<!--#include file="../inc/conn.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=<%=q_Charset%>" />
</head>

<body>
<%
session("admin_name")=""
session("admin_password")=""
session("id")=""
%>
<SCRIPT LANGUAGE="javascript">
window.alert('成功退出!');window.self.location.href="index.asp"
</SCRIPT>
</body>
</html>