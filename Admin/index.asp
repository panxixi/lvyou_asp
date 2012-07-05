<!--#include file="../inc/utf.asp"-->
<!--#include file="../inc/conn.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=<%=q_Charset%>" />
<title>QCMS后台管理</title>
<style type="text/css">
<!--
*{vertical-align:baseline;font-weight:inherit;font-family:inherit;font-style:inherit;font-size:100%;outline:0;border:0;margin:0;padding:0}
a{color:#333;font-size:12px}
.login{
	position:absolute;
	top:50%;
	left:50%;
	height:300px;
	width:350px;
	margin:-140px auto auto -180px;
	background-color: #FFF;
	border: 1px solid #7C7C7C;
}
.login_1{padding-left:44px;padding-right:44px}
#login_2 input{
	border:1px dotted #CCC;
	height:22px;
	width:80px;
	color:#666;
	font-size:12px;
	background-color:#E6E6E6
}
#login_2{margin-top:20px;width:255px}
H1{font-size:12px;color:#333;margin-top:8px;margin-bottom:3px}
.en1{font-size:10px;font-family:Verdana}
.input1{font-family:verdana;background-color:#eee;border-bottom:#FFF 1px solid;border-left:#CCC 1px solid;border-right:#FFF 1px solid;border-top:#CCC 1px solid;font-size:12px;height:22px}
.input1-bor{
	font-family:verdana;
	background-color:#F0F8FF;
	font-size:12px;
	border:1px solid #3CF;
	height:22px;
	line-height: 22px;
	text-indent: 5px;
}
</style>
<style type="text/css">
<!--
.en1{font-size:10px;font-family:Verdana}
.input1,.input11{
	font-family:verdana;
	background-color:#eee;
	border-bottom:#FFF 1px solid;
	border-left:#CCC 1px solid;
	border-right:#FFF 1px solid;
	border-top:#CCC 1px solid;
	font-size:12px;
	height:22px;
	line-height: 22px;
	text-indent: 5px;
}
body {
	background-color: #CCC;
}
-->
</style>
</head>
<body>
<div class="login">
  <div><img src="styles/advanced/images/login.gif" width="350" height="57" />
  <form method="post" action="login.asp" name="login">
    <div class="login_1">
      <h1><a>Name</a></h1>
      <input size=42 name="admin_name" class="input11" onblur="this.className='input1'" onfocus="this.className='input1-bor'" id="admin_name" />
      <h1><a>Pass</a></h1>
      <input name="admin_password" type="password" class="input11" id="admin_password" onFocus="this.className='input1-bor'" onBlur="this.className='input1'" size=42>
      <h1><a>Code</a></h1>
      <input size=42 name="admin_yzm" class="input11" onBlur="this.className='input1'" onFocus="this.className='input1-bor'" id="admin_yzm">
      <div id="login_2">
        <table width="100%" border="1">
          <tr>
            <td width="139" align="center"><input name="提交" type="submit" value="登陆" /></td>
            <td width="100"><input type="reset" value="重置" /></td>
          </tr>
        </table>
      </div>
    </div>
    </form>
    
  </div>
</div>
</body>
</html>
