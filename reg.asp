<!--#include file="inc/utf.asp"-->
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/s.asp"-->
<!--#include file="inc/news.asp"-->
<!--#include file="inc/category.asp"-->
<!--#include file="inc/sys.asp"-->
<!--#include file="inc/tfunction.asp"-->
<!--#include file="admin/md5.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=<%=q_Charset%>" />
<title>用户注册</title>
<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
	line-height: 20px;
}
table {
	border-top-width: 1px;
	border-left-width: 1px;
	border-top-style: solid;
	border-left-style: solid;
	border-top-color: #CCC;
	border-left-color: #CCC;
}
table tr td {
	border-right-width: 1px;
	border-bottom-width: 1px;
	border-right-style: solid;
	border-bottom-style: solid;
	border-right-color: #CCC;
	border-bottom-color: #CCC;
	padding-left: 10px;
}
-->
</style></head>
<body>

<script language="javascript">
function id_keyup(txtinput)
{
    txtinput.value=txtinput.value.replace(/(^\s*)|(\s*$)/g, "");
}
function form_onsubmit()
{
	if (document.form1.Username.value=="")
	{
	alert ("用户名不能为空!");
	document.form1.Username.focus();
	return false;
	}
	if (document.form1.Username.value.length <3 || document.form1.Username.value.length >12)
	{
	alert ("用户名限制在3-6位数组或字母组合！");
	document.form1.Username.focus();
	return false;
	}
	if (document.form1.Password.value=="")
	{
	alert ("密码不能为空!");
	document.form1.Password.focus();
	return false;
	}
	if (document.form1.Password.value.length <6 || document.form1.Password.value.length >15)
	{
	alert ("密码限制在6-15位数组或字母组合！");
	document.form1.Password.focus();
	return false;
	}
	if (document.form1.Passwordtwo.value=="")
	{
	alert ("确认密码不能为空!");
	document.form1.Passwordtwo.focus();
	return false;
	}
	if (document.form1.Password.value!=document.form1.Passwordtwo.value)
	{
	alert ("新密码和确认密码不一致。");
	document.form1.Password.value='';
	document.form1.Passwordtwo.value='';
	document.form1.Password.focus();
	return false;
}
if (document.form1.Mail.value.length == 0) {
		alert("EMAIL不能为空.");
		document.form1.Mail.focus();
		return false;
	}
	tChk = /^\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$/;
	if(!tChk.exec(document.form1.Mail.value)){
		alert("E-Mail格式不正确！");
		document.form1.Mail.focus();
		return false;
	}

	if (document.form1.Person.value=="")
	{
	alert ("真实姓名不能为空!");
	document.form1.Person.focus();
	return false;
	}
	if (document.form1.Person.value.length <2 || document.form1.Person.value.length >6)
	{
	alert ("真实姓名限制在2-6字！");
	document.form1.Username.focus();
	return false;
	}
    var tucktel = document.getElementById('Phone').value;
	if (tucktel=="")
	{
		alert ("电话不能为空!");
		document.form1.Phone.focus();
		return false;
	}else{
		var strFormat = "0123456789- ";
		var tel2="";
		for(var i=0;i<tucktel.length;i++){
		   if(strFormat.indexOf(tucktel.substr(i ,1)) == -1){
			  alert("你输入的电话号码格式不正确，请输入正确号码！以确保能正确联系！");
			  return false;
		   }
		}
		strFormat = "0123456789";
		for(i=0;i<tucktel.length;i++){
			if(strFormat.indexOf(tucktel.substr(i ,1)) != -1){
				tel2 = tel2 + tucktel.substr(i ,1);
			}	
		}
		tucktel = tel2;	
		var telnum=tucktel.length;
		if(telnum != 8 && telnum != 11  && telnum != 10 && telnum != 12 && telnum != 15 ){
			alert("你输入的电话号码不正确，请填写真实号码，以确保能正确联系！");
			return false;
		}
	}
	if (document.form1.Address.value=="")
	{
	alert ("地址不能为空!");
	document.form1.Address.focus();
	return false;
	}
}
</script>
<div class="container">
<form id="form1" name="form1" method="post" action="reg.asp?active=reg" onSubmit="return form_onsubmit()">
  <table width="60%" height="400" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td colspan="2" ><strong><font color="#FF0000">用户注册</font></strong></td>
              </tr>
              <tr>
                <td><span>用&nbsp;&nbsp;&nbsp;&nbsp;户</span></td>
                <td><input name="Username" type="text" class="inp" id="Username" maxlength="12" onFocus="this.className='admin_input_active';" onBlur="this.className='inp'" onkeyup="value=value.replace(/[\W]/g,'') " onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))" /> <span class="red">*</span><span>  3-12位英文字母与阿拉伯数字、区分大小写。</span></td>
              </tr>
              <tr>
                <td><span>密&nbsp;&nbsp;&nbsp;&nbsp;码</span></td>
                <td><input name="Password" type="password" class="inp" id="Password" maxlength="15" onFocus="this.className='admin_input_active';" onBlur="this.className='inp'" onkeyup="value=value.replace(/[\W]/g,'') " onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))" />                  <span class="red">*</span><span>  6-15个英文字母、数字，区分大小写。</span></td>
              </tr>
              <tr>
                <td><span>确认密码</span></td>
                <td><input name="Passwordtwo" type="password" class="inp" id="Passwordtwo" maxlength="15" onFocus="this.className='admin_input_active';" onBlur="this.className='inp'" onkeyup="value=value.replace(/[\W]/g,'') " onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))" />                  <span class="red">*</span><span>  请再输入一遍您上面输入的密码。</span></td>
              </tr>
      
                      <tr>
                <td><span>性&nbsp;&nbsp;&nbsp;&nbsp;别</span></td>
                <td>
                  <input name="sex" type="radio" id="radio" value="男" checked="checked" />
                  男<input type="radio" name="sex" id="radio" value="女" />
                  女
                </td>
              </tr>
              
              <tr>
                <td><span>QQ</span></td>
                <td>
                 <input name="qq" type="text" class="inp" id="qq" maxlength="30" />
                </td>
              </tr>
              
              <tr>
                <td><span>邮&nbsp;&nbsp;&nbsp;&nbsp;箱</span></td>
                <td><input name="Mail" type="text" class="inp" id="Mail" maxlength="30" onFocus="this.className='admin_input_active';" onBlur="this.className='inp'" onkeyup="id_keyup(this)"/>                  <span class="red">*</span> <font color="#91BE19">务必填写真实的邮件！</font></td>
              </tr>
              <tr>
                <td><span> 真实姓名</span></td>
                <td><input name="Person" type="text" class="inp" maxlength="4" onFocus="this.className='admin_input_active';" onBlur="this.className='inp'" onkeyup="value=value.replace(/[^\u4E00-\u9FA5]/g,'')" onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\u4E00-\u9FA5]/g,''))" />                  <span class="red">*</span> 为了方便与您联系 请填写真实中文姓名。</td>
              </tr>
              <tr>
                <td><span>电&nbsp;&nbsp;&nbsp;&nbsp;话</span></td>
                <td><input name="Phone" type="text" class="inp" onFocus="this.className='admin_input_active';" onBlur="this.className='inp'" onkeyup="id_keyup(this)"  />                  <span class="red">*</span> 为了方便与您联系 请填写有效电话。</td>
              </tr>
              <tr>
                <td><span>地&nbsp;&nbsp;&nbsp;&nbsp;址</span></td>
                <td><input name="Address" type="text" class="inp" maxlength="30" onFocus="this.className='admin_input_active';" onBlur="this.className='inp'"  onkeyup="id_keyup(this)" />                  <span class="red">*
                  <input name="act" type="hidden" id="act" value="reg" />
                </span></td>
              </tr>
    <tr>
                <td colspan="2" align="center">
                  
                  <input type="submit" name="button" id="button" value=" 提 交 " style="border:1px #000000 solid;vertical-align:middle;height:25px" /></td>
      </tr>
</table>
	      </form>
</body>
</html>
<%
act=request.QueryString("active")
if act="reg" Then
username=nohtml(request.Form("Username"))
userpwd=md5(nohtml(request.Form("Password")))
sex=nohtml(request.Form("sex"))
qq=nohtml(request.Form("qq"))
email=nohtml(request.Form("Mail"))
truename=nohtml(request.Form("Person"))
tel=nohtml(request.Form("Phone"))
address=nohtml(request.Form("Address"))
sql="insert into users(username,userpwd,sex,email,qq,tel,truename,address) values ('"&username&"','"&userpwd&"','"&sex&"','"&email&"','"&qq&"','"&tel&"','"&truename&"','"&address&"')"
conn.execute(sql)
xs_err("<font color='red'>注册失败！</font>")
jump=zx_url("注册成功!","index.asp")
%>

<%
end If
%>
