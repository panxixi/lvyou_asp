<!--#include file="../../inc/utf.asp"-->
<!--#include file="../../inc/const.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=<%=q_Charset%>" />
<title>无标题文档</title>
<style type="text/css">
<!--
body {
	margin-top: 5px;
	margin-bottom: 0px;
	margin-left: 5px;
	margin-right: 0px;
	background-color: #F1F4F7;
}
.upload {
	border: 1px solid #999;
	margin: 0px;
	padding: 0px;
}
.upk {
	margin: 0px;
	padding: 0px;
	float: left;
	height: 24px;
	width: 600px;
}
-->
</style>
</head>

<body>
	<form action="upload.asp" enctype="multipart/form-data" name="form1" method="post" class="upk">
      <input name="strPhoto" type="file" class="upload" id="strPhoto" size="20" />
            <input type="submit" name="Submit" value="提  交" style="border:1px #999999 solid;vertical-align:middle;height:18px" />
	</form>
</body>
</html>
