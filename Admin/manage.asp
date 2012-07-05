<!--#include file="../inc/utf.asp"-->
<!--#include file="../inc/const.asp"-->
<html>
      <head>
      <SCRIPT LANGUAGE="JAVASCRIPT" TYPE="TEXT/JAVASCRIPT">
      <!--
      if(top.location != self.location)
      {
        top.location.replace(self.location)
      }
      -->
      </SCRIPT>
 
      <title>QCMS网站管理系统</title>
	  <meta http-equiv="Content-Type" content="text/html; charset=<%=q_Charset%>">
      </head>
 
      <frameset cols="185,*" framespacing="0" border="0" frameborder="0" frameborder="no">
        <frame src="menu.asp" name="leftFrame" scrolling="yes" frameborder="0" border="no" />
		<frame src="main.asp" name="mainFrame" scrolling="YES" />
	  </frameset>
 
      <noframes>
        <body>对不起，您的浏览器不支持框架，程序无法运行！</body>
      </noframes>
      </html>