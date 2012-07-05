<!--#include file="../inc/utf.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="../inc/conn.asp"-->
<!--#include file="../inc/news.asp"-->
<!--#include file="../inc/sys.asp"-->
<!--#include file="../inc/category.asp"-->
<!--#include file="../inc/do.asp"-->
<!--#include file="isadmin.asp"--> 
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=<%=q_Charset%>" />
  <link rel=stylesheet href='styles/advanced/menu.css' />
  <script type="text/javascript">
  var cookieName = 'weenbizzMenuCookie';
 
  if(top.location == self.location)
  {
          self.location.replace("index.html")
  }
 
  function setCookie(Cookie, value, expiredays)
  {
          var ExpireDate = new Date ();
          ExpireDate.setTime(ExpireDate.getTime() + (expiredays * 24 * 3600 * 1000));
          document.cookie = Cookie + "=" + escape(value) +
          ((expiredays == null) ? "" : "; expires=" + ExpireDate.toGMTString());
  }
 
  function getCookie(Cookie)
  {
          if (document.cookie.length > 0)
          {
                  begin = document.cookie.indexOf(Cookie+"=");
                  if (begin != -1)
                  {
                          begin += Cookie.length+1;
                          end = document.cookie.indexOf(";", begin);
                          if (end == -1) end = document.cookie.length;
                          return unescape(document.cookie.substring(begin, end));
                  }
          }
          return null;
  }
 
  function LoadMenu()
  {
          cookieMenu = getCookie(cookieName);
          if(cookieMenu != null)
          {
                  for(i = 0; i < divNames.length; i++)
                  {
                          if(cookieMenu.indexOf(divNames[i]) >= 0)
                          {
                                  document.getElementById('div_' + divNames[i]).style.display = 'block';
                                  document.getElementById('img_' + divNames[i]).src = 'styles/advanced/images/closed.gif';
                          }
                          else
                          {
                                  document.getElementById('div_' + divNames[i]).style.display = 'none';
                                  document.getElementById('img_' + divNames[i]).src = 'styles/advanced/images/open.gif';
                          }
                  }
          }
  }
 
  function SaveMenu()
  {
          var cookiestring = '';
 
          for(i = 0; i < divNames.length; i++)
          {
                  var block = document.getElementById('div_' + divNames[i]);
 
                  if(block.style.display != 'none')
                  {
                          cookiestring += divNames[i] + '|';
                  }
          }
 
          setCookie(cookieName,cookiestring,1);
  }
 
  function ToggleDiv(className)
  {
          var img = document.getElementById('img_' + className);
          var div = document.getElementById('div_' + className);
 
          if(div.style.display == 'none')
          {
                  img.src = 'styles/advanced/images/closed.gif';
                  div.style.display = 'block';
          }
          else
          {
                  img.src = 'styles/advanced/images/open.gif';
                  div.style.display = 'none';
          }
  }
 
  function nav_goto(targeturl)
  {
          parent.frames.mainFrame.location = targeturl;
  }
  </script>
</head>
 
<body>
 
 
 
<table width="100%" border="0" cellpadding="0" cellspacing="0" align="center">
<tr>
  <td align="center"><a href="../" target="_blank" class="normal" onFocus="this.blur()"><img src="styles/advanced/images/logo.gif" border="0" /></a></td>
</tr>
<tr>
  <td align="center" style="padding-top: 8px; padding-bottom: 12px;">
    <a href="main.asp" target="mainFrame" class="normal">后台首页</a>&nbsp;|&nbsp;
    
    <a href="logout.asp" class="normal" target="_parent" onClick="return confirm('确定退出QCMS网站管理系统吗?');">安全退出</a>  </td>
</tr>
</table>
 
<table width="150" border="0" cellpadding="0" cellspacing="0" align="center">
<tr>
  <td valign="top" align="center">
 
  <table cellpadding="0" cellspacing="0" border="0" width="100%">
        <tr>
          <td bgcolor="#FFFFFF" style="padding-left: 2px; padding-top: 2px; padding-right: 2px; padding-bottom: 0px;"><div class="menutitle" onClick="ToggleDiv('系统管理');"><img id="img_系统管理" src="styles/advanced/images/closed.gif" border="0" align="absmiddle">&nbsp;&nbsp;系统管理</div>
        <div id="div_系统管理" style="display:block;"><div class="menulink-normal" onClick="nav_goto('sys.asp?id=1');" onMouseOver="this.className='menulink-hover';" onMouseOut="this.className='menulink-normal'"><img src="styles/advanced/images/menu_1.gif" border="0" align="absmiddle">&nbsp;<a href="sys.asp?id=1" target="mainFrame">基本设置</a></div><div class="menulink-normal" onClick="nav_goto('userAdmin.asp?id=2');" onMouseOver="this.className='menulink-hover';" onMouseOut="this.className='menulink-normal'"><img src="styles/advanced/images/menu_1.gif" border="0" align="absmiddle">&nbsp;<a href="userAdmin.asp?id=2" target="mainFrame">用户管理</a></div>
       <div class="menulink-normal" onClick="nav_goto('sys.asp?id=8');" onMouseOver="this.className='menulink-hover';" onMouseOut="this.className='menulink-normal'"><img src="styles/advanced/images/menu_1.gif" border="0" align="absmiddle">&nbsp;<a href="sys.asp?id=8" target="mainFrame">更新缓存</a></div><div class="menulink-normal" onClick="nav_goto('sys.asp?id=7');" onMouseOver="this.className='menulink-hover';" onMouseOut="this.className='menulink-normal'"><img src="styles/advanced/images/menu_1.gif" border="0" align="absmiddle">&nbsp;<a href="sys.asp?id=7" target="mainFrame">重置路径</a></div><div class="menulink-normal" onClick="nav_goto('sys.asp?id=3');" onMouseOver="this.className='menulink-hover';" onMouseOut="this.className='menulink-normal'"><img src="styles/advanced/images/menu_1.gif" border="0" align="absmiddle">&nbsp;<a href="sys.asp?id=3" target="mainFrame">管理员设置</a></div><div class="menulink-normal" onClick="nav_goto('sys.asp?id=4');" onMouseOver="this.className='menulink-hover';" onMouseOut="this.className='menulink-normal'"><img src="styles/advanced/images/menu_1.gif" border="0" align="absmiddle">&nbsp;<a href="sys.asp?id=4" target="mainFrame">修复数据库</a></div><div class="menulink-normal" onClick="nav_goto('sys.asp?id=88');" onMouseOver="this.className='menulink-hover';" onMouseOut="this.className='menulink-normal'"><img src="styles/advanced/images/menu_1.gif" border="0" align="absmiddle">&nbsp;<a href="sys.asp?id=88" target="mainFrame">参数设置</a></div>
	   <div class="menulink-normal" onClick="nav_goto('http://www.q-cms.cn/help/');" onMouseOver="this.className='menulink-hover';" onMouseOut="this.className='menulink-normal'"><img src="styles/advanced/images/menu_1.gif" border="0" align="absmiddle">&nbsp;<a href="http://www.q-cms.cn/help/" target="mainFrame">模版代码参考</a></div>
	   </div>
        
          </td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF" style="padding-left: 2px; padding-top: 2px; padding-right: 2px; padding-bottom: 0px;"><div class="menutitle" onClick="ToggleDiv('分类管理');"><img id="img_分类管理" src="styles/advanced/images/open.gif" border="0" align="absmiddle">&nbsp;&nbsp;分类管理</div>
        <div id="div_分类管理" style="display:none;"><div class="menulink-normal" onClick="nav_goto('Channel.asp?id=1');" onMouseOver="this.className='menulink-hover';" onMouseOut="this.className='menulink-normal'"><img src="styles/advanced/images/menu_1.gif" border="0" align="absmiddle">&nbsp;<a href="Channel.asp?id=1" target="mainFrame">分类管理</a></div><div class="menulink-normal" onClick="nav_goto('Channel.asp?id=2&cz=add');" onMouseOver="this.className='menulink-hover';" onMouseOut="this.className='menulink-normal'"><img src="styles/advanced/images/menu_1.gif" border="0" align="absmiddle">&nbsp;<a href="Channel.asp?id=2&cz=add" target="mainFrame">添加分类</a></div></div>
        
          </td>
        </tr>
        
        <tr>
          <td bgcolor="#FFFFFF" style="padding-left: 2px; padding-top: 2px; padding-right: 2px; padding-bottom: 0px;"><div class="menutitle" onClick="ToggleDiv('报名系统');"><img id="img_报名系统" src="styles/advanced/images/open.gif" border="0" align="absmiddle">&nbsp;&nbsp;报名系统</div>
	        <div id="div_报名系统" style="display:block;">
	        	<div class="menulink-normal" onClick="nav_goto('Huodong.asp');" onMouseOver="this.className='menulink-hover';" onMouseOut="this.className='menulink-normal'"><img src="styles/advanced/images/menu_1.gif" border="0" align="absmiddle">&nbsp;<a href="Huodong.asp" target="mainFrame">活动管理</a> &nbsp;&nbsp;<a href="Huodong.asp?Action=AddView" target="mainFrame">添加</a></div>
	        	<div class="menulink-normal" onClick="nav_goto('carlocation.asp?Action=ListView');" onMouseOver="this.className='menulink-hover';" onMouseOut="this.className='menulink-normal'"><img src="styles/advanced/images/menu_1.gif" border="0" align="absmiddle">&nbsp;<a href="carlocation.asp?Action=ListView" target="mainFrame">乘车地点</a></div>
	        	<div class="menulink-normal" onClick="nav_goto('Baoming.asp?ss=18');" onMouseOver="this.className='menulink-hover';" onMouseOut="this.className='menulink-normal'"><img src="styles/advanced/images/menu_1.gif" border="0" align="absmiddle">&nbsp;<a href="Baoming.asp" target="mainFrame">报名管理</a></div>
	        	
	        	<div class="menulink-normal" onClick="nav_goto('Jiesuan.asp');" onMouseOver="this.className='menulink-hover';" onMouseOut="this.className='menulink-normal'"><img src="styles/advanced/images/menu_1.gif" border="0" align="absmiddle">&nbsp;<a href="Jiesuan.asp" target="mainFrame">积分结算</a></div>
	        	
	        	<div class="menulink-normal" onClick="nav_goto('Huodong.asp?Action=AddView');" onMouseOver="this.className='menulink-hover';" onMouseOut="this.className='menulink-normal'"><img src="styles/advanced/images/menu_1.gif" border="0" align="absmiddle">&nbsp;<a href="Baoming.asp?Action=ChooseHuodong" target="mainFrame">帮会员报名</a></div>
	        	
	        	<div class="menulink-normal" onClick="nav_goto('Baoming.asp?ss=21');" onMouseOver="this.className='menulink-hover';" onMouseOut="this.className='menulink-normal'"><img src="styles/advanced/images/menu_1.gif" border="0" align="absmiddle">&nbsp;<a href="Baoming.asp?Action=HuishouView" target="mainFrame">报名回收站</a></div>
	        	
	        	<div class="menulink-normal" onClick="nav_goto('Logs.asp');" onMouseOver="this.className='menulink-hover';" onMouseOut="this.className='menulink-normal'"><img src="styles/advanced/images/menu_1.gif" border="0" align="absmiddle">&nbsp;<a href="Logs.asp" target="mainFrame">关键操作日志记录(未完)</a></div>
	        </div>
          </td>
        </tr>
        
        <tr>
          <td bgcolor="#FFFFFF" style="padding-left: 2px; padding-top: 2px; padding-right: 2px; padding-bottom: 0px;"><div class="menutitle" onClick="ToggleDiv('内容管理');"><img id="img_内容管理" src="styles/advanced/images/open.gif" border="0" align="absmiddle">&nbsp;&nbsp;内容管理</div>
        <div id="div_内容管理" style="display:none;"><div class="menulink-normal" onClick="nav_goto('content.asp?active=list');" onMouseOver="this.className='menulink-hover';" onMouseOut="this.className='menulink-normal'"><img src="styles/advanced/images/menu_1.gif" border="0" align="absmiddle">&nbsp;<a href="content.asp?active=list" target="mainFrame">内容管理</a></div><div class="menulink-normal" onClick="nav_goto('content.asp?active=add');" onMouseOver="this.className='menulink-hover';" onMouseOut="this.className='menulink-normal'"><img src="styles/advanced/images/menu_1.gif" border="0" align="absmiddle">&nbsp;<a href="content.asp?active=add" target="mainFrame">添加内容</a></div><div class="menulink-normal" onClick="nav_goto('content.asp?active=Replace');" onMouseOver="this.className='menulink-hover';" onMouseOut="this.className='menulink-normal'"><img src="styles/advanced/images/menu_1.gif" border="0" align="absmiddle">&nbsp;<a href="content.asp?active=Replace" target="mainFrame">批量替换</a></div><div class="menulink-normal" onClick="nav_goto('content.asp?active=hiRep');" onMouseOver="this.className='menulink-hover';" onMouseOut="this.className='menulink-normal'"><img src="styles/advanced/images/menu_1.gif" border="0" align="absmiddle">&nbsp;<a href="content.asp?active=hiRep" target="mainFrame">高级替换</a></div></div>
        
          </td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF" style="padding-left: 2px; padding-top: 2px; padding-right: 2px; padding-bottom: 0px;"><div class="menutitle" onClick="ToggleDiv('生成静态');"><img id="img_生成静态" src="styles/advanced/images/open.gif" border="0" align="absmiddle">&nbsp;&nbsp;生成静态</div>
        <div id="div_生成静态" style="display:none;"><div class="menulink-normal" onClick="nav_goto('content.asp?active=list');" onMouseOver="this.className='menulink-hover';" onMouseOut="this.className='menulink-normal'"><img src="styles/advanced/images/menu_1.gif" border="0" align="absmiddle">&nbsp;<a href="html.asp?active=index" target="mainFrame">更新主页</a></div><div class="menulink-normal" onClick="nav_goto('html.asp?active=list');" onMouseOver="this.className='menulink-hover';" onMouseOut="this.className='menulink-normal'"><img src="styles/advanced/images/menu_1.gif" border="0" align="absmiddle">&nbsp;<a href="html.asp?active=list" target="mainFrame">生成分类</a></div><div class="menulink-normal" onClick="nav_goto('html.asp?active=cont');" onMouseOver="this.className='menulink-hover';" onMouseOut="this.className='menulink-normal'"><img src="styles/advanced/images/menu_1.gif" border="0" align="absmiddle">&nbsp;<a href="html.asp?active=cont" target="mainFrame">生成内容</a></div><div class="menulink-normal" onClick="nav_goto('html.asp?active=google');" onMouseOver="this.className='menulink-hover';" onMouseOut="this.className='menulink-normal'"><img src="styles/advanced/images/menu_1.gif" border="0" align="absmiddle">&nbsp;<a href="html.asp?active=google" target="mainFrame">Google地图</a></div><div class="menulink-normal" onClick="nav_goto('html.asp?active=baidu');" onMouseOver="this.className='menulink-hover';" onMouseOut="this.className='menulink-normal'"><img src="styles/advanced/images/menu_1.gif" border="0" align="absmiddle">&nbsp;<a href="html.asp?active=baidu" target="mainFrame">Baidu地图</a></div><div class="menulink-normal" onClick="nav_goto('html.asp?active=rss');" onMouseOver="this.className='menulink-hover';" onMouseOut="this.className='menulink-normal'"><img src="styles/advanced/images/menu_1.gif" border="0" align="absmiddle">&nbsp;<a href="html.asp?active=rss" target="mainFrame">RSS地图</a></div></div>
        
          </td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF" style="padding-left: 2px; padding-top: 2px; padding-right: 2px; padding-bottom: 0px;"><div class="menutitle" onClick="ToggleDiv('其他设置');"><img id="img_其他设置" src="styles/advanced/images/open.gif" border="0" align="absmiddle">&nbsp;&nbsp;其他设置</div>
        <div id="div_其他设置" style="display:none;"><div class="menulink-normal" onClick="nav_goto('other.asp?open=review&act=list');" onMouseOver="this.className='menulink-hover';" onMouseOut="this.className='menulink-normal'"><img src="styles/advanced/images/menu_1.gif" border="0" align="absmiddle">&nbsp;<a href="other.asp?open=review&act=list" target="mainFrame">回复管理</a></div><div class="menulink-normal" onClick="nav_goto('other.asp?open=guest&act=list');" onMouseOver="this.className='menulink-hover';" onMouseOut="this.className='menulink-normal'"><img src="styles/advanced/images/menu_1.gif" border="0" align="absmiddle">&nbsp;<a href="other.asp?open=guest&act=list" target="mainFrame">留言管理</a></div><div class="menulink-normal" onClick="nav_goto('other.asp?open=bq&act=list');" onMouseOver="this.className='menulink-hover';" onMouseOut="this.className='menulink-normal'"><img src="styles/advanced/images/menu_1.gif" border="0" align="absmiddle">&nbsp;<a href="other.asp?open=bq&act=list" target="mainFrame">标签管理</a></div><div class="menulink-normal" onClick="nav_goto('other.asp?open=diy&act=list');" onMouseOver="this.className='menulink-hover';" onMouseOut="this.className='menulink-normal'"><img src="styles/advanced/images/menu_1.gif" border="0" align="absmiddle">&nbsp;<a href="other.asp?open=diy&act=list" target="mainFrame">自定义页面</a></div><div class="menulink-normal" onClick="nav_goto('other.asp?open=ilink&act=list');" onMouseOver="this.className='menulink-hover';" onMouseOut="this.className='menulink-normal'"><img src="styles/advanced/images/menu_1.gif" border="0" align="absmiddle">&nbsp;<a href="other.asp?open=ilink&act=list" target="mainFrame">站内连接</a></div><div class="menulink-normal" onClick="nav_goto('other.asp?open=js&act=list');" onMouseOver="this.className='menulink-hover';" onMouseOut="this.className='menulink-normal'"><img src="styles/advanced/images/menu_1.gif" border="0" align="absmiddle">&nbsp;<a href="other.asp?open=js&act=list" target="mainFrame">外部调用</a></div><div class="menulink-normal" onClick="nav_goto('other.asp?open=temp&act=list');" onMouseOver="this.className='menulink-hover';" onMouseOut="this.className='menulink-normal'"><img src="styles/advanced/images/menu_1.gif" border="0" align="absmiddle">&nbsp;<a href="other.asp?open=temp&act=list" target="mainFrame">模版管理</a></div><div class="menulink-normal" onClick="nav_goto('other.asp?open=plus&act=list');" onMouseOver="this.className='menulink-hover';" onMouseOut="this.className='menulink-normal'"><img src="styles/advanced/images/menu_1.gif" border="0" align="absmiddle">&nbsp;<a href="other.asp?open=plus&act=list" target="mainFrame">插件中心</a></div></div>
        
          </td>
        </tr>
        
        <tr>
          <td bgcolor="#FFFFFF" style="padding-left: 2px; padding-top: 2px; padding-right: 2px; padding-bottom: 0px;"><div class="menutitle" onClick="ToggleDiv('插件中心');"><img id="img_插件中心" src="styles/advanced/images/open.gif" border="0" align="absmiddle">&nbsp;&nbsp;插件中心</div>
        <div id="div_插件中心" style="display:none;">
        <%
		set cn=new config
		rs=cn.plus_list_on()
		do while not cn.rs.eof
		%>
        <div class="menulink-normal" onClick="nav_goto('<%=cn.rs("plus_manage")%>');" onMouseOver="this.className='menulink-hover';" onMouseOut="this.className='menulink-normal'"><img src="styles/advanced/images/menu_1.gif" border="0" align="absmiddle">&nbsp;<a href="<%=cn.rs("plus_manage")%>" target="mainFrame"><%=cn.rs("plus_name")%></a></div>
        
        <%
		cn.rs.movenext
		loop
		%>
        </div>
          </td>
        </tr>
        </table>
 
        <div style="margin-bottom:24px"></div>
  </td>
</tr>
</table>
<script type="text/javascript"> 
var divNames = new Array("系统管理");
window.onload = LoadMenu;
window.onbeforeunload = SaveMenu;
</script>
</body>
</html>
