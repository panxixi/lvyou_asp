<%
userName = vbsUnEscape(Replace(Request.Cookies(CookieName)("UserName"),"'","''"))
password = Replace(Request.Cookies(CookieName)("password"),"'","''")
if not checkLoginBool(userName,password) then
	Tourl("./Login.asp")
end if
%>