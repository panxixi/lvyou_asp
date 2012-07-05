<!--#include file="../../inc/const.asp"-->
<%
dim path,myFile,read,write,cntNum
path=server.mappath("counter.qcms")
read=1
write=2
Set myFso = Server.CreateObject("Scripting.FileSystemObject")
set myFile = myFso.opentextfile(path,read)
cntNum=myFile.ReadLine
myFile.close
cntNum=FmtNumber(cntNum+1,6)
set myFile = myFso.opentextfile(path,write,TRUE)
myFile.write(cntNum)
myFile.close
set myFile=nothing
set myFso=nothing
'格式化数字
function FmtNumber(n,length)
 if len(n)>=length then
  fmtNumber = n
  exit function
 end if
 FmtNumber = string(length - len(n), "0") & n
end function
'替换成图片
function tihuan(cntNum)
cntNum=replace(cntNum,"1","<img src='"&install&"plus/count/img/1.gif'/>")
cntNum=replace(cntNum,"2","<img src='"&install&"plus/count/img/2.gif'/>")
cntNum=replace(cntNum,"3","<img src='"&install&"plus/count/img/3.gif'/>")
cntNum=replace(cntNum,"4","<img src='"&install&"plus/count/img/4.gif'/>")
cntNum=replace(cntNum,"5","<img src='"&install&"plus/count/img/5.gif'/>")
cntNum=replace(cntNum,"6","<img src='"&install&"plus/count/img/6.gif'/>")
cntNum=replace(cntNum,"7","<img src='"&install&"plus/count/img/7.gif'/>")
cntNum=replace(cntNum,"8","<img src='"&install&"plus/count/img/8.gif'/>")
cntNum=replace(cntNum,"9","<img src='"&install&"plus/count/img/9.gif'/>")
cntNum=replace(cntNum,"0","<img src='"&install&"plus/count/img/0.gif'/>")
tihuan=cntNum
end Function

function dm2js(dm)
dm=replace(dm,"""","\""")
dm=replace(dm,"'","\""")
dm=replace(dm,"/","\/")
dm= Replace(dm, CHR(13),"") 
dm= Replace(dm, CHR(10),"")
dm= Replace(dm, CHR(9),"")
dm2js=dm
end Function
%>
document.write('<%=dm2js(tihuan(cntNum))%>');