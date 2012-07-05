<%
'SQL防注入字符过滤
Dim GetFlag Rem(提交方式)
Dim ErrorSql Rem(非法字符) 
Dim RequestKey Rem(提交数据)
Dim ForI Rem(循环标记)
ErrorSql = "'~;~and~(~)~exec~insert~select~delete~update~count~*~%~chr~mid~master~truncate~char~declare" Rem(每个敏感字符或者词语请使用半角 "~" 格开)
ErrorSql = split(ErrorSql,"~")
If Request.ServerVariables("REQUEST_METHOD")="GET" Then
GetFlag=True
Else
GetFlag=False
End If
If GetFlag Then
For Each RequestKey In Request.QueryString
For ForI=0 To Ubound(ErrorSql)
If Instr(LCase(Request.QueryString(RequestKey)),ErrorSql(ForI))<>0 Then
die "<Script Language='javascript'>alert('包含非法字符，返回！');history.back(-1)</Script>"
End If
Next
Next 
Else
For Each RequestKey In Request.Form
For ForI=0 To Ubound(ErrorSql)
If Instr(LCase(Request.Form(RequestKey)),ErrorSql(ForI))<>0 Then
die "<Script Language='javascript'>alert('包含非法字符，返回！');history.back(-1)</Script>"
End If
Next
Next
End If
%>