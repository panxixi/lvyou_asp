<%
if Trim(Request.QueryString("Action"))="autotag" then
	title = Server.URLEncode(request("title"))
	Message = Server.URLEncode(request("Message"))
	'myxml="http://keyword.discuz.com/related_kw.html?title="&title&"&content="&Message&"&ics=utf-8&ocs=utf-8"
	myxml="http://search.qihoo.com/sint/related_kw.html?title="&title&"&content="&Message&"&ics=utf-8&ocs=utf-8"
	set xml = server.CreateObject("Microsoft.XMLDOM")
	xml.async = "false"
	xml.resolveExternals = "false"
	xml.setProperty "ServerHTTPRequest", true
	xml.load(myxml)
	If xml.getElementsByTagName("info")(0).selectSingleNode("count").Text > 0 Then
	Set objNodes = xml.getElementsByTagName("item") 
	For i = 0 to objNodes.length - 1 
	  showtag = showtag & Trim(objNodes(i).selectSingleNode("kw").Text)&","
	Next
	Response.write Left(showtag,Len(showtag)-1)
	Else
	Response.write "无法自动分词"
	End If

else
%>
<script>
//自动提取标签
function autotag() {
    var title = document.forms["form1"].huoDongName.value;
    //var Message = oEditor.EditorDocument.body.innerHTML;
    var Message = document.forms["form1"].huoDongXingcheng.value;
	//alert(Message);
	var Str="title=" + escape(title) + "&Message=" + escape(Message) + "&m=" + Math.random();
    document.forms["form1"].atuotag.disabled = true;
    if (title.length < 1) {
        alert("错误提示：没有填写标题！    ");
        document.forms["form1"].atuotag.disabled = false;
        return false
    }
    if (Message.length < 1) {
        alert("错误提示：内容不能为空！    ");
        document.forms["form1"].atuotag.disabled = false;
        return false
    }
	else {
		if (window.ActiveXObject) {
			xmlHttp = new ActiveXObject("Microsoft.XMLHTTP");
		}
		else if (window.XMLHttpRequest) {
			xmlHttp = new XMLHttpRequest();
		}
		var url = "autoTag.asp?Action=autotag";
		var query="title=" + escape(title) + "&Message=" + escape(Message) + "&m=" + Math.random();
		xmlHttp.onreadystatechange = function() {
			if(xmlHttp.readyState == 4) {
				if(xmlHttp.status == 200) {
					var TempStr = xmlHttp.responseText;
					document.forms["form1"].atuotag.disabled = false;
					if (TempStr == 0) {
						alert("没有找到可用标签！")
					} else {
						document.forms["form1"].nkeyword.value += TempStr
					}
				}
				else{
				alert("Ajax无返回！")
				}
			}
		}
		xmlHttp.open("POST", url, true);
 		xmlHttp.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
		xmlHttp.send(query);
	}
}

</script>

<%End if%>
<%
sub getTag()
	response.Write "<input name=""atuotag"" type=""button"" class=""ActionBtn Pointer"" value=""提取标签"" onclick=""autotag();""/>"
End sub
%>