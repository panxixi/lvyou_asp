//QiAjax.js文件内容开始
/** 欢迎使用!
*程序名称:QiAjax
*文件名: QiAjax.js
*作者 :漆夜的空旷(QQ昵称)
*email:yaqy@qq.com
*主页: http://hi.baidu.com/tangtou

*说明:程序的使用方法:
例一:以GET的方式访问
   var ha = new QiAjax();
   ha.ajax_get("http://localhost:8085/", call_back);
   function call_back(text,xml)
   {
     alert(text);
   }
例二:以POST的方式访问
    var ha = new QiAjax();
    ha.ajax_post("post.php","name=yangQingrong&birth=1985",call_back2);
   function call_back2(text,xml)
   {
     alert(text);
   }
   回调函数call_back中的参数有两个,一个是返回的xmlHttp.responseText,另一个是xmlHttp.responseXml.
*/
function QiAjax()
{
   var ObjThis = this;
   // -----------------------------------------------------------
   this.getXmlHttp = function()
   {
      var xmlHttp;
      try
      {
         xmlHttp = new XMLHttpRequest;
      }
      catch(e)
      {
         try
         {
            xmlHttp = new ActiveXObject("MSXML2.XMLHTTP");
         }
         catch(e2)
         {
            try
            {
               xmlHttp = new ActiveXObject("Microsoft.XMLHTTP");
            }
            catch(e3)
            {
               xmlHttp = false;
            }
         }
      }
      return xmlHttp;
   }
   // -----------------------------------------------------------
   this.data_back = function()
   {
      if (this.xmlHttp.readyState == 4)
      {
         var response = this.xmlHttp.responseText;
         var xml=this.xmlHttp.responseXML;
         var func_callback=ObjThis.callback;
         func_callback(response,xml);
      }
   }
   // -----------------------------------------------------------
   this.xmlHttp = this.getXmlHttp();
   // -----------------------------------------------------------
   this.ajax_get = function(url, callback)
   {
      this.callback = callback;
      this.xmlHttp.open("GET", url, true);
      this.xmlHttp.onreadystatechange = function()
      {
         ObjThis.data_back();
      };
      this.xmlHttp.send(null);
   }
   // -----------------------------------------------------------
   this.ajax_post = function(url,param,callback)
   {
      this.callback = callback;
      this.xmlHttp.open("POST", url, true);
      this.xmlHttp.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
      this.xmlHttp.onreadystatechange = function()
      {
         ObjThis.data_back();
         
      };
      this.xmlHttp.send(param);
   }
   // -----------------------------------------------------------
}
function $(obj){
  return document.getElementById(obj);    
}

function $F(obj){
  return document.getElementById(obj).value;
}
function $U(str){
  return encodeURIComponent(escape(str));    
}
//QiAjax.js文件内容结束
