// Geasion Media Network
function $(obj){
  return document.getElementById(obj);    
}
//---------------------------------------------------------- 自定义Event事件响应
function addEvent(obj, EventName, callBack){
	if(obj.addEventListener){ //FF
		obj.addEventListener(EventName, callBack, false);
	}else if(obj.attachEvent){ //IE
		obj.attachEvent('on' + EventName, callBack);
	}else{
		obj["on" + EventName] = callBack;
	}
}

//---------------------------------------------------------- 自定义Dom节点操作
function c$(na, id, cn){
	try{
		var em = document.createElement(na);
		if(id) em.id = id;
		if(cn) em.className = cn;
		return em;
	}catch(e){}
	return null;
}

//---------------------------------------------------------- Flash object
function fsso(src, width, height, fsn){
document.write('<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0" width="'+width+'" height="'+height+'">')
document.write('<param name="movie" value="'+src+'">')
document.write('<param name="Src" value="'+src+'">')
document.write('<param name="wmode" value="transparent">')
if(fsn==1){
document.write('<param name="allowFullScreen" value="true"/>')
}
document.write('<embed src="'+src+'" quality="high" width="'+width+'" height="'+height+'" TYPE="application/x-shockwave-flash" PLUGINSPAGE="http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash"></embed>');
document.write('</object>');
}


//---------------------------------------------------------- 内容翻滚
function domPlay(id,t,d,h,fq){ //id,间隔,方向,翻滚高度
	var box=$(id),can=true,t=t||1500,fq=fq||30,h=h||17,d=d==-1?-1:1;
	box.innerHTML+=box.innerHTML;
	box.onmouseover=function(){can=false};
	box.onmouseout=function(){can=true};
	var max=parseInt(box.scrollHeight/2);
	new function (){
		var stop=box.scrollTop%h==0&&!can;
		if(!stop){
			var set=d>0?[max,0]:[0,max];
			box.scrollTop==set[0]?box.scrollTop=set[1]:box.scrollTop+=d;
		};
		setTimeout(arguments.callee,box.scrollTop%h?fq:t);
	};
};

//---------------------------------------------------------- Tab菜单
var oo;
function chgTab(o){
	if(o==oo) return;
	oo=o;
	var op = o.parentNode;var opp = op.parentNode;var k=0;
	var opa= op.getElementsByTagName("a");var oppd = opp.getElementsByTagName("ul");
	for(var i=0,j=opa.length;i<j;i++){opa[i].className = "tab";oppd[i].style.display = "none"; if(o==opa[i])k=i;}
	o.className = "tab tab2";
	oppd[k].style.display = "block";
}

//---------------------------------------------------------- Image Error
function errImg(o){
	if(!o) o=window.event.srcElement;
	//alert(o.innerHTML);
	o.style.background = "url(images/err.jpg) no-repeat center center";
}
function setImg(){
	var is = document.getElementsByTagName("img");
	for(var i=0,j=is.length; i<j; i++){
		is[i].onerror = function(){errImg(this);}
		//addEvent(is[i],"error",errImg);
		//is[i].style.background = "url(images/err.jpg) no-repeat center center";
	}
}

