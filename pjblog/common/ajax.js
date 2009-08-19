//Author:evio

//edit by evio start  AJAX 模式
//创建HTTPREQUEST对象系列
function echo(obj,html)
{
	$(obj).innerHTML=html;
}
function fopen(obj)
{
	$(obj).style.display="";
}
function fclose(obj)
{
	$(obj).style.display="none";
}
function CreateXMLHTTP()
{
	var xmlhttp=false;
	try	{
  		xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
 	} 
	catch (e) {
  		try {
   			xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
  		} 
		catch (e) {
   			xmlhttp = false;
 		}
 	}
	if (!xmlhttp && typeof XMLHttpRequest!='undefined') {
  		xmlhttp = new XMLHttpRequest();
				if (xmlhttp.overrideMimeType) {//设置MiME类别
			xmlhttp.overrideMimeType('text/xml');
		}
	}	

	return xmlhttp;	
}
function check(url,obj1,obj2)
{
		
		var xmlhttp = CreateXMLHTTP();
		if(!xmlhttp)
		{
			alert("你的浏览器不支持XMLHTTP！！");
			return;
		}
		xmlhttp.onreadystatechange=requestdata;
		xmlhttp.open("GET",url,true);
		xmlhttp.send(null);
		function requestdata()
		{
			
				fopen(obj1);
				echo(obj1,"<img src='images/loading.gif'>");
				if(xmlhttp.readyState==4)
				{
					if(xmlhttp.status==200)
					{
						if(obj1!=obj2){fclose(obj1);};
						echo(obj2,xmlhttp.responseText);
						
					}
				}
			
		}
}
//edit by evio start

//edit by 戒聊 start 
//判断验证码
var ajax_request_type = "GET";
var ajax_debug_mode = false;
function ajaxcheckcode(objid){
    var objname = getById(objid).value;
	if (objname){
		if (objname.length>=4)
		{
			ajaxcheckdata(objname);
		}			
	}
}

function initAjaxObject() {
	ajaxDebugPrint("initAjaxObject() called..");	
	var RetValue;
	try {
			RetValue = new ActiveXObject("Msxml2.XMLHTTP");
		} catch (e) {
		try {
		RetValue = new ActiveXObject("Microsoft.XMLHTTP");
		} catch (oc) {
		RetValue = null;
		}
	}
	if(!RetValue && typeof XMLHttpRequest != "undefined")
		RetValue = new XMLHttpRequest();
		if (!RetValue)
			ajaxDebugPrint("Could not create connection object.");
		return RetValue;
}

function getById(id) {
	itm = null;
	if (document.getElementById) {
		itm = document.getElementById(id);
	} else if (document.all)	{
		itm = document.all[id];
	} else if (document.layers) {
		itm = document.layers[id];
	}
	return itm;
}

function ajaxcheckdata(x_objname) {
	var xmlhttp = initAjaxObject();
	var post_data = null;
	var send_url = "common/checkcode.asp";
	try{
		if (ajax_request_type == "GET") {
			send_url = send_url + "?value="+ x_objname;
			post_data = null;
		}else{
			post_data = "value="+ x_objname;
		}
		xmlhttp.open(ajax_request_type, send_url.replace("&&","&"), true);
		if (ajax_request_type == "POST") {
			xmlhttp.setRequestHeader("Method", "POST " + send_url + " HTTP/1.1");
			xmlhttp.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
		}
		xmlhttp.onreadystatechange = function() {
			if (xmlhttp.readyState != 4 || xmlhttp.status != 200) 
				return;
			var response = xmlhttp.responseText;
			ajaxDebugPrint("received " + response.substring(0));
			var innerEl = getById('isok_checkcode');
			if(typeof(innerEl)=='object')
			innerEl.innerHTML = response;
		}
		xmlhttp.send(post_data);
		ajaxDebugPrint("url = " + send_url + "/post = " + post_data);
	}catch(e){}
	delete xmlhttp;
}

function ajaxDebugPrint(text) {
	if (ajax_debug_mode)
	alert("RSD: " + text);
}
//edit by 戒聊 end
