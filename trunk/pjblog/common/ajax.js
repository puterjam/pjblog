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