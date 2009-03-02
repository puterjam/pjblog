/*------------------------------------------
Author: xujiwei
Website: http://www.xujiwei.cn
E-mail: vipxjw@163.com
Copyright (c) 2006, All Rights Reserved
------------------------------------------*/
function AJAXRequest() {
	var xmlObj = false;
	var CBfunc,ObjSelf;
	ObjSelf=this;
	try { xmlObj=new XMLHttpRequest; }
	catch(e) {
		try { xmlObj=new ActiveXObject("MSXML2.XMLHTTP"); }
		catch(e2) {
			try { xmlObj=new ActiveXObject("Microsoft.XMLHTTP"); }
			catch(e3) { xmlObj=false; }
		}
	}
	if (!xmlObj) return false;
	if(arguments[0]) this.url=arguments[0]; else this.url="";
	if(arguments[1]) this.callback=arguments[1]; else this.callback=function(obj){return};
	if(arguments[2]) this.content=arguments[2]; else this.content="";
	if(arguments[3]) this.method=arguments[3]; else this.method="POST";
	if(arguments[4]) this.async=arguments[4]; else this.async=true;
	this.send=function() {
		var purl,pcbf,pc,pm,pa;
		if(arguments[0]) purl=arguments[0]; else purl=this.url;
		if(arguments[1]) pc=arguments[1]; else pc=this.content;
		if(arguments[2]) pcbf=arguments[2]; else pcbf=this.callback;
		if(arguments[3]) pm=arguments[3]; else pm=this.method;
		if(arguments[4]) pa=arguments[4]; else pa=this.async;
		if(!pm||!purl||!pa) return false;
		xmlObj.open (pm, purl, pa);
		if(pm=="POST") xmlObj.setRequestHeader("Content-Type","application/x-www-form-urlencoded");
		xmlObj.onreadystatechange=function() {
			if(xmlObj.readyState==4) {
				if(xmlObj.status==200) {
					pcbf(xmlObj);
				}
				else {
					pcbf(null);
				}
			}
		}
		if(pm=="POST")
			xmlObj.send(pc);
		else
			xmlObj.send("");
	}
	this.get=function() {
		var purl,pcbf;
		if(arguments[0]) purl=arguments[0]; else purl=this.url;
		if(arguments[1]) pcbf=arguments[1]; else pcbf=this.callback;
		if(!purl&&!pcbf) return false;
		this.send(purl,"",pcbf,"GET",true);
	}
	this.post=function() {
		var fo,pcbf,purl,pc,pm;
		if(arguments[0]) fo=arguments[0]; else return false;
		if(arguments[1]) pcbf=arguments[1]; else pcbf=this.callback;
		if(arguments[2])
			purl=arguments[2];
		else if(fo.action)
			purl=fo.action;
		else
			purl=this.url;
		if(arguments[3])
			pm=arguments[3];
		else if(fo.method)
			pm=fo.method.toLowerCase();
		else
			pm="post";
		if(!pcbf&&!purl) return false;
		pc=this.formToStr(fo);
		if(!pc) return false;
		if(pm) {
			if(pm=="post")
				this.send(purl,pc,pcbf,"POST",true);
			else
				if(purl.indexOf("?")>0)
					this.send(purl+"&"+pc,"",pcbf,"GET",true);
				else
					this.send(purl+"?"+pc,"",pcbf,"GET",true);
		}
		else
			this.send(purl,pc,pcbf,"POST",true);
	}
	// formToStr
	// from SurfChen <surfchen@gmail.com>
	// @url     http://www.surfchen.org/
	// @license http://www.gnu.org/licenses/gpl.html GPL
	// modified by xujiwei
	// @url     http://www.xujiwei.cn/
	this.formToStr=function(fc) {
		var i,query_string="",and="";
		for(i=0;i<fc.length;i++) {
			e=fc[i];
			if (e.name!='') {
				if (e.type=='select-one') {
					element_value=e.options[e.selectedIndex].value;
				}
				else if (e.type=='checkbox' || e.type=='radio') {
					if (e.checked==false) {
						continue;	
					}
					element_value=e.value;
				}
				else {
					element_value=e.value;
				}
				element_value=encodeURIComponent(element_value);
				query_string+=and+e.name+'='+element_value;
				and="&";
			}
		}
		return query_string;
	}
}