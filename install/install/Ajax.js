/*--------------------------------------
Name: AJAXRequest
Version: 0.8.06
Author: xujiwei
E-mail: vipxjw@163.com
Website: http://www.xujiwei.cn/
License: http://www.gnu.org/licenses/gpl.html GPL
Copyright (c) 2006-2008, xujiwei, All Rights Reserved

AjaxRequest Deveoper Manual:
    http://ajax.xujiwei.cn/
--------------------------------------*/
/**
 * AJAXRequest类
 * @author	xujiwei
 * @constructor
 * @class		AJAXRequest
 * @param		{Object}	[initObj]	初始化参数对象
 * @param		{String}	[initObj.url=""]			要请求的Url
 * @param		{String}	[initObj.content=""]		要发送的内容
 * @param		{String}	[initObj.method="GET"]		请求方法，GET 或 POST
 * @param		{String}	[initObj.charset]			发送数据时使用的编码
 * @param		{Boolean}	[initObj.async=true]		请求类型，true 为异步，false 为同步
 * @param		{Number}	[initObj.timeout=3600000]	请求超时时间，单位为毫秒
 * @param		{function}	[initObj.ontimeout]			请求超时时的回调函数
 * @param		{function}	[initObj.onrequeststart]		请求开始时的回调函数
 * @param		{function}	[initObj.onrequestend]		请求结束时的回调函数
 * @param		{function}	[initObj.oncomplete]		请求正确完成时的回调函数
 * @param		{function}	[initObj.onexception]		请求发生异常时的回调函数
 * @property	{String}	url			要请求的Url
 * @property	{String}	content		要发送的内容
 * @property	{String}	method		请求方法，GET 或 POST
 * @property	{String}	charset		发送数据时使用的编码
 * @property	{Boolean}	async		请求类型，true 为异步，false 为同步
 * @property	{Number}	timeout		请求超时时间，单位为毫秒
 * @property	{function}	ontimeout	请求超时时的回调函数
 * @property	{function}	onrequeststart	请求开始时的回调函数
 * @property	{function}	onrequestend	请求结束时的回调函数
 * @property	{function}	oncomplete	请求正确完成时的回调函数
 * @property	{function}	onexception	请求发生异常时的回调函数
 * @example
 * var ajax1 = new AJAXRequest();
 * var ajax2 = new AJAXRequest({
 * 	url: "getdata.asp",	// 从getdata.asp获取数据
 * 	method: "GET",		// GET方式
 * 	oncomplete: function(obj) {
 * 		alert(obj.responesText);	// 显示getdata.asp输出的内容
 * });
 */
function Ajax() {
	var xmlPool=[],objPool=[],AJAX=this,ac=arguments.length,av=arguments;
	var xmlVersion=["MSXML2.XMLHTTP","Microsoft.XMLHTTP"];
	var ec=emptyFun=function(){};
	av=ac>0?typeof(av[0])=="object"?av[0]:{}:{};
	_av=av;
	var encode=$GEC(av.charset+"");
	var prop=['url','content','method','async','timeout','ontimeout','onrequeststart','onrequestend','oncomplete','onexception'];
	var defval=['','','GET',true,3600000,ec,ec,ec,ec,ec],l=prop.length;
	while(l--){this[prop[l]]=getp(av[prop[l]],defval[l]);}
	if(!getObj()) return false;
	function getp(p,d) { return p!=undefined?p:d; }
	function getObj() {
		var i,j,tmpObj;
		for(i=0,j=xmlPool.length;i<j;i++) if(xmlPool[i].readyState==0||xmlPool[i].readyState==4) return xmlPool[i];
		try { tmpObj=new XMLHttpRequest; }
		catch(e) {
			for(i=0,j=xmlVersion.length;i<j;i++) {
				try { tmpObj=new ActiveXObject(xmlVersion[i]); } catch(e2) { continue; }
				break;
			}
		}
		if(!tmpObj) return false;
		else { xmlPool[xmlPool.length]=tmpObj; return xmlPool[xmlPool.length-1]; }
	}
	function $(id){return document.getElementById(id);}
	function $N(d){var n=d*1;return(isNaN(n)?0:n);}
	function $VO(v){return(typeof(v)=="string"?(v=$(v))?v:false:v);}
	function $GID(){return((new Date)*1);}
	function $SOP(id,ct){objPool[id+""]=ct;}
	function $LOP(id){return(objPool[id+""]);}
	function $SRP(f,r,p){return(function(s){s=f(s);for(var i=0;i<r.length;i++) s=s.replace(r[i],p[i]);return(s);});}
	function $GEC(cs){
		if(cs.toUpperCase()=="UTF-8") return(encodeURIComponent);
		else return($SRP(escape,[/\+/g],["%2B"]));
	}
	function $ST(obj,text) {
		var nn=obj.nodeName.toUpperCase();
		if("INPUT|TEXTAREA|OPTION".indexOf(nn)>-1) obj.value=text;
		else try{obj.innerHTML=text;} catch(e){};
	}
	function $CB(cb) {
		if(typeof(cb)=="function") return cb;
		else {
			cb=$VO(cb);
			if(cb) return(function(obj){$ST(cb,obj.responseText);});
			else return this.oncomplete; }
	}
	function $GP(p,v,d,f) {
		var i=0;
		while(i<v.length){p[i]=v[i]?f[i]?f[i](v[i]):v[i]:d[i];i++;}
		while(i<d.length){p[i]=d[i];i++;}
	}
	function send(purl,pc,pcbf,pm,pa) {
		var ct,ctf=false,xmlObj=getObj(),ac=arguments.length,av=arguments;
		if(!xmlObj) return false;
		var pmp=pm.toUpperCase()=="POST"?true:false;
		if(!pm||!purl) return false;
		var ev={url:purl, content:pc, method:pm};
		purl+=(purl.indexOf("?")>-1?"&":"?")+"timestamp="+$GID();
		xmlObj.open(pm,purl,pa);
		AJAX.onrequeststart(ev);
		if(pmp) xmlObj.setRequestHeader("Content-Type","application/x-www-form-urlencoded");
		ct=setTimeout(function(){ctf=true;xmlObj.abort();},AJAX.timeout);
		var rc=function() {
			if(ctf) { AJAX.ontimeout(ev); AJAX.onrequestend(ev); }
			else if(xmlObj.readyState==4) {
				ev.status=xmlObj.status;
				try{ clearTimeout(ct); } catch(e) {};
				try{ 
					if(xmlObj.status==200){
						pcbf(xmlObj); 
					}else {
						AJAX.onexception(ev);
					}
				}catch(e) {
					AJAX.onexception(ev); 
				}
				AJAX.onrequestend(ev);
			}
		}
		xmlObj.onreadystatechange=rc;
		if(pmp) xmlObj.send(pc); else xmlObj.send("");
		if(pa==false) rc();
		return true;
	}
	/**
	 * 设置发送请求时使用的编码
	 * @param	{String}	charset	编码名称，UTF-8 或 GB2312，或者其他
	 * @example
	 * var ajax = new AJAXRequest();
	 * ajax.setcharset("GB2312");
	 */
	this.setcharset=function(cs) { encode=$GEC(cs); }
	/**
	 * 使用GET方法请求指定的Url
	 * @param	{String}	[url]		要请求的Url
	 * @param	{function|Object|String}	[oncomplete]	正确返回时的回调函数，或者要更新的对象，或者要更新对象的ID
	 * @example
	 * var ajax = new AJAXRequest();
	 * ajax.get("getdata.asp", function(obj) {
	 * 	alert(obj.responseText);	// 显示从getdata.asp得到的数据
	 * });
	 * ajax.get("getdata.asp", "txtData");	// 将从getdata.asp得到的数据更新到ID为txtData的HTML元素
	 */
	this.get=function() {
		var p=[],av=arguments;
		$GP(p,av,[this.url,this.oncomplete],[null,$CB]);
		if(!p[0]&&!p[1]) return false;
		return(send(p[0],"",p[1],"GET",this.async));
	}
	/**
	 * 按一定时间间隔请求指定的Url并更新指定的对象
	 * @param	{function|Object|String}	oncomplete	正确返回时的回调函数，或者要更新的对象，或者要更新对象的ID
	 * @param	{String}	url			请求的Url
	 * @param	{Number}	interval	发送请求的时间间隔
	 * @param	{Number}	times		发送请求的次数
	 * @return	{String}	update过程的标识符，用于停止update
	 * @see		AJAXRequest#stopupdate
	 * @example
	 * var ajax = new AJAXRequest();
	 * ajax.update(function(obj) {
	 * 		alert(obj.responseText); 
	 * 	},
	 * 	"getdata.asp",	// 从getdata.asp获取数据
	 * 	1000,	// 每1秒更新一将
	 * 	3		// 更新3次
	 * );
	 */
	this.update=function() {
		var p=[],purl,puo,pinv,pcnt,av=arguments;
		$GP(p,av,[this.oncomplete,this.url,-1,-1],[$CB,null,$N,$N]);
		if(p[2]==-1) p[3]=1;
		var sf=function(){send(p[1],"",p[0],"GET",AJAX.async);};
		var id=$GID();
		var cf=function(cc) {
			sf(); cc--; if(cc==0) return;
			$SOP(id,setTimeout(function(){cf(cc);},p[2]));
		}
		cf(p[3]);
		return id;
	}
	/**
	 * 停止更新对象
	 * @param	{String}	update_id	update方法返回的标识符
	 * @see		AJAXRequest#update
	 * @example
	 * var ajax = new AJAXRequest();
	 * var up = ajax.update("txtData", "getdata.asp");
	 * ajax.stopupdate(up);
	 */
	this.stopupdate=function(id) {
		clearTimeout($LOP(id));
	}
	/**
	 * 发送数据到指定Url
	 * @param	{String}	[url]			请求的Url
	 * @param	{String}	[content]		要发送的数据
	 * @param	{function|Object|String}	[oncomplete]	正确返回时的回调函数，或者要更新的对象，或者要更新对象的ID
	 * @see		AJAXRequest#postf
	 * @example
	 * var ajax = new AJAXRequest();
	 * ajax.post("postdata.asp", "the data to post", function(){});
	 */
	this.post=function() {
		var p=[],av=arguments;
		$GP(p,av,[this.url,this.content,this.oncomplete],[null,null,$CB]);
		if(!p[0]&&!p[2]) return false;
		return(send(p[0],p[1],p[2],"POST",this.async));
	}
	/**
	 * 发送指定的表单到指定的Url
	 * @param	{String|Object}	formObject	表单对象或表单对象的ID
	 * @param	{function|Object|String}	[oncomplete]	正确返回时的回调函数，或者要更新的对象，或者要更新对象的ID
	 * @see		AJAXRequest#post
	 * @example
	 * var ajax = new AJAXRequest();
	 * ajax.postf("dataForm", function(obj) {
	 * 	alert(obj.responseText);
	 * });
	 */
	this.postf=function() {
		var p=[],fo,vaf,pcbf,purl,pc,pm,ac=arguments.length,av=arguments;
		fo=ac>0?$VO(av[0]):false;
		if(!fo||(fo&&fo.nodeName!="FORM")) return false;
		vaf=fo.getAttribute("onvalidate");
		vaf=vaf?(typeof(vaf)=="string"?new Function(vaf):vaf):null;
		if(vaf&&!vaf()) return false;
		$GP(p,[av[1],fo.getAttribute("action"),fo.getAttribute("method")],[this.oncomplete,this.url,this.method],[$CB,null,null]);
		pcbf=p[0];purl=p[1];
		if(!pcbf&&!purl) return false;
		pc=this.formToStr(fo); if(!pc) return false;
		if(p[2].toUpperCase()=="POST")
			return(send(purl,pc,pcbf,"POST",true));
		else {
			purl+=(purl.indexOf("?")>-1?"&":"?")+pc;
			return(send(purl,"",pcbf,"GET",true));
		}
	}
	/**
	 * 将表单对象转换化UrlEncode的字符串
	 * @author	SurfChen <surfchen@gmail.com>
	 * @link	http://www.surfchen.org/
	 * @param	{Object} formObject
	 * @returns {String} 表单字符串
	 * @see		AJAXRequest#postf
	 * @ignore
	 */
	this.formToStr=function(fc) {
		var i,qs="",and="",ev="";
		for(i=0;i<fc.length;i++) {
			e=fc[i];
			if (e.name!='') {
				if (e.type=='select-one'&&e.selectedIndex>-1) ev=e.options[e.selectedIndex].value;
				else if (e.type=='checkbox' || e.type=='radio') {
					if (e.checked==false) continue;
					ev=e.value;
				}
				else ev=e.value;
				ev=encode(ev); qs+=and+e.name+'='+ev; and="&";
			}
		}
		return qs;
	}
	if(ac>0){eval("this." + av.method.toLowerCase() + "()");}
}