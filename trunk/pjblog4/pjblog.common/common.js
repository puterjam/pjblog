// JavaScript Document
/*
	重构一些内置方法
*/
String.prototype.trim = function(){return this.toString().replace(/^(\s+)|(\s+)$/,"");};
String.prototype.reg = function(r){return r.test(this.toString());};
String.prototype.DeCode = function(){return decodeURIComponent(this.toString()).replace(/\+/g," ");}
String.prototype.utf8 = function(){return encodeURIComponent(this.toString());}
String.prototype.toNum = function(){return parseInt(this.toString());};
Array.prototype.trim = function(){
	var a = new Array();
	for(var i=0;i<this.length;i++){
		if(typeof(this[i])=="string"){
			a[i] = this[i].trim();
		}else{
			a[i] = this[i];
		}
	}
    return a;
};
String.prototype.json = function(){
	try{
		eval("var jsonStr = (" + this.toString() + ")");
	}catch(ex){
		var jsonStr = null;
	}
	return jsonStr;
};

/* ------------------ 定义一些正则判断方法 --------------------*/
var REGEXP = {
	REG_USERNAME : /^[0-9A-Za-z_]{5,16}$/, //用户名
	REG_QQ : /^\d{5,11}$/, //QQ
	REG_EMAIL : /^\w+((-\w+)|(\.\w+))*\@[A-Za-z0-9]+((\.|-)[A-Za-z0-9]+)*\.[A-Za-z0-9]+$/, // Email
	REG_WEBURL : /^(http|ftp|https):\/\/[\w\-_]+(\.[\w\-_]+)+([\w\-\.,@?^=%&:/~\+#]*[\w\-\@?^=%&/~\+#])?$/, //网址
	REG_TELPHONE : /^0\d{2,3}-\d{1,}$/, //电话
	REG_MOBILE : /^1\d{10}$/, // 手机
	REG_POSTCODE : /^\d{6}$/, //邮政编码
	REG_CHECKCODE : /^\d{4}$/, //验证码
	REG_ISCHINA : /^[\u4e00-\u9fa5]+$/ //是否包含汉字
}

/* ------------------ 定义Cookie的JS设置和获取 --------------------*/
var cookie = {
    SET	: function(name, value, days) {var expires = "";if (days) {var d = new Date();d.setTime(d.getTime() + days * 24 * 60 * 60 * 1000);expires = "; expires=" + d.toGMTString();}document.cookie = name + "=" + value + expires + "; path=/";},
	GET	: function(name) {var re = new RegExp("(\;|^)[^;]*(" + name + ")\=([^;]*)(;|$)");var res = re.exec(document.cookie);return res != null ? res[3] : null;}
};

/* ------------------ 定义获取ID对象的方法 --------------------*/
function $(){ 
    var elements = new Array(); 
    for (var i = 0; i < arguments.length; i++) 
    { 
        var element = arguments[i]; 
        if (typeof element == 'string') 
            element = document.getElementById(element); 
        if (element) {
        } else {
            element = null;
        }
        if (arguments.length == 1) {
            return element; 
        } else {
            elements.push(element); 
        }
    } 
    return elements; 
}