// JavaScript Document
// Author : evio 
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

/* ------------------ 获取对象的绝对坐标 --------------------*/
function ABS(a){
	var b = { x: a.offsetLeft, y: a.offsetTop};
	a = a.offsetParent;
	while (a) {
		b.x += a.offsetLeft;
		b.y += a.offsetTop;
		a = a.offsetParent;
	}
	return b;
}
/* ------------------ 8位窗口 --------------------*/
var Box = {
	selfWidth : false, // 是否自定义宽度
	selefHeight : false, // 是否自定义高度
	offsetBoder : {
		HasBorder : true,
		Border : 1
	}, // 偏移量
	FollowBox : function(obj, newlayerWidth, newlayerHeight, postion, html){
	/*
		@ obj 				被定位的对象
		@ newlayerWidth		打开的新层的宽度
		@ newlayerHeight 	打开新层的高度
		@ postion 			打开新层的位置 共8个位置, 从顶部开始, 顺时针方向; 还有个位置是[0,0]
		@ html 				HTML代码
	*/
		// 创建一个上传的层
		var div = document.createElement("div");
		document.body.appendChild(div);
		div.innerHTML = html;
		
		var Ps = ABS(obj);
		var left = Ps.x;
		var top = Ps.y;
		var width = obj.offsetWidth;
		var height = obj.offsetHeight;
		if (!this.selfWidth){newlayerWidth = div.offsetWidth;}
		if (!this.selefHeight){newlayerHeight = div.offsetHeight;}
		
		var offPach = this.offsetBoder.HasBorder ? this.offsetBoder.Border : 0;
		var pos = new Array();
			pos.push(left);						// 0
			pos.push(top);						// 1
			pos.push(width);					// 2
			pos.push(height);					// 3
			pos.push(newlayerWidth);			// 4
			pos.push(newlayerHeight);			// 5
			pos.push(this.offsetBoder.Border); 	// 6
		var newPos = this.ChoosePostion(postion, pos);
		
		if (this.selfWidth){div.style.width = newlayerWidth;}
		if (this.selefHeight){div.style.height = newlayerHeight;}
		/*div.style.cssText += ";position: absolute; left:" + (newPos[0]) + "px; top:" + (newPos[1]) + "px;";*/
		div.style.position = "absolute";
		div.style.display = "block";
		div.style.left = newPos[0] + "px";
		div.style.top = newPos[1] + "px";
		return div; // 返回该新创建的对象,用于控制CSS
	},
	ChoosePostion : function(i, Arrays){
		var left, top;
		switch (i){
			case 1 :
				left = Arrays[0] - Arrays[6];
				top = Arrays[1] - Arrays[5] - (Arrays[6] * 2 + 1);
				break;
			case 2 :
				left = Arrays[0] + Arrays[2];
				top = Arrays[1] - Arrays[5] - (Arrays[6] * 2 + 1);
				break;
			case 3 :
				left = Arrays[0] + Arrays[2] + 1;
				top = Arrays[1];
				break;
			case 4 :
				left = Arrays[0] + Arrays[2] + 1;
				top = Arrays[1] + Arrays[3] + 1;
				break;
			case 5 :
				left = Arrays[0] - Arrays[6];
				top = Arrays[1] + Arrays[3] + 1;
				break;
			case 6 :
				left = Arrays[0] - Arrays[4] - (Arrays[6] * 2 + 1);
				top = Arrays[1] + Arrays[3] + 1;
				break;
			case 7 :
				left = Arrays[0] - Arrays[4] - (Arrays[6] * 2 + 2);
				top = Arrays[1];
				break;
			case 8:
				left = Arrays[0] - Arrays[4] - (Arrays[6] * 2 + 1);
				top = Arrays[1] - Arrays[5] - (Arrays[6] * 2 + 1);
				break;
			default :
				left = 0;
				top = 0;
		}
		return [left, top];
	}
}