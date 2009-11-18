// JavaScript Document
// Author : evio 
// 特殊数组 : 用来压缩插件的信息
var PluginOutPutString = new Array(), PluginTempValue;

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

/* ---------------------------- 创建文件夹规则 example: ------------------------------ */
//<input onblur="ReplaceInput(this,window.event)" onkeyup="ReplaceInput(this,window.event)" />
function ReplaceInput(obj, cevent){
	var str = ["<", ">", "/", "\\", ":", "*", "?", "|", "\"", /[\u4E00-\u9FA5]/g];
	if(cevent.keyCode != 37 && cevent.keyCode != 39){
		//obj.value = obj.value.replace(/[\u4E00-\u9FA5]/g,'');
		for (var i = 0 ; i < str.length ; i++){
			obj.value = obj.value.replace(str[i], "");
		}
	}
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

function outputSideBar(html){
	html = html.replace(/[\n\r]/g,"");
	html = html.replace(/\\/g,"\\\\");
	html = html.replace(/\'/g,"\\'");
	return html;
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
		/*div.style.cssText += ";position: absolute; left:" + (newPos[0]) + "px; top:" + (newPos[1]) + "px;";*/
		div.style.position = "absolute";
		div.style.display = "block";
		div.style.left = newPos[0] + "px";
		div.style.top = newPos[1] + "px";
		if (this.selfWidth){div.style.width = newlayerWidth + "px";}
		if (this.selefHeight){div.style.height = newlayerHeight + "px";}
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
				left = Arrays[0];
				top = Arrays[1];
		}
		return [left, top];
	}
}

/* JS 操作表格 */
function TableAddRow(obj){ this._obj = obj ; this.Mark = "AddRowMark"; }

TableAddRow.prototype.GetRowId = function(){
	var arr = this._obj.rows, o = [false, 0];
	for (var i = 0 ; i < arr.length ; i++){
		if (arr[i].id == this.Mark){o = [true, i]; break;}
	}
	return o;
}

TableAddRow.prototype.AddRow = function(offset){
	var arr = this.GetRowId();
	if (arr[0]){
		return this._obj.insertRow(parseInt(arr[1] + parseInt(offset)));
	}else{
		return null;
	}
}

/* 添加新分类 */
function AddNewCateRow(){
	var NewCate = new TableAddRow($("CateTable"));
	if (!$(NewCate.Mark + "_del")){
		// 判断按钮
		var button = $("Addbutton");
		if (!button.disabled){button.disabled = true;}else{button.disabled = false;}
		
		var tr = NewCate.AddRow(0);//  插入一行
		if (tr != null){
			tr.id = NewCate.Mark + "_del";
			tr.style.lineHeight = "30px";
			tr.style.backgroundColor = "#FFFF99";
			// 插入第一个数据 0 
			var _td = tr.insertCell(0);
			var div = document.createElement("div");
			div.innerHTML = "<a href=\"javascript:;\" onclick=\"try{$('" + NewCate.Mark + "_del').parentNode.removeChild($('" + NewCate.Mark + "_del'));$('Addbutton').disabled=false;}catch(e){}\" id=\"new_selectid\"><img src=\"../images/check_error.gif\" border=\"0\"></a>";
			_td.appendChild(div);
			//  插入第二个数据 1
			_td = tr.insertCell(1);
			div = document.createElement("div");
			div.innerHTML = "<input type=\"text\" id=\"new_order\" class=\"text\" size=\"2\" name=\"cate_Order\" />";
			_td.appendChild(div);
			//  插入第三个数据 2
			_td = tr.insertCell(2);
			div = document.createElement("div");
			div.innerHTML = "<img id=\"iconimg\" src=\"" + carr[1] + "\" width=\"16\" height=\"16\" /> <select id=\"new_icon\" onchange=\"$('iconimg').src=this.options[this.options.selectedIndex].value\" style=\"width:120px;\" name=\"Cate_icons\">" + carr[0] + "</select>";
			_td.appendChild(div);
			// 插入第四个数据 3
			_td = tr.insertCell(3);
			div = document.createElement("div");
			div.innerHTML = "<input id=\"new_name\" type=\"text\" class=\"text\" value=\"\" size=\"14\" name=\"Cate_Name\" />";
			_td.appendChild(div);
			// 插入第五个数据 4
			_td = tr.insertCell(4);
			div = document.createElement("div");
			div.innerHTML = "<input id=\"new_Intro\" type=\"text\" class=\"text\" size=\"20\" name=\"Cate_Intro\"/>";
			_td.appendChild(div);
			// 插入第六个数据 5
			_td = tr.insertCell(5);
			div = document.createElement("div");
			div.innerHTML = "<input id=\"new_Part\" type=\"text\" class=\"text\" size=\"16\" onblur=\"ReplaceInput(this,window.event)\" onkeyup=\"ReplaceInput(this,window.event)\" name=\"Cate_Part\" />";
			_td.appendChild(div);
			// 插入第七个数据 6
			_td = tr.insertCell(6);
			div = document.createElement("div");
			div.innerHTML = "<input id=\"new_URL\" type=\"text\" size=\"30\" class=\"text\" name=\"cate_URL\" />";
			_td.appendChild(div);
			// 插入第八个数据 7
			_td = tr.insertCell(7);
			div = document.createElement("div");
			div.innerHTML = "<select id=\"new_local\" name=\"Cate_local\"><option value=\"0\">同时</option><option value=\"1\">顶部</option><option value=\"2\">侧边</option></select>";
			_td.appendChild(div);
			// 插入第九个数据 8
			_td = tr.insertCell(8);
			div = document.createElement("div");
			div.innerHTML = " <select id=\"new_Secret\" name=\"cate_Secret\"><option value=\"0\" style=\"background:#0f0\">公开</option><option value=\"1\" style=\"background:#f99\">保密</option></select>";
			_td.appendChild(div);
			// 插入第十个数据 9
			_td = tr.insertCell(9);
			div = document.createElement("div");
			div.innerHTML = "<a href=\"javascript:;\" onclick=\"CheckForm.Category.Add(this);\">保存新分类</a>";
			_td.appendChild(div);
		}
	}
}

/* --------------- JS操作权限 ------------------ */
var JsCopy = {
	index :{
		edit : function(){
			var c = document.getElementsByTagName("span");
			for (var i = 0 ; i < c.length ; i++){
				if (c[i].className == "index_edit") c[i].parentNode.removeChild(c[i]);
			}
		},
		del : function(){
			var c = document.getElementsByTagName("span");
			for (var i = 0 ; i < c.length ; i++){
				if (c[i].className == "index_del") c[i].parentNode.removeChild(c[i]);
			}
		},
		enterArticle : function(id){
			cookie.SET(cookies + "_enterArticle", id, 1);
		}
	}
}

/* --------------- JS提示框 ------------------ */
var Pos = 0, Dir = 2, len = 0;
var time = 0, en = null, limit = 6, Source = null, Target = null;
function flash(obj){
	if (time % 2 == 0){
		$(obj).style.cssText += ";background:" + Target;
	}else{
		$(obj).style.cssText += ";background:" + Source;
	}
	if (time > limit){
		clearTimeout(en);
	}else{
		time++;
		en = setTimeout("flash(\"" + obj + "\")", 100);
	}
}

function animate()
{
var elem = $('progress');
	if(elem != null) {
		if (Pos == 0) len += Dir;
		if (len > 32 || Pos > 249) Pos += Dir;
		if (Pos > 249) len -= Dir;
		if (Pos > 249 && len == 0) Pos = 0;
		elem.style.left = Pos + "px";
		elem.style.width = len + "px";
	}
}
var Tip = {
	coint : null,
	timer : 3,
	create : function(obj, id){
		var temp = document.createElement(obj);
		temp.id = id;
		return temp;
	},
	CreateLayer : function(title, html){
		var _this = this;
		var loader_container = this.create("div", "loader_container");
		var loader = this.create("div", "loader");
		var load_title = this.create("div", "load_title");
		var load_body = this.create("div", "load_body");
		var loader_bg = this.create("div", "loader_bg");
		var progress = this.create("div", "progress");
		load_title.style.textAlign = "center";
		load_body.style.textAlign = "center";
		load_title.innerHTML = title;
		load_body.innerHTML = html;
		loader_bg.appendChild(progress);
		loader.appendChild(load_title);
		loader.appendChild(load_body);
		loader.appendChild(loader_bg);
		loader_container.appendChild(loader);
		document.body.appendChild(loader_container);
		loader_container.style.top = (document.documentElement.clientHeight - loader_container.offsetHeight)/2 + document.documentElement.scrollTop + "px";
		this.coint = setInterval(animate, 10);
		setTimeout(_this.remove, 3000)
	},
	remove : function(){
		clearInterval(this.coint);
		document.body.removeChild($("loader_container"));
	}
}