// JavaScript Document
function $(el){
	return document.getElementById(el);	
}
 

var Dragtemplate = {
	down : false,
	MouserPole : {x : 0, y : 0}, // 记录鼠标位置, 初始化鼠标位置.
	DivPole : {x : 0, y : 0},
	DivCss : {width : 0, height : 0},
	DivBorder : "",
	DivDom : {box : new Array(), frame : 4},
	parent : null,
	filter : "-moz-opacity:0.5;opacity:0.5;filter:alpha(opacity=50);-ms-filter:'alpha(opacity=50)';background-color: #E1E1E1",
	pressdown : function(obj){
		
		_this = this;
		this.down = true;
		
		this.parent = obj;
		
		var bi = this.getxy(obj);
		var x = bi.l;
		var y = bi.t;
		_this.DivBorder = obj.style.cssText;
		obj.style.cssText += ";border:1px solid #ff7788";
		var width = obj.offsetWidth - 2;
		var height = obj.offsetHeight - 2;
		_this.DivPole.x = x;
		_this.DivPole.y = y;
		_this.DivCss.width = width;
		_this.DivCss.height = height;

		_this.CreateTempDiv(obj);
		
		obj.style.cursor = "move";
		obj.style.cssText += ";" + _this.filter;
		obj.style.position = "absolute";
		obj.style.left = x + 3 + "px";
		obj.style.top = y + 3 + "px";
		obj.style.width = _this.DivCss.width + "px";
		obj.style.height = _this.DivCss.height + "px";
				
		var e = e || window.event;
		var o = _this.divPole(e, obj); // 获得鼠标位置
		_this.MouserPole.x = o.x; // 更新鼠标位置 x
		_this.MouserPole.y = o.y; // 更新鼠标位置 y
		obj.onmouseup = function(){
			_this.pressup(this);
		}
		obj.onmousemove = function(){
			_this.pressmove(this);
		}
	},
	pressup : function(obj){
		this.down = false;
		this.parent = null;
		obj.style.cursor = "";
		obj.style.position = "";
		obj.style.cssText = this.DivBorder;
		$("TempDragDiv").parentNode.replaceChild(obj, $("TempDragDiv"))
		this.MouserPole.x = 0;
		this.MouserPole.y = 0;
		this.DivPole.x = 0;
		this.DivPole.y = 0;
		obj.onmousemove = function(){return;}
	},
	pressmove : function(obj){
		_this = this;
		var e = e || window.event;
		if (_this.down){
			var c = _this.getPole(e);
			obj.style.top = c.y + "px";
			obj.style.left = c.x + "px";
			_this.toTempDiv(e);
		}
	},
	/*获取对象位置*/
	getPole : function(e){
		_this = this;
		var c = {};
		if (document.all){
			c.x = e.clientX - _this.MouserPole.x;
			c.y = e.clientY - _this.MouserPole.y;
		}else{
			c.x = e.pageX;
			c.y = e.pageY;
		}
 		return c;
	},
	/*获取鼠标在div中的位置*/
	divPole : function(e, obj){
		var to = this.getxy(obj);
		var left = to.l;
		var top = to.t;
		left = e.clientX - left;
		top = e.clientY - top;
		return {
			x : left,
			y : top
		}
	},
	/*获取坐标相关*/
	getxy : function(m){
		var t = m.offsetTop;
		var l = m.offsetLeft;
		var w = m.offsetWidth;
		var h = m.offsetHeight;
		while(m = m.offsetParent){
			t += m.offsetTop;
			l += m.offsetLeft;
		}
  		return {t : t, l : l, w : w, h : h};
	},
	/*创建临时模板*/
	CreateTempDiv : function(obj){
		_this = this;
		var div = document.createElement("div");
		div.id = "TempDragDiv";
		obj.parentNode.insertBefore(div, obj);
		div.style.cssText = "border:1px dotted green; width:" + (_this.DivCss.width) + "px; height:" + (_this.DivCss.height) + "px";
	},
	/*删除临时模板*/
	RemoveTempDiv : function(){$("TempDragDiv").parentNode.removeChild($("TempDragDiv"));},
	pike : function(e, obj){
		var o = this.getxy(obj);
		if ((e.x > o.l) && (e.x < (o.l + o.w)) && (e.y > o.t) && (e.y < (o.t + o.h))){
			if (e.y < (o.t + o.h / 2)){
				return 1;
			}else{
				return 2;
			}
		}else{
			return 0;
		}
	},
	toTempDiv : function(e){
		for (var i = 0 ; i < this.DivDom.box.length ; i++){
			if (!this.DivDom.box[i])
				continue
			if (this.DivDom.box[i] == this.parent)
				continue
			var result = this.pike(e, this.DivDom.box[i]);
			if (result == 0)
				continue
			if (result == 1){
				this.DivDom.box[i].parentNode.insertBefore($("TempDragDiv"), this.DivDom.box[i]);
			}else{
				if(this.DivDom.box[i].nextSibling == null){
					this.DivDom.box[i].parentNode.appendChild($("TempDragDiv"))
				}else{
					this.DivDom.box[i].parentNode.insertBefore($("TempDragDiv"), this.DivDom.box[i].nextSibling)
				}
			}
			$("TempDragDiv").style.width = "100%";
		}
		
		for(var j = 1 ;j <= this.DivDom.frame ; j++){
			var r = 0, g = $("fame_" + j).getElementsByTagName("div");
			for (var b = 0 ; b < g.length ; b++) { if (g[b].className == "drag"){r++;} }
			if (r == 0){
			var k = this.getxy($("fame_" + j))
				if(e.x > k.l && e.x < (k.l + k.w)){
					$("fame_" + j).appendChild($("TempDragDiv"))
					$("TempDragDiv").style.width = "100%";
				}
			}
		}
	},
	/*给每个DIV增加拖动事件*/
	AddEvent : function(){
		_this = this;
		document.onselectstart = function(){return false}
		var c = document.getElementsByTagName("div");
		for (var t = 0 ; t < c.length ; t++){
			if (c[t].className == "drag") {_this.AddArrElement(c[t]); c[t].onmousedown = function(){_this.pressdown(this);}}
		}
	},
	AddArrElement : function(obj){this.DivDom.box.push(obj);}
}
window.onerror=function(){return false};
window.onload = function(){Dragtemplate.AddEvent();}