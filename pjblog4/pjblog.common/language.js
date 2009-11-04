// JavaScript Document
var lang = new Object();
lang.Set = {
	ins : "",
	Asp : function(i){
		this.ins = i;
		var _this;
		switch (this.ins){
			case 1 : _this = "操作成功!"; break;
			case 2 : _this = "操作失败!"; break;
			case 3 : _this = "非法赋值!"; break;
			case 4 : _this = "验证码"; break;
			case 5 : _this = "看不清楚？点击刷新验证码！"; break;
			default : _this = "非法操作!";
		}
		return _this;
	},
	Js : function(i){
		this.ins = i;
		var _this;
		switch (this.ins){
			case 1 : _this = "测试"; break;
			default : _this = "非法操作!";
		}
		return _this;
	}
}