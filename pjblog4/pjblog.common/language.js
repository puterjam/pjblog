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
			case 6 : _this = "错误信息"; break;
			case 7 : _this = "请将信息输入完整"; break;
			case 8 : _this = "请返回重新输入"; break;
			case 9 : _this = "请输入登录验证码"; break;
			case 10 : _this = "非法用户名！<br/>请尝试使用其他用户名！"; break;
			case 11 : _this = "验证码有误"; break;
			case 12 : _this = "用户名与密码错误"; break;
			case 13 : _this = "登录成功"; break;
			case 14 : function(name){ _this = "<b>" + name + "</b>，欢迎你的再次光临。"}; break;
			case 15 : _this = "点击返回"; break;
			case 16 : _this = "点击返回登录前页面"; break;
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