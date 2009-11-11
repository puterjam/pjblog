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
			case 14 : _this = function(name){ return "<b>" + name + "</b>，欢迎你的再次光临。";}; break;
			case 15 : _this = "点击返回"; break;
			case 16 : _this = "点击返回登录前页面"; break;
			case 17 : _this = "用户登录"; break;
			case 18 : _this = "用户名"; break;
			case 19 : _this = "密码"; break;
			case 20 : _this = "保存我的登录信息"; break;
			case 21 : _this = "登录"; break;
			case 22 : _this = "注册新用户"; break;
			case 23 : _this = "后台管理"; break;
			case 24 : _this = "登录失败"; break;
			case 25 : _this = "您没有权限发表新日志!"; break;
			case 26 : _this = "反馈信息"; break;
			case 27 : _this = "发表日志"; break;
			case 28 : _this = "请选择编辑器"; break;
			case 29 : _this = "UBB在线编辑器"; break;
			case 30 : _this = "FCK在线编辑器"; break;
			case 31 : _this = "下一步"; break;
			case 32 : _this = function(name){return "在 【" + name + "】 发表日志";}; break;
			case 33 : _this = "标题"; break;
			case 34 : _this = "别名"; break;
			case 35 : _this = "日志设置"; break;
			case 36 : _this = "自动转成开头大写拼音"; break;
			case 37 : _this = "自动转成小写拼音"; break;
			case 38 : _this = "晴天"; break;
			case 39 : _this = "多云"; break;
			case 40 : _this = "疾风"; break;
			case 41 : _this = "冰雹"; break;
			case 42 : _this = "阴天"; break;
			case 43 : _this = "下雨"; break;
			case 44 : _this = "阵雨"; break;
			case 45 : _this = "下雪"; break;
			case 46 : _this = "评论倒序"; break;
			case 47 : _this = "禁止评论"; break;
			case 48 : _this = "日志置顶"; break;
			case 49 : _this = "隐私及Meta"; break;
			case 50 : _this = "设置日志隐私"; break;
			case 51 : _this = "自定义日志页Meta信息"; break;
			case 52 : _this = "私密日志"; break;
			case 53 : _this = "加密日志"; break;
			case 54 : _this = "加密日志允许客人输入正确的密码即可查看"; break;
			case 55 : _this = "密码提示"; break;
			case 56 : _this = "不需要加密则留空"; break;
			case 57 : _this = "不需要提示则留空"; break;
			case 58 : _this = "自定义日志页面头的Meta信息，留空则默认为Tag和日志摘要"; break;
			case 59 : _this = "关键字"; break;
			case 60 : _this = "填写你的keywords，利于搜索引擎，不需要则留空"; break;
			case 61 : _this = "描述"; break;
			case 62 : _this = "填写你的Description，利于搜索引擎，不需要则留空"; break;
			case 63 : _this = "来源"; break;
			case 64 : _this = "本站原创"; break;
			case 65 : _this = "网址"; break;
			case 66 : _this = "发表时间"; break;
			case 67 : _this = "当前时间"; break;
			case 68 : _this = "自定义日期"; break;
			case 69 : _this = "格式:yyyy-mm-dd hh:mm:ss"; break;
			case 70 : _this = "标签"; break;
			case 71 : _this = "插入已经使用的标签"; break;
			case 72 : _this = "tag之间用英文的空格或逗号分割"; break;
			case 73 : _this = "内容"; break;
			case 74 : _this = "禁止显示图片"; break;
			case 75 : _this = "禁止表情转换"; break;
			case 76 : _this = "禁止自动转换链接"; break;
			case 77 : _this = "禁止自动转换关键字"; break;
			case 78 : _this = "统计字数"; break;
			case 79 : _this = "清空内容"; break;
			case 80 : _this = "内容摘要"; break;
			case 81 : _this = "编辑内容摘要"; break;
			case 82 : _this = "附件上传"; break;
			case 83 : _this = "引用通告"; break;
			case 84 : _this = "请输入网络日志项的引用通告URL。可以用逗号分隔多个引用通告地址. "; break;
			case 85 : _this = "保存日志"; break;
			case 86 : _this = "提交日志"; break;
			case 87 : _this = "保存并取消草稿"; break;
			case 88 : _this = "保存为草稿"; break;
			case 89 : _this = "删除该日志"; break;
			case 90 : _this = "友情提示:保存草稿后，日志不会在日志列表中出现。只有再次编辑，<b>取消草稿</b>后才显示出来。"; break;
			case 91 : _this = "上传"; break;
			case 92 : _this = "返回你所发表的日志"; break;
			case 93 : _this = "公开日志"; break;
			case 94 : _this = "隐藏日志"; break;
			default : _this = "非法操作!";
		}
		return _this;
	},
	Js : function(i){
		this.ins = i;
		var _this;
		switch (this.ins){
			case 1 : _this = "请选择分类"; break;
			default : _this = "非法操作!";
		}
		return _this;
	}
}