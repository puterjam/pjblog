// -----------------------------------------------
//	Language Jscript For PJBlog Write By evio
//	Http://www.evio.name & Http://www.pjhome.net
// 	2009-09-19 13:15
//	CopyRight @ PJBlog Http://www.pjhome.net
// -----------------------------------------------
var lang = {
	/*
		Err Class. Deal with error description!
	*/
	Err : {
		Conn : "数据库连接出错，请检查连接字串!",
		Enable : "Blog不欢迎你的访问。",
		info : function(t){
			var temp;
			switch (t)
			{
				case 1 : temp = "用户名已注册"; break;
				case 2 : temp = "用户名未注册"; break;
				case 3 : temp = "非法参数"; break;
				case 4 : temp = "不允许外部链接提交数据"; break;
				case 5 : temp = "验证码有误，请返回重新输入"; break;
				case 6 : temp = "网站名称不能为空!"; break;
				case 7 : temp = "网站地址不能为空!"; break;
				case 8 : temp = "很抱歉，本站无此内容！"; break;
				case 9 : temp = "很抱歉，未经允许请勿盗链本站资源！"; break;
				case 10 : temp = "无法找到相应的模块!"; break;
				case 11 : temp = "加载模块发生异常!"; break;
				default : temp = "非法操作"; // 999 号错误
			}
			return temp;
		}
	},
	/*
		Action Class.
	*/
	Action : {
		Login : "登录",
		Logout : "退出系统",
		LoginSuc : function(t){
			var temp;
			switch (t)
			{
				case 1 : temp = "退出登录成功"; break;
				case 2 : temp = "三秒后自动返回登录前页面"; break;
				default : temp = "非法操作";
			}
			return temp;
		},
		Submit : "提交",
		ReSet : "重写",
		ins : "插入",
		Close : "关闭",
		Register : "注册新用户",
		PassWord : "密码",
		PassWordTip : "密码提示",
		Title : "标题",
		Category : "分类",
		AddNewCate : "增加新分类",
		ChooseCate : "选择分类",
		Cname : "别名",
		UpperPinyin : "自动转成开头大写拼音",
		LowerPinyin : "自动转成小写拼音",
		LogSetting : "日志设置",
		Control : "后台管理",
		Weather : {
			Sun : "晴天",
			Cloud : "多云",
			Wind : "疾风",
			Ice : "冰雹",
			Cloudy : "阴天",
			Rain : "下雨",
			Shower : "阵雨",
			Snow : "下雪"
		},
		logs : {
			CommDesc : "评论倒序",
			CommUnAble : "禁止评论",
			LogTop : "日志置顶",
			Meta : "隐私及Meta",
			PrivacySet : "设置日志隐私",
			MetaSet : "自定义日志页Meta信息",
			PrivateLog : "私密日志",
			EncryptedLog : "加密日志",
			PrivateInfo : "私密日志只有主人和作者能查阅",
			EncryptedInfo : "加密日志允许客人输入正确的密码即可查看",
			EnCloneLogSe : "不需要加密则留空",
			EnCloneLogTip : "不需要提示则留空",
			CloneTitle : "加密标题",
			CloneComment : "加密评论",
			MetaTagSet : "自定义日志页面头的Meta信息，留空则默认为Tag和日志摘要",
			LogKeyWordAlt : "填写你的keywords，利于搜索引擎，不需要则留空",
			KeyWords : "KeyWords",
			Description : "Description",
			LogDescriptionAlt : "填写你的Description，利于搜索引擎，不需要则留空",
			ComeFrom : "来源",
			ComeFromValue : "本站原创",
			WebSite : "网址",
			PostTime : "发表时间",
			LocalTime : "当前时间",
			AutoTime : "自定义日期",
			TimeForMat : "格式:yyyy-mm-dd hh:mm:ss",
			Tags : "Tags",
			InsTags : "插入已经使用的Tag",
			TagInfo : "tag之间用英文的空格或逗号分割",
			Content : "内容",
			DisImg : "禁止显示图片",
			DisSim : "禁止表情转换",
			DisUrl : "禁止自动转换链接",
			AutoKey : "禁止自动转换关键字",
			TotalLetter : "统计字数",
			ClearContent : "清空内容",
			SummyContent : "内容摘要",
			EditSummyContent : "编辑内容摘要",
			Upload : "附件上传",
			Quote : "引用通告",
			QuoteInfo : "请输入网络日志项的引用通告URL。可以用逗号分隔多个引用通告地址. ",
			SaveLog : "保存日志",
			PostLog : "提交日志",
			Draft : "保存并取消草稿",
			SaveDraft : "保存为草稿",
			DelLog : "删除该日志",
			DelLogInfo : "是否要删除该日志",
			CannerEdit : "取消编辑",
			EditType : "选择编辑类型",
			EditDraftInfo : "友情提示:保存草稿后，日志不会在日志列表中出现。只有再次编辑，<b>取消草稿</b>后才显示出来。",
			RetArticle : function(t){
				var temp;
				switch (t)
				{
					case 1 : temp = "无相关文章!"; break;
					case 2 : temp = "模式"; break;
					case 3 : temp = "全部显示"; break;
					case 4 : temp = "分页显示"; break;
					case 5 : temp = "分页"; break;
					case 6 : temp = "转到第1页"; break;
					case 7 : temp = "当前页"; break;
					case 8 : temp = function(t){return "转到第" + t + "页";}
							 break;
					case 9 : temp = function(t){return "共" + t + "个相关文章"}
							 break;
					default : temp = "非法操作";
				}
				return temp;
			},
			LogErr : function(t){
				var temp;
				switch (t)
				{
					case 1 : temp = "你没有没有权限修改日志"; break;
					case 2 : temp = "您没有没有权限发表新日志"; break;
					default : temp = "非法操作";
				}
				return temp;
			},
			LogeditErr : "日志发表错误",
			UBB : function(t){
				var temp;
				switch (t)
				{
					case 1 : temp = "UBB编辑器"; break;
					case 2 : temp = "UBBeditor"; break;
					default : temp = "非法操作";
				}
				return temp;
			},
			FCK : function(t){
				var temp;
				switch (t)
				{
					case 1 : temp = "FCK在线编辑器"; break;
					case 2 : temp = "FCKeditor"; break;
					default : temp = "非法操作";
				}
				return temp;
			},
			CateName : function(name){
				return "在 【" + name + "】 发表日志";
			}
		},
		Rss : function(a, b){
			return "订阅 " + a + " 所有文章(" + b + ")"
		}
	},
	
	/*
		Tip Class.
	*/
	Tip : {
		UBB : {
			Down : function(t){
				var temp;
				switch (t)
				{
					case 1 : temp = "下载文件"; break;
					case 2 : temp = "该文件只允许会员下载!"; break;
					case 3 : temp = "下载次数"; break;
					case 4 : temp = "文件大小超出，请返回重新上传!"; break;
					case 5 : temp = "文件格式非法，请返回重新上传!"; break;
					case 6 : temp = "文件上传成功，请返回继续上传!"; break;
					case 7 : temp = "确定上传"; break;
					case 8 : temp = "只对UBB编辑有效,媒体文件包括图片无效"; break;
					case 9 : temp = "文字水印"; break;
					case 10 : temp = "图片水印"; break;
					case 11 : temp = "允许盗链"; break;
					case 12 : temp = "水印类型"; break;
					case 13 : temp = "对不起，你没有权限上传附件!"; break;
					default : temp = "非法操作";
				}
				return temp;
			}	
		},
		Comment : function(t){
			var temp;
			switch (t)
			{
				case 1 : temp = "未审核评论,仅管理员可见."; break;
				default : temp = "非法操作";
			}
			return temp;
		},
		logs : {
			post : function(t){
				var temp;
				switch (t)
				{
					case 1 : temp = "返回你所发表的日志"; break;
					case 2 : temp = "发表日志"; break;
					case 3 : temp = "您必须选择分类后才可以进行下一步操作."; break;
					case 4 : temp = "如果您需要立即增加新分类请点击\"增加新分类\"链接."; break;
					case 5 : temp = "请选择编辑器类型,默认为UBB编辑器."; break;
					default : temp = "非法操作";
				}
				return temp;
			},
			edit : function(t){
				var temp;
				switch (t)
				{
					case 1 : temp = "返回你所编辑的日志"; break;
					case 2 : temp = "编辑日志"; break;
					default : temp = "非法操作";
				}
				return temp;
			},
			Ajax : function(t){
				var temp;
				switch (t)
				{
					case 1 : 
						temp = function(date){return "草稿于 " + date + " <font color='#9C0024'>保存</font>成功,请不要刷新本页,以免丢失信息!";}
						break;
					case 2 : 
						temp = function(date){return "草稿于 " + date + " <font color='#AF5500'>更新</font>成功,请不要刷新本页,以免丢失信息!";}
						break;
					default : temp = "非法操作";
				}
				return temp;
			}
		},
		SysTem : function(t){
			var temp;
			switch (t)
			{
				case 1 : temp = "错误信息"; break;
				case 2 : temp = "单击返回"; break;
				case 3 : temp = "单击返回上一页"; break;
				case 4 : temp = "单击返回首页"; break;
				case 5 : temp = "反馈信息"; break;
				case 6 : temp = "返回"; break;
				case 7 : temp = "下一步"; break;
				case 8 : temp = "单击返回退出前页面"; break;
				case 9 : temp = "返回登录前页面"; break;
				default : temp = "非法操作";
			}
			return temp;
		},
		Link : {
			Err : "友情链接发表出错",
			Suc : "友情链接添加成功",
			Aduit : "网站友情链接添加成功,请等待审核!",
			ImgLink : function(t){
				var temp;
				switch (t)
				{
					case 1 : temp = "图象链接"; break;
					case 2 : temp = "Image Link"; break;
					default : temp = "非法操作";
				}
				return temp;
			},
			TextLink : function(t){
				var temp;
				switch (t)
				{
					case 1 : temp = "文字链接"; break;
					case 2 : temp = "Text Link"; break;
					default : temp = "非法操作";
				}
				return temp;
			},
			Form : function(t){
				var temp;
				switch (t)
				{
					case 1 : temp = "申请友情链接"; break;
					case 2 : temp = "网站名称"; break;
					case 3 : temp = "网站地址"; break;
					case 4 : temp = "网站Logo"; break;
					case 5 : temp = "验证码"; break;
					case 6 : temp = "提示:带<span style=\"color:#f00\">*</span>项为必填项，网站的Logo和地址要写完整，必须包含 http://"; break;
					default : temp = "非法操作";
				}
				return temp;
			}
		},
		Control : function(t){
			var temp;
			switch (t)
			{
				case 1 : temp = "管理员密码"; break;
				case 2 : temp = "管理员密码错误!"; break;
				default : temp = "非法操作";
			}
			return temp;
		}
	}
}