// JavaScript Document
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

String.prototype.reg = function(r){return r.test(this.toString());};

String.prototype.json = function(){
	try{
		eval("var jsonStr = (" + this.toString() + ")");
	}catch(ex){
		var jsonStr = null;
	}
	return jsonStr;
};

var cookie = {
    SET	: function(name, value, days) {var expires = "";if (days) {var d = new Date();d.setTime(d.getTime() + days * 24 * 60 * 60 * 1000);expires = "; expires=" + d.toGMTString();}document.cookie = name + "=" + value + expires + "; path=/";},
	GET	: function(name) {var re = new RegExp("(\;|^)[^;]*(" + name + ")\=([^;]*)(;|$)");var res = re.exec(document.cookie);return res != null ? res[3] : null;}
};
var PluginOutPutString = new Array(), PluginTempValue;
function outputSideBar(html){
	html = html.replace(/[\n\r]/g,"");
	html = html.replace(/\\/g,"\\\\");
	html = html.replace(/\'/g,"\\'");
	return html;
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

/*
	@ 全局方法开始
*/
var ceeevio = {
	/*
		' @ 清理缓存
		' @ 参数2个
		' @ obj 对象实例
	*/
	clearApplication : function(obj){
		cChecked(obj);
		$("#clear").html("<div id='flade'>Waiting...</div>");
		ci = true;
		fadeout();
		$.get(
			"../pjblog.logic/control/log_category.asp?s=" + Math.random(),
			{action : "clear"},
			function(Data){
				$("#Sbox").html(Data);
				scollbottom();
				setTimeout("ci = false; $('#clear').html('Complete...')", 1000);
			}
		);
		
	},	
	/*
		' @ 获得验证码
		' @ 参数2个
		' @ {Str : "文件地址", id : "对象ID名"}
	*/	
	GetCode : function(Str, id){
		$("#" + id).html("<img id=\"vcodeImg\" src=\"about:blank\" onerror=\"this.onerror=null;this.src='" + Str + "?s='+Math.random();\" alt=\"" + lang.Set.Asp(4) + "\" title=\"" + lang.Set.Asp(5) + "\" style=\"margin-right:40px;cursor:pointer;width:40px;height:18px;margin-bottom:-4px;margin-top:3px;\" onclick=\"src='" + Str + "?s='+Math.random()\"/>");
	},
	/*
		' @ 静态化过程
	*/
	Static : {
		/* ------------------- 重建首页 ------------------ */
		s_defailt : function(obj){
			cChecked(obj);
			$("#Sbox").empty();
			$("#clear").html("<div id='flade'>Waiting...</div>");
			ci = true;
			fadeout();
			$.getJSON(
				"../pjblog.logic/log_webconfig.asp?s=" + Math.random(),
				{action : "default"},
				function(Data){
					$("#Sbox").append("<div style=\"text-align:right; line-height:18px\"><span style=\"float:left\">重建首页及首页分页完毕</span>" + Data.Info + "</div>");
					scollbottom();
					setTimeout("ci = false; $('#clear').html('Complete...')", 1000);
				}
			);
		},
		/* ------------------- 重建日志页 ------------------ */
		s_Article : {
			//	@  获取日志ID总数
			getID : function(obj){
				var _this = this, _obj = obj;
				cChecked(obj);
				$("#Sbox").empty();
				$("#clear").html("<div id='flade'>Waiting...</div>");
				$("#Sbox").append("<div style=\"text-align:right; line-height:18px;\"><span style=\"float:left\">准备获取所有日志ID, 请稍后...</span></div><div class=\"clear\"></div>");
				ci = true;
				fadeout();
				$.getJSON(
					"../pjblog.logic/log_webconfig.asp?s=" + Math.random(),
					{action : "getarticleid"},
					function(Data){
						if (Data.Suc){
							if (Data.total = 0){
								$("#Sbox").append("<div style=\"text-align:right; line-height:18px\"><span style=\"float:left\">没有日志需要重建</span></di<div class=\"clear\"></div>v>");
								scollbottom();
								setTimeout("ci = false; $('#clear').html('Complete...')", 1000);
							}else if(Data.total < 0){
								$("#Sbox").append("<div style=\"text-align:right; line-height:18px\"><span style=\"float:left\">获取ID过程出错</span></div<div class=\"clear\"></div>>");
								scollbottom();
								setTimeout("ci = false; $('#clear').html('Complete...')", 1000);
							}else{
								ArtArray = Data.id;
								var n = 1;
								$("#Sbox").append("<div style=\"text-align:right; line-height:18px\"><span style=\"float:left\">获取ID成功,开始重建所有日志.</span></div><div class=\"clear\"></div>");
								scollbottom();
								$("#clear").prepend("<div class=\"rebuilt\">正在重建第 <span id=\"RebuiltID\">0</span> 篇日志<div>")
								_this.Article(ArtArray[n - 1], n, Data.id.length, _obj);
							}
						}else{
							$("#Sbox").append("<div style=\"text-align:right; line-height:18px\"><span style=\"float:left\">没有日志需要重建</span></div><div class=\"clear\"></div>");
							scollbottom();
							setTimeout("ci = false; $('#clear').html('Complete...')", 1000);
						}
					}
				);
			},
			Article : function(i, n, t, obj){
				//  @  单个日志重建
				if (n > t){
					ci = false;
					$("#Sbox").append("<div style=\"text-align:right; line-height:18px\"><span style=\"float:left\">所有日志重建完毕,共重建 " + t + " 篇日志...</span></div><div class=\"clear\"></div>");
					$("#clear").html("Complate...");
					scollbottom();
					ArtArray = new Array();
					return;
				}
				$("#RebuiltID").text(n);
				var _this = this, _obj = obj, _n = n, _t = t;
				$.getJSON(
					"../pjblog.logic/log_webconfig.asp?s=" + Math.random(),
					{action : "article", id : i},
					function(Data){
						if (Data.Suc){
							$("#Sbox").append("<div style=\"text-align:right; line-height:18px\"><span style=\"float:left;\">&nbsp;&nbsp;&nbsp;&nbsp;<font color=\"#0000ff\">" + Data.Title + "</font></span>" + Data.Path + "</div><div class=\"clear\"></div>");
							scollbottom();
						}else{
							$("#Sbox").append("<div style=\"text-align:right; line-height:18px\"><span style=\"float:left\"><font color=\"#ff0000\">重建日志ID为" + _n + "的日志失败</font></span></div><div class=\"clear\"></div>");
							scollbottom();
						}
						_n++;
						_this.Article(ArtArray[_n - 1], _n, _t, _obj);
					}
				);
			}
		},
		s_category : {
			//	@  获取分类ID总数
			getID : function(obj){
				var _this = this, _obj = obj;
				cChecked(obj);
				$("#Sbox").empty();
				$("#clear").html("<div id='flade'>Waiting...</div>");
				$("#Sbox").append("<div style=\"text-align:right; line-height:18px;\"><span style=\"float:left\">准备获取所有分类ID, 请稍后...</span></div><div class=\"clear\"></div>");
				ci = true;
				fadeout();
				$.getJSON(
					"../pjblog.logic/log_webconfig.asp?s=" + Math.random(),
					{action : "getcategoryid"},
					function(Data){
						if (Data.Suc){
							if (Data.total = 0){
								$("#Sbox").append("<div style=\"text-align:right; line-height:18px\"><span style=\"float:left\">没有分类需要重建</span></di<div class=\"clear\"></div>v>");
								scollbottom();
								setTimeout("ci = false; $('#clear').html('Complete...')", 1000);
							}else if(Data.total < 0){
								$("#Sbox").append("<div style=\"text-align:right; line-height:18px\"><span style=\"float:left\">获取ID过程出错</span></div<div class=\"clear\"></div>>");
								scollbottom();
								setTimeout("ci = false; $('#clear').html('Complete...')", 1000);
							}else{
								ArtArray = Data.id;
								var n = 1;
								$("#Sbox").append("<div style=\"text-align:right; line-height:18px\"><span style=\"float:left\">获取分类ID成功,开始重建所有日分类.</span></div><div class=\"clear\"></div>");
								scollbottom();
								$("#clear").prepend("<div class=\"rebuilt\">正在重建第 <span id=\"RebuiltID\">0</span> 号分类<div>")
								_this.category(ArtArray[n - 1], n, ArtArray.length, _obj);
							}
						}else{
							$("#Sbox").append("<div style=\"text-align:right; line-height:18px\"><span style=\"float:left\">没有分类需要重建</span></div><div class=\"clear\"></div>");
							scollbottom();
							setTimeout("ci = false; $('#clear').html('Complete...')", 1000);
						}
					}
				);
			},
			category : function(i, n, t, obj){
				//  @  单个分类重建
					if (n > t){
						ci = false;
						$("#Sbox").append("<div style=\"text-align:right; line-height:18px\"><span style=\"float:left\">所有分类重建完毕,共重建 " + t + " 个分类...</span></div><div class=\"clear\"></div>");
						$("#clear").html("Complate...");
						scollbottom();
						ArtArray = new Array();
						return;
					}
					$("#RebuiltID").text(n);
					var _this = this, _obj = obj, _n = n, _t = t, f = i[1];
					$.getJSON(
						"../pjblog.logic/log_webconfig.asp?s=" + Math.random(),
						{action : "category", id : i[0]},
						function(Data){
							if (Data.Suc){
								$("#Sbox").append("<div style=\"text-align:right; line-height:18px\"><span style=\"float:left;\">&nbsp;&nbsp;&nbsp;&nbsp;<font color=\"#0000ff\">静态文件夹为 " + f + " 的分类重建成功</font></span>" + Data.Info + "</div><div class=\"clear\"></div>");
								scollbottom();
							}else{
								$("#Sbox").append("<div style=\"text-align:right; line-height:18px\"><span style=\"float:left\"><font color=\"#ff0000\">重建日志ID为" + _n + "的日志失败</font></span></div><div class=\"clear\"></div>");
								scollbottom();
							}
							_n++;
							_this.category(ArtArray[_n - 1], _n, _t, _obj);
						}
					);
			}
		}//,
	},
	category : {
		add : function(){
			$(".AddA").attr("disabled", true);
			canSubmit = false;
			var str = "<tr class=\"SecTd\" id=\"newAdd\"><td id=\"delimg\"><a href=\"javascript:;\" onclick=\"try{ceeevio.category.remove();}catch(e){}\"><img src=\"../images/check_error.gif\" border=\"0\"></a></td>";
			str += "<td><input type=\"text\" id=\"new_order\" class=\"text\" size=\"2\" name=\"cate_Order\" /></td>";
			str += "<td><img id=\"iconimg\" src=\"" + carr[1] + "\" width=\"16\" height=\"16\" /> <select id=\"new_icon\" onchange=\"$('#iconimg').attr('src', this.options[this.options.selectedIndex].value)\" style=\"width:120px;\" name=\"Cate_icons\">" + carr[0] + "</select></td>";
			str += "<td><input id=\"new_name\" type=\"text\" class=\"text\" value=\"\" size=\"14\" name=\"Cate_Name\" /></td>";
			str += "<td><input id=\"new_Intro\" type=\"text\" class=\"text\" size=\"20\" name=\"Cate_Intro\"/></td>";
			str += "<td><input id=\"new_Part\" type=\"text\" class=\"text\" size=\"16\" onblur=\"ReplaceInput(this,window.event)\" onkeyup=\"ReplaceInput(this,window.event)\" name=\"Cate_Part\" /></td>";
			str += "<td><input id=\"new_URL\" type=\"text\" size=\"30\" class=\"text\" name=\"cate_URL\" /></td>";
			str += "<td><select id=\"new_local\" name=\"Cate_local\"><option value=\"0\">同时</option><option value=\"1\">顶部</option><option value=\"2\">侧边</option></select></td>"
			str += "<td><select id=\"new_Secret\" name=\"cate_Secret\"><option value=\"0\" style=\"background:#0f0\">公开</option><option value=\"1\" style=\"background:#f99\">保密</option></select</td>>";
			str += "<td id=\"singlesave\"><a href=\"javascript:;\" onclick=\"ceeevio.category.SinglePost(this);\">保存新分类</a></td></tr>";
			$("#AddRowMark").before(str);
		},
		remove : function(){
			$("#newAdd").remove();
			$(".AddA").attr("disabled", false);
			canSubmit = true;
		},
		SinglePost : function(obj){
			
			$(".AddA").attr("disabled", false);
			
			var new_order 	= 	$("#new_order").val();
			var new_icon 	= 	$("#new_icon").val();
			var new_name 	= 	$("#new_name").val();
			var new_Intro 	= 	$("#new_Intro").val();
			var new_Part 	= 	$("#new_Part").val();
			var new_URL 	= 	$("#new_URL").val();
			var new_local 	= 	$("#new_local").val();
			var new_Secret 	= 	$("#new_Secret").val();
			
			// 验证数据
			if ((new_order.legnth < 1) || (!/^\d+$/.test(new_order))){
				alert("您所填写的排序格式不正确!"); $("#new_order").select(); $(".AddA").attr("disabled", false); return;
			}
			if (new_icon.length < 1){
				alert("您所选择的图标不正确!"); $(".AddA").attr("disabled", false); return;
			}
			if (new_name.length < 2){
				alert("标题格式不正确或为空, 标题应大于2位字符!"); $("#new_name").select(); $(".AddA").attr("disabled", false); return;
			}
			
			$.post(
				"../pjblog.logic/control/log_category.asp?action=add&s=" + Math.random(),
				{cate_Order : new_order, cate_icon : new_icon, cate_Name : new_name, cate_Intro : new_Intro, cate_Folder : new_Part, cate_URL : new_URL, cate_local : new_local, cate_Secret : new_Secret},
				function(Data){
					var json = Data.json();
					if (json.Suc){
						canSubmit = true;
						$(".AddA").attr("disabled", false);
						var _this = $("#newAdd");
						//动画效果
						_this.seekAttention(); // 将遮盖效果换成闪动效果
//						var already = effect.windows.open(_this, {
//														  			position 	: 		0,
//																	html 		:		"保存成功",
//																	offset		:		0
//														  });
//						jQuery(already).addClass("opacity");
//						jQuery(already).attr("id", "ctip");
//						jQuery(already).css({
//																	width 			: 		_this.outerWidth() + "px",
//																	height 			: 		_this.outerHeight() + "px",
//																	"line-height"	:		_this.innerHeight() + "px",
//																	color			: 		"#000",
//																	"text-align"	: 		"center",
//																	"font-size"		:		"13px",
//																	"font-family"	:		"微软雅黑"
//											});
						// 	@	 替换文本
						$("#delimg").replaceWith("<td width=\"25\"><input type=\"checkbox\" value=\"" + json.Info + "\" name=\"SelectID\" /><input type=\"hidden\" value=\"" + json.Info + "\" name=\"Cate_ID\" /></td>");
						$("#singlesave").replaceWith("<td><input type=\"text\" class=\"text\" name=\"cate_count\" value=\"0\" size=\"2\" readonly=\"readonly\" style=\"background:#ffe\"/> 篇</td>");
						
						
						//移除ID
						try{
						$("#newAdd").removeAttr("id");
						$("#new_order").removeAttr("id");
						$("#new_icon").removeAttr("id");
						$("#new_name").removeAttr("id");
						$("#new_Intro").removeAttr("id");
						$("#new_Part").removeAttr("id");
						$("#new_URL").removeAttr("id");
						$("#new_local").removeAttr("id");
						$("#new_Secret").removeAttr("id");
						}catch(e){alert(e.description)}
						setTimeout("$('#ctip').remove();", 2000);
					}else{
						alert(json.Info);
						$(".AddA").attr("disabled", false);
					}
				}
			);
		}
	},
	theme : {
		layout : function(f1, obj){
			var _obj = obj, _f1 = f1;
			$.post(
				"../pjblog.logic/control/log_template.asp?action=Choose&s=" + Math.random(),
				{f1 : f1},
				function(data){
					var json = data.json();
					if (json.Suc){
						var c = "<div style=\"padding:1px 20px 0px 20px;\">";
						c += "<form action=\"../pjblog.logic/control/log_template.asp?action=update\" method=\"post\" style=\"font-size:12px;\">";
						c += "<input type=\"hidden\" value=\"" + _f1 + "\" name=\"f1\">";
						c += "<div>请选择样式</div>";
						c += "<div style=\"list-style:none\">";
						c += json.Info;
						c += "</div>";
						c += "<input type=\"submit\" value=\"确定\"> &nbsp; <input type=\"button\" value=\"关闭\"onclick=\"$('#templateTip').remove();\">";
						c += "</form></div>";
						var m = jQuery(_obj);
						jQuery(effect.windows.open( m, {
								position 	: 		3,
								html 		:		c,
								offset 		: 		1
						}))
						.css({
								border		:		"1px solid #000",
								background	:		"#fff",
								padding     :		"10px 2px 10px 2px"
						})
						.attr("id", "templateTip")
						.seekAttention();
					}else{
						alert(json.Info)
					}
				}
			);
		},
		AddPlus : function(f1, folder, tp_pluginSingleMark, _obj){
			jQuery(_obj).html("正在添加...")
			var obj = _obj;
			$(".doAddplus")
				.attr("disabled", true)
				.bind("click", function(){return false;});
			$.post(
				"../pjblog.logic/control/log_template.asp?action=AddPlus&s=" + Math.random(),
				{f1 : f1, folder : folder, tp_pluginSingleMark : tp_pluginSingleMark},
				function(data){
					jQuery(obj).html("操作成功!");
					$(".doAddplus").attr("disabled", false)
					setTimeout("location.reload()", 500);
				}
			);
		},
		ImportPlus : function(data){
			var json = data.json();
			if (json.Suc){
				effect.WarnTip.open({
						title	: 	"恭喜 操作成功",
						html	:	json.Info
				}
			)}else{
				effect.WarnTip.open({
						title	: 	"抱歉 操作失败",
						html	:	json.Info
				});
			}
		}//,
	},
	Comment : {
		post : function(){
			$("#postform").ajaxForm({
				dataType : "json",
				beforeSubmit : function(){
					// 判断用户名
					var reg;
					if ($("#comm_Author").val().length == 0){
						alert("请填写用户名!");
						$("#comm_Author").seekAttention();
						return false;
					}
					// 判断邮箱
					if ($("#comm_Email").val().length > 0){
						reg = /\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*/ig;
						if (!reg.test($("#comm_Email").val())){
							alert("邮箱格式不正确");
							$("#comm_Email").seekAttention();
							return false;
						}
					}
					// 判断网址
					if ($("#comm_WebSite").val().length > 0){
						reg = /(http):\/\/[\w\-_]+(\.[\w\-_]+)+([\w\-\.,@?^=%&:/~\+#]*[\w\-\@?^=%&/~\+#])?/ig;
						if (!reg.test($("#comm_WebSite").val())){
							alert("网址格式不正确");
							$("#comm_WebSite").seekAttention();
							return false;
						}
					}
					if ($("#comm_Content").val().length == 0){
						alert("评论内容不能为空");
						$("#comm_Content").seekAttention();
						return false;
					}
				},
				success : function(data){
					$("#postform").resetForm();
					var str = cee.decode(data.Info);
					var div = document.createElement("div");
					jQuery(div).attr("id", "comment_" + data.id).addClass("CommPart").html(str);
					$("#commentBox").prepend(div);
					if (CommPageSize < $(".CommPart").length){
						$(".CommPart:last").remove();
					}
					jQuery(div).seekAttention();
				}
			});	
		},
		page : function(id, needTip, url){
			if (needTip){$("#commentBox").html("正在加载评论, 请稍后...");}
			var cJS = document.createElement("script");
			jQuery(cJS).attr("chatset", "utf-8").attr("src", url + "&page=" + id + "&s=" + Math.random());
			$("head:first").append(jQuery(cJS));
		},
		replyBox : function(id){
			var _id = id;
			try{$("#commentcontentreply_" + id).css("display", "block");}catch(e){}
			$.getJSON(
				"../pjblog.logic/log_comment.asp?action=replybox&s=" + Math.random(),
				{id : id},
				function(data){
					if (data.Suc){
						var div = document.createElement("div");
						var Str = cee.decode(data.Info);
						jQuery(div).html(Str).attr("id", "replyBox");
						if ($("#commcontent_" + _id)){
							$("#commcontent_" + _id).append(jQuery(div));
						}
						try{$("#commentcontentreply_" + _id).css("display", "none");}catch(e){}
					}else{}
				}
			);
		},
		replyBoxRemove : function(id){
			try{$("#replyBox").remove();$("#commentcontentreply_" + id).css("display", "block");}catch(e){}
		},
		doReply : function(id){
			var content = $("#replyBox_" + id).val(), _this = this, _id = id;
			$.post(
				"../pjblog.logic/log_comment.asp?action=savereply&s=" + Math.random(),
				{id : id, content : content},
				function(data){
					var json = data.json();
					if (json.Suc){
						effect.WarnTip.open({title : "恭喜 操作成功", html : "回复评论成功!"});
						try{
						  _this.replyBoxRemove();
						  $("#commentcontentreply_" + _id).css("display", "block");
						  $("#commentcontentreply_" + _id).html(cee.decode(json.Info));//这里改
					  }catch(e){alert(e.description)}
					}else{effect.WarnTip.open({title : "抱歉 操作失败", html : json.Info});}
				}
			);
		},
		Aduit : function(id, static, _obj){
			var _s = static, _this = this, _id = id, __obj = _obj;
			$.getJSON(
			  "../pjblog.logic/log_comment.asp?s=" + Math.random(),
			  {action : "Aduit", id : id, which : static},
			  function(data){
					if (data.Suc){
						var __this = _this, __id = _id;
						var ___obj = __obj;
						/*
							@ 1 : 通过审核
							@ 0 : 取消审核
						*/
						if (_s == 0){
							effect.WarnTip.open({title : "恭喜 操作成功", html : "<strong>取消</strong> 审核成功!"});
							jQuery(__obj).text("通过审核")
										 .bind("click", function(){
											ceeevio.Comment.Aduit(__id, 1, ___obj);
										 });
						}else{
							effect.WarnTip.open({title : "恭喜 操作成功", html : "<strong>通过</strong> 审核成功!"});
							jQuery(__obj).text("取消审核")
										 .bind("click", function(){
											ceeevio.Comment.Aduit(__id, 0, ___obj);
										 });
						}
					}else{
						effect.WarnTip.open({title : "错误信息", html : data.Info});
					}
			  }
			);
		},
		del : function(id){
			var _id = id;
			$.getJSON(
				"../pjblog.logic/log_comment.asp?s=" + Math.random(),
				{action : "del", id : id},
				function(data){
					if (data.Suc){
						effect.WarnTip.open({title : "恭喜 操作成功", html : "删除评论成功"});
						$("#comment_" + _id).remove();
					}else{alert(data.Info)}
				}
			);
		},
		quote : function(id, mem){
			$("#comm_Content").val("[quote=" + mem + "]" + $("#commcontent_" + id).text() + "[/quote]");
			location = "#postform";
		}
	}//,
}

var effect = {
	windows : {
		open : function(obj, str){
			
			var position, html, offset;

			position = str.position; html = str.html; offset = str.offset;
			
			var div = document.createElement("div"); $(document.body).append(div); jQuery(div).html(html);
			
			var el = this.position(obj, position, offset);

			jQuery(div).css({
							position	: 	"absolute",
							left 		: 	el.left,
							top 		: 	el.top
			});
			return div;
		},
		position : function(obj, i, o){
			var n = obj.offset(), h = obj.outerHeight(), w = obj.outerWidth();
			var l = 0, t = 0;
			switch (i){
				case 1 : 
					l = n.left; 			t = n.top - h - o; 				break;
				case 2 :
					l = n.left + w + o; 	t = n.top - h - o; 				break;
				case 3 :
					l = n.left + w + o; 	t = n.top; 						break;
				case 4 :
					l = n.left + w + o; 	t = n.top + h + o; 				break;
				case 5 :
					l = n.left; 			t = n.top + h + o;				break;
				case 6 :
					l = n.left - w - o;		t = n.top + h + o;				break;
				case 7 :
					l = n.left - w - o; 	t = n.top;						break;
				case 8 :
					l = n.left - w - o;		t = n.top - h - o;				break;
				default :
					l = n.left;				t = n.top;
			}
			return {left : l, top : t}
		}//,
	},
	WarnTip : {
		coint : null,
		open : function(str){
			try{this.close();}catch(e){}
			var loader_container = this.create("div", "loader_container");
			var loader = this.create("div", "loader");
			var load_title = this.create("div", "load_title");
			var load_body = this.create("div", "load_body");
			var loader_bg = this.create("div", "loader_bg");
			var progress = this.create("div", "progress");
			
			jQuery(load_title).css({textAlign : "center"});
			jQuery(load_body).css({textAlign : "center"});
			
			jQuery(load_title).html(str.title);
			jQuery(load_body).html(str.html);
			
			jQuery(loader_bg).append(jQuery(progress));
			jQuery(loader).append(jQuery(load_title));
			jQuery(loader).append(jQuery(load_body));
			jQuery(loader).append(jQuery(loader_bg));
			jQuery(loader_container).append(jQuery(loader));
			$(document.body).append(jQuery(loader_container));
			
			jQuery(loader_container).css(
			   "top",
			   (document.documentElement.clientHeight - jQuery(loader_container).outerHeight())/2 + document.documentElement.scrollTop + "px"
			);
			
			this.coint = setInterval(animate, 10);
			var _this = this;
			setTimeout(_this.close, 3000);
		},
		create : function(obj, id){
			var temp = document.createElement(obj);
			jQuery(temp).attr("id", id);
			return temp;
		},
		close : function(){
			clearInterval(this.coint);
			$("#loader_container").remove();
		}
	},
	MakeBox : {
		open : function(html){
			var div = document.createElement("div");
			$(document.body).append(jQuery(div)
				.css({position : "absolute", display : "block"})
				.html(html)
			);
			var height 	= 	jQuery(div).outerHeight();
			var width 	= 	jQuery(div).outerWidth();
			var top 	= 	(document.documentElement.clientHeight - height) / 2 + document.documentElement.scrollTop;
			var left 	= 	(document.documentElement.clientWidth - width) / 2 + document.documentElement.scrollLeft
			jQuery(div).css({
				top 	: 	top,
				left 	: 	left
			});
			return div;
		},
		PlusInsert : function(html){
			try{$("#PlusCode").remove();}catch(e){}
			var c = this.open(html);
			jQuery(c).css({
					border			: 		"2px solid #9EB3E0",
					background		: 		"#fff",
					padding 		:		"4px 4px 4px 4px",
					"font-size"		:		"12px"
			}).attr("id", "PlusCode").seekAttention();
		},
		PlusCode : function(id){
			var _id = id, _this = this;
			$.getJSON(
				"../pjblog.logic/control/log_template.asp?action=GetPlusCode&s=" + Math.random(),
				{id : id},
				function(data){
					if (data.Suc){
						var c = "";
						c += "<div style=\"width:500px;\"><div style=\"float:left; width:200px;\"><input type=\"button\" value=\"保存代码\" onclick=\"effect.MakeBox.SavePlusCode(" + _id + ")\" class=\"button\"><span id=\"saveinfo\"></span></div><div style=\"float:right; line-height:12px; width:12px; margin-top:6px\"><a href=\"javascript:;\" class=\"close\" onclick=\"try{$('#PlusCode').remove();}catch(e){}\">&nbsp;&nbsp;&nbsp;</a></div></div>";
						c += "<div class='clear'></div><textarea id=\"doSave\">" + cee.decode(data.Info) + "</textarea>";
						_this.PlusInsert(c);
					}else{
						alert(data.Info)
					}
				}
			);
		},
		SavePlusCode : function(id){
			var c = $("#doSave").val();
			$("#saveinfo").text("正在提交数据...").css({color : "red"});
			$.post(
				"../pjblog.logic/control/log_template.asp?action=SavePlusCode&id=" + id + "&s=" + Math.random(),
				{Content : c},
				function(data){
					var json = data.json();
					if (json.Suc){
						$("#saveinfo").text("保存数据成功...");
						setTimeout("try{$('#PlusCode').remove();}catch(e){}", 1000)
					}else{
						$("#saveinfo").text(json.Info);
					}
				}
			);
		},
		GetPlusTag : function(value){
			var c = "";
			c += "<div style=\"width:500px;\"><div style=\"float:left; width:400px;\">将以下代码插入到页面中即可</div><div style=\"float:right; line-height:12px; width:12px; margin-top:6px\"><a href=\"javascript:;\" class=\"close\" onclick=\"try{$('#PlusCode').remove();}catch(e){}\">&nbsp;&nbsp;&nbsp;</a></div></div>";
			c += "<div style=\" margin:4px 0px 8px 0px; text-align:center; width:500px\"><input type=\"text\" class=\"text\"value=\"" + value + "\" style=\"width:500px; margin:0 auto; font-family: 'Trebuchet MS', Arial, Helvetica, sans-serif;line-height:16px;font-style:oblique\" /></div>";
			c += "<div style=\"text-align:right; padding:0px 5px 0px 20px; color:#ccc; font-family:Arial, Helvetica, sans-serif; font-size:10px;font-style:oblique;width:485px\">CopyRight @ PJBlog4</div>";
			this.PlusInsert(c);
		}
	}//,
}

/*
	'	@	特效支持函数
*/
var Pos = 0, Dir = 2, len = 0;
function animate()
{
var elem = $('#progress');
	if(elem != null) {
		if (Pos == 0) len += Dir;
		if (len > 32 || Pos > 249) Pos += Dir;
		if (Pos > 249) len -= Dir;
		if (Pos > 249 && len == 0) Pos = 0;
		elem.css({
				 left 	: 	Pos,
				 width 	: 	len
		});
	}
}