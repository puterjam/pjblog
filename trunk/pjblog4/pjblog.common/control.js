// JavaScript Document
$(function(){
	$('a, input[type="button"]').bind('focus',function(){
		if(this.blur){ //如果支持 this.blur
			this.blur();
		}
	});
	var timer;
	try{
		for (var j = 0; j < 2 ; j++){
			$(".lay_" + j).hover(
				function(){
					clearTimeout(timer);
					$('#overT').eBoxClose();
					var ii = $(this).attr("class");
					var iii = parseInt(ii.replace("lay_", ""));
					$.eBoxClick(this, {
						css : {
							width : "172px",
							height : "60px",
							background : "url(../../images/Control/overlay_1.png) no-repeat"
						},
						n 	: 	10,
						html 	: 	ReadDom(iii),
						fix 	: 	false,
						id 	: 	"overT",
						oft 	: 	0,
						ofl 	: 	0
					});
				},
				function(){
					timer = setTimeout("$('#overT').eBoxClose()", 500);
				}
			);
		}
	}catch(e){}
	
	function ReadDom(j){
		return "<div class=\"ovlay-content\"><div class=\"con\">" + $("#ovlayinfo #lay_" + j).text() + "</div></div>";
	}
});

var conMain = {}

conMain.seesion = "";

conMain.tip = {
	edit : function(){
		conMain.seesion = $(".te").html();
		var word = $(".word").text();
		var c = "<form action=\"../pjblog.logic/control/log_general.asp?action=SaveTipWord\" method=\"post\" id=\"post_tip\">"
					+ "<input type=\"text\" value=\"" + word + "\" class=\"text\" style=\"width:235px;\" name=\"tip_word\" /><br />"
					+ "<a href=\"javascript:;\" onclick=\"conMain.tip.post()\">保存</a>&nbsp;&nbsp;<a href=\"javascript:;\" onclick=\"conMain.tip.close()\">取消</a>&nbsp;&nbsp;<span id=\"tip01\"></span>"
					+ "</form>";
		$(".te").html(c);
	},
	close : function(){
		$(".te").html(conMain.seesion);
	},
	post : function(){
		var c = $("input[name='tip_word']").val();
		$("#post_tip").ajaxSubmit({
			dataType : "json",
			beforeSubmit : function(){
				$("#tip01").text("wait, posting..").css({
					"font-style" : "oblique",
					"font-size" : "10px",
					"line-height" : "20px",
					color : "#ff0000"
				});
			},
			success : function(data){
				if (data.Suc){
					conMain.tip.close();
					$(".word").text(c);
				}else{
					jAlert(data.Info, "错误信息");
					$("#tip01").text("")
				}
			}			   
		});
	}
}

conMain.PersonSet = function(obj){
	try{$('#PersonSet').eBoxClose();}catch(e){}
	var c = "<div class=\"person_top\" onclick=\"$('#PersonSet').eBoxClose()\"></div>"
			+ "<div class=\"person_mid\"><div class=\"cContent\"><ul><li><a href=\"\">修改我的基本资料</a></li><li><a href=\"\">修改整站基本信息</a></li><li><a href=\"\">编辑网站基本权限</a></li></ul></div></div>"
			+ "<div class=\"person_bom\"></div>";
	$.eBoxClick(obj, {
		css : {
			background : "#ffffff",
			width : "176px"
		},
		n 		: 	11,
		html 	: 	c,
		fix 	: 	false,
		id 		: 	"PersonSet",
		oft 	: 	0,
		ofl 	: 	0
	});
}

conMain.Base = {
	seo : {
		session : null,
		ping : function(obj){
			var div = document.createElement("<tr><td>aaa</td></tr>")
			$(div).insertBefore($(obj).parent().parent());
		},
		Post : function(){
			$("#SEO").ajaxForm({
				dataType : "json",
				beforeSubmit : function(){
					$("input").attr("disabled", true);
				},
				success : function(data){
					var c = "";
					data.Suc ? c = "更新数据成功" : c = "更新数据失败";
					jAlert(c, '反馈信息');
					$("input").attr("disabled", false);
				}
			});
		},
		see : function(obj, html){
			try{$("#seePingUrl").eBoxClose();}catch(e){}
			$.eBoxClick(obj, {
				css : {
					background : "#ffffff",
					width : "250px",
					"line-height" : "16px",
					padding : "1px 10px 1px 10px",
					border : "1px solid #000",
					"font-style" : "oblique",
					"font-size" : "11px",
					cursor : "pointer",
					color : "#A61416"
				},
				n 		: 	9,
				html 	: 	html,
				fix 	: 	false,
				id 		: 	"seePingUrl",
				oft 	: 	0,
				ofl 	: 	0
			});
			$("#seePingUrl").click(function(){$("#seePingUrl").eBoxClose();});
		},
		editPing : function(id, url){
			try{this.cancelPing();}catch(e){}
			this.session = {id : id, html : $("#ping_" + id).html()};
			var c = "<td>"
					+ "<input type=\"text\" class=\"text\" value=\"" + $.trim($("#ping_" + id + " > td > span").eq(0).text()) + "\" style=\"width:130px\" id=\"Ping_Name\">&nbsp;&nbsp;"
					+ "<input type=\"text\" class=\"text\" value=\"" + $.trim(url) + "\" style=\"width:140px\" id=\"Ping_url\">"
					+ "</td><td><input type=\"button\" class=\"button\" value=\"更新\" onclick=\"conMain.Base.seo.SavePing(" + id + ")\">&nbsp;&nbsp;"
					+ "<input type=\"button\" class=\"button\" value=\"取消\" onclick=\"conMain.Base.seo.cancelPing()\">"
					+ "</td>";
			$("#ping_" + id).html(c);
		},
		cancelPing : function(){
			var obj = "#ping_" + this.session.id;
			$(obj).html(this.session.html);
		},
		SavePing : function(id){
			var _this = this, _id = id;
			var _Ping_Name = $("#Ping_Name").val();
			var _Ping_url = $("#Ping_url").val();
			$.getJSON(
				"../pjblog.logic/control/log_general.asp?s=" + Math.random(),
				{
					action : "savePing",
					id : id,
					Ping_Name : _Ping_Name,
					Ping_url : _Ping_url
				},
				function(data){
					if (data.Suc){
						_this.cancelPing();
						$("#ping_" + _id + " > td > span").eq(1).text(_Ping_Name);
						$("#ping_url_" + _id).val(_Ping_url);
					}else{
						jAlert(data.Info, '错误对话框');
					}
				}
			);
		}
	},
	BaseSetting : function(){
		$("#BaseSetting").ajaxForm({
			dataType : "json",
			beforeSubmit : function(){
				$("input").attr("disabled", true);
			},
			success : function(data){
				var c = "";
				data.Suc ? c = "更新数据成功" : c = "更新数据失败";
    			jAlert(c, '反馈信息');
				$("input").attr("disabled", false);
			}
		});	
	},
	ArtComm : function(){
		$("#ArtComm").ajaxForm({
			dataType : "json",
			beforeSubmit : function(){
				$("input").attr("disabled", true);
			},
			success : function(data){
				var c = "";
				data.Suc ? c = "更新数据成功" : c = "更新数据失败";
    			jAlert(c, '反馈信息');
				$("input").attr("disabled", false);
			}
		});
	},
	WebMode : function(){
		$("#WebMode").ajaxForm({
			dataType : "json",
			beforeSubmit : function(){
				$("input").attr("disabled", true);
			},
			success : function(data){
				var c = "";
				data.Suc ? c = "更新数据成功" : c = "更新数据失败";
    			jAlert(c, '反馈信息');
				$("input").attr("disabled", false);
			}
		});
	},
	WapMail : function(){
		$("#WapMail").ajaxForm({
			dataType : "json",
			beforeSubmit : function(){
				$("input").attr("disabled", true);
			},
			success : function(data){
				var c = "";
				data.Suc ? c = "更新数据成功" : c = "更新数据失败";
    			jAlert(c, '反馈信息');
				$("input").attr("disabled", false);
			}
		});
	},
	RegFilter : function(){
		$("#RegFilter").ajaxForm({
			dataType : "json",
			beforeSubmit : function(){
				$("input").attr("disabled", true);
			},
			success : function(data){
				var c = "";
				data.Suc ? c = "更新数据成功" : c = "更新数据失败";
    			jAlert(c, '反馈信息');
				$("input").attr("disabled", false);
			}
		});	
	}
}

conMain.initialize = {
	Get : function(){
		if ($("#s1").attr("checked")) this.s1();
		if ($("#s2").attr("checked")) this.s2();
		if ($("#s3").attr("checked")) this.s3();
		if ($("#s4").attr("checked")) this.s4();
	},
	s1 : function(){conMain.ReBuild.s1();},
	s2 : function(){conMain.ReBuild.s2();},
	s3 : function(){conMain.ReBuild.s3();},
	s4 : function(){conMain.ReBuild.s4();},
	conWidth : function(obj, delay, obj2, callBack){
		var step = 5;
		var parWidth = $(obj).parent().outerWidth();
		var selWidth = $(obj).outerWidth();
		var newStep = selWidth + step;
		if (newStep <= parWidth){
			$(obj).css({
				width : newStep + "px"		   
			});
			var c = parseInt((newStep /  parWidth) * 100);
			$(obj2 + " > td").eq(2).text(c + "%");
			setTimeout("conMain.initialize.conWidth('" + obj + "', " + delay + ", '" + obj2 + "', " + callBack + ")", delay);
		}else{callBack();}
	}
}

conMain.ReBuild = {}

conMain.ReBuild.s1 = function(){
	$("#s1").attr("disabled", true);
	$("#n1").css({
		width : "10%"
	});
	$.getJSON(
		"../pjblog.logic/control/log_general.asp?s=" + Math.random(),
		{
			action : "ReBuidWebData"
		},
		function(data){
			if (data.Suc){
				conMain.initialize.conWidth("#n1", 10, "#t1", function(){$("#s1").attr("disabled", false);});
			}
		}
	);
}

conMain.ReBuild.s2 = function(){
	$("#s2").attr("disabled", true);
	$("#n2").css({
		width : "10%"
	});
	$.getJSON(
		"../pjblog.logic/control/log_general.asp?s=" + Math.random(),
		{
			action : "ClearVistor"
		},
		function(data){
			if (data.Suc){
				conMain.initialize.conWidth("#n2", 10, "#t2", function(){$("#s2").attr("disabled", false)});
			}
		}
	);
}

conMain.ReBuild.s4 = function(){
	$("#s4").attr("disabled", true);
	$.get(
		"../pjblog.logic/control/log_general.asp?s=" + Math.random(),
		{
			action : "AppCount"
		},
		function(data){
			if (data > 0){
				conMain.ReBuild.app(0, data);
			}
		}
	);
}

conMain.ReBuild.app = function(index, count){
	if ((index + 1) <= count){
		var _index = index, _count = count;
		$.getJSON(
			"../pjblog.logic/control/log_general.asp?s=" + Math.random(),
			{
				action : "AppRemove",
				index : index
			},
			function(data){
				if (data.Suc){
					var c = parseInt(((_index + 1) / _count) * 100);
					$("#n4").css({
						width : c + "%"
					});
					$("#t4 > td").eq(2).text(c + "%");
					conMain.ReBuild.app((_index + 1), _count);
				}
			}
		);
	}else{$("#s4").attr("disabled", false);}
}

conMain.category = {
	add : function(obj){
		try{$('#cateAdd').eBoxClose();}catch(e){}
		$.eBoxClick(obj, {
			css : {
				width : "420px",
				border : "1px solid #7BBCF6",
				background : "#fff",
				padding : "10px 15px"
			},
			n 		: 	5,
			html 	: 	this.AddString(),
			fix 	: 	false,
			id 		: 	"cateAdd",
			oft 	: 	0,
			ofl 	: 	2,
			complete : function(){
				$("#postNewCate").ajaxForm({
					dataType : "json",
					beforeSubmit : function(){
						if ($("#cateAdd input[name='cate_Name']").val().length <= 0){
							jAlert("分类名称不能为空", "错误信息");
							return false;
						}
						$("#cateAdd input").attr("disabled", true);
						jAlert("正在发送数据,请稍后...", "确认信息");
					},
					success : function(data){
						if (data.Suc){
							jConfirm('新分类添加成功, 确定要刷新页面显示该分类吗?', '确认对话框', function(r){
								if (r){
									window.location.reload();
								}
							});
						}else{
							jAlert(data.Info, "错误信息");
						};
						$("#cateAdd input").attr("disabled", false);
					}			   
				});
			}
		});
	},
	AddString : function(){
		var c =   "<div class=\"title\">"
					+ "<div class=\"left\">增加新分类</div>"
					+ "<div class=\"right\" onclick=\"$('#cateAdd').eBoxClose();\">关闭</div>"
					+ "<div class=\"clear\"></div>"
				+ "</div>"
				+ "<form action=\"../pjblog.logic/control/log_category.asp?action=add\" method=\"post\" id=\"postNewCate\">"
				+ "<div class=\"mainBody\">"
					+ "<div>分类名称 <input type=\"text\" name=\"cate_Name\" class=\"text\" /></div>"
					+ "<div>分类说明 <input type=\"text\" name=\"cate_Intro\" class=\"text\" style=\"width:250px\" /></div>"
					+ "<div>静态目录 <input type=\"text\" name=\"cate_Folder\" class=\"text\" /> 在全静态模式下用到.</div>"
					+ "<div>分类排序 <input type=\"text\" name=\"cate_Order\" class=\"text\" value=\"0\" size=\"4\" /></div>"
					+ "<div>分类图标 <select name=\"cate_icon\">"
					+	"<option value=\"\">选择分类</option>"
					+ "</select></div>"
					+ "<div>外部链接 <input type=\"text\" name=\"cate_URL\" class=\"text\" style=\"width:250px\" /></div>"
					+ "<div>显示位置 <input type=\"radio\" name=\"cate_local\" value=\"0\" checked /> 同时 <input type=\"radio\" name=\"cate_local\" value=\"1\" /> 顶部 <input type=\"radio\" name=\"cate_local\" value=\"2\" /> 侧边</div>"
					+ "<div>分类私密 <input type=\"radio\" name=\"cate_Secret\" value=\"0\" checked /> 公开 <input type=\"radio\" name=\"cate_Secret\" value=\"1\" /> 隐藏</div>"
					+ "<div class=\"submit\"><input type=\"submit\" value=\"保存\" class=\"button\" /></div>"
				+ "</div>"
				+ "</form>";
		return c;
	},
	expert : function(obj){
		try{this.expertClose();}catch(e){}
		this.expertSession = $(obj);
		$(obj).addClass("lowBox");
		$.eBoxClick(obj, {
			css : {
				width : "420px",
				border : "1px solid #7BBCF6",
				background : "#fff",
				padding : "10px 15px",
				"z-index" : "0"
			},
			n 		: 	11,
			html 	: 	this.editString(),
			fix 	: 	false,
			id 		: 	"cateAdd",
			oft 	: 	-1,
			ofl 	: 	0,
			complete : function(){
				$("#postNewCate").ajaxForm({
					dataType : "json",
					beforeSubmit : function(){
						if ($("#cateAdd input[name='cate_Name']").val().length <= 0){
							jAlert("分类名称不能为空", "错误信息");
							return false;
						}
						$("#cateAdd input").attr("disabled", true);
						jAlert("正在发送数据,请稍后...", "确认信息");
					},
					success : function(data){
						if (data.Suc){
							jConfirm('编辑分类成功, 确定要刷新页面重载该分类信息吗?', '确认对话框', function(r){
								if (r){
									window.location.reload();
								}
							});
						}else{
							jAlert(data.Info, "错误信息");
						};
						$("#cateAdd input").attr("disabled", false);
					}			   
				});
			}
		});
	},
	expertSession : null,
	expertClose : function(){
		$("#cateAdd").eBoxClose();
		this.expertSession.removeClass("lowBox");
	},
	editString : function(){
		var x = this.expertSession.parents("tr:first").prev().find("td > .hideData > div");
		var c =   "<div class=\"title\">"
					+ "<div class=\"left\">编辑分类信息</div>"
					+ "<div class=\"right\" onclick=\"conMain.category.expertClose()\">关闭</div>"
					+ "<div class=\"clear\"></div>"
				+ "</div>"
				+ "<form action=\"../pjblog.logic/control/log_category.asp?action=edit\" method=\"post\" id=\"postNewCate\"><input type=\"hidden\" name=\"cate_ID\" value=\"" + $.trim(x.eq(0).text()) + "\">"
				+ "<div class=\"mainBody\">"
					+ "<div>分类名称 <input type=\"text\" name=\"cate_Name\" class=\"text\" value=\"" + $.trim(x.eq(3).text()) + "\" /></div>"
					+ "<div>分类说明 <input type=\"text\" name=\"cate_Intro\" class=\"text\" style=\"width:250px\" value=\"" + $.trim(x.eq(4).text()) + "\" /></div>"
					+ "<div>静态目录 <input type=\"text\" name=\"cate_Folder\" class=\"text\" value=\"" + $.trim(x.eq(5).text()) + "\" /> 在全静态模式下用到.</div>"
					+ "<div>分类排序 <input type=\"text\" name=\"cate_Order\" class=\"text\" value=\"" + $.trim(x.eq(1).text()) + "\" size=\"4\" /></div>"
					+ "<div>分类图标 <select name=\"cate_icon\">"
					+	"<option value=\"\">选择分类</option>"
					+ "</select></div>"
					+ "<div>外部链接 <input type=\"text\" name=\"cate_URL\" class=\"text\" style=\"width:250px\" value=\"" + $.trim(x.eq(6).text()) + "\" /></div>"
					+ "<div>显示位置 <input type=\"radio\" name=\"cate_local\" value=\"0\" " + (parseInt($.trim(x.eq(7).text())) == 0 ? "checked" : "") + " /> 同时 <input type=\"radio\" name=\"cate_local\" value=\"1\" " + (parseInt($.trim(x.eq(7).text())) == 1 ? "checked" : "") + " /> 顶部 <input type=\"radio\" name=\"cate_local\" value=\"2\" " + (parseInt($.trim(x.eq(7).text())) == 2 ? "checked" : "") + " /> 侧边</div>"
					+ "<div>分类私密 <input type=\"radio\" name=\"cate_Secret\" value=\"0\" " + ($.trim(x.eq(8).text()).toLowerCase() == "false" ? "checked" : "") + "/> 公开 <input type=\"radio\" name=\"cate_Secret\" value=\"1\" " + ($.trim(x.eq(8).text()).toLowerCase() == "true" ? "checked" : "") + " /> 隐藏</div>"
					+ "<div class=\"submit\"><input type=\"submit\" value=\"保存\" class=\"button\" /></div>"
				+ "</div>"
				+ "</form>";
		return c;
	},
	del : function(id){
		jConfirm('确定要删除此分类?<br />如果是日志分类链接, 删除后连同该分类下的日志全部删除<br />', '确认对话框', function(r){
			if (r){
				$.post(
					"../pjblog.logic/control/log_category.asp?action=del&s=" + Math.random(),
					{SelectID : id},
					function(data){
						var json = data.json();
						if (json.Suc){
							jConfirm("删除分类成功!<br />是否要刷新本页?", "确认对话框", function(r){
								if (r) window.location.reload();												
							})
						}else{
							jAlert(json.Info, "错误信息");
						}
					}
				);
			}
		});
	}
}

conMain.skin = {}

conMain.skin.set = {
	read : {
		html : function(obj, s1, s2, s3){
			var info = $(obj).find(".readinfo div");
			var styleinfo;
			var m = "";
			if (s1 == s3) m = s2;
			var c =   "<div class=\"title\" id=\"SkinSetTitle\" title=\"点击此处拖动本框\">"
					+ 	"<span class=\"left\">查看主题</span>"
					+ 	"<span class=\"right\" onclick=\"$('#cateAdd').eBoxClose();\">关闭</span>"
					+ 	"<div class=\"clear\"></div>"
					+ "</div>"
					+ "<div class=\"mainBody\">"
					+	"<div><strong>主题名称</strong> " + $.trim(info.eq(0).text()) + "</div>"
					+	"<div><strong>主题作者</strong> " + $.trim(info.eq(1).text()) + "</div>"
					+	"<div><strong>发布时间</strong> " + $.trim(info.eq(2).text()) + "</div>"
					+	"<div><strong>作者网站</strong> " + $.trim(info.eq(3).text()) + " <a href=\"" + $.trim(info.eq(3).text()) + "\" target=\"_blank\" style=\"margin-left:10px;\">访问</a></div>"
					+	"<div><strong>作者邮箱</strong> " + $.trim(info.eq(4).text()) + " <a href=\"mailto:" + $.trim(info.eq(4).text()) + "\" style=\"margin-left:10px;\">发邮件给他</a></div>"
					+	"<div><strong>主题版本</strong> " + $.trim(info.eq(5).text()) + "</div>"
					+	"<div><strong>主题说明</strong> " + $.trim(info.eq(8).text()) + "</div>"
					+ 	"<div class=\"submit\"><form action=\"../pjblog.logic/control/log_template.asp?action=update\"  method=\"post\" id=\"sinset\">"
					+	"<strong>样式选择:</strong><input type=\"hidden\" value=\"" + m + "\" name=\"skin_style\"><input type=\"hidden\" value=\"" + s3 + "\" name=\"skin_common\">"
				   	+	"<div class=\"style\">";
				$(obj).find(".readStyle .item").each(function(){
					styleinfo = $(this).find("div");
					c += "<li><img src=\"" + $.trim(styleinfo.eq(6).text()) + "\" width=\"40\" height=\"40\" title=\"" + $.trim(styleinfo.eq(0).text()) + "\" class=\"" + (((s1 == s3)&&($.trim(styleinfo.eq(9).text()) == s2)) ? "selectimg" : "cimg") + "\" rel=\"" + $.trim(styleinfo.eq(9).text()) + "\" onerror=\"this.src='../images/control/sPreview.jpg'\"/></li>"					  
				});					
				c	+=	"<div class=\"clear\"></div></div>"
					+	"<input type=\"submit\" value=\"确定\" class=\"button\" />"
					+	"</form></div>"
					+ "</div>"
			return c;
		},
		open : function(obj, s1, s2, s3){
			// s1 : 默认主题文件夹
			// s2 : 默认样式文件夹
			// s3 : 当前主题的文件夹
			try{$("#cateAdd").eBoxClose();}catch(e){}
			$.eBoxClick(obj, {
				css : {
					width : "420px",
					border : "1px solid #7BBCF6",
					background : "#fff",
					padding : "10px 15px"
				},
				html 	: 	this.html(obj, s1, s2, s3),
				fix 	: 	true,
				id 		: 	"cateAdd",
				oft 	: 	0,
				ofl 	: 	0,
				complete : function(){
					$('#cateAdd').draggable({handle : "#SkinSetTitle"});
					$(".cimg").click(function(){
						try{
							$(".selectimg").attr("class", "cimg");
						}catch(e){}				  
						$(this).attr("class", "selectimg");
						var c = $(this).attr("rel");
						$("input[name='skin_style']").val(c);
					});
					$("#sinset").ajaxForm({
						dataType : "json",
						beforeSubmit : function(){
							if ($("#sinset input[name='skin_style']").val().length <= 0){
								jAlert("分类名称不能为空", "错误信息");
								return false;
							}
							if ($("#sinset input[name='skin_common']").val().length <= 0){
								jAlert("分类名称不能为空", "错误信息");
								return false;
							}
							$("#sinset input").attr("disabled", true);
							jAlert("正在发送数据,请稍后...", "确认信息");
						},
						success : function(data){
							if (data.Suc){
								jConfirm('激活主题和样式成功, 是否要关闭本样式列表, 并且刷新数据?', '确认对话框', function(r){
									if (r){
										window.location.reload();
									}
								});
							}else{
								jAlert(data.Info, "错误信息");
							};
							$("#sinset input").attr("disabled", false);
						}				  
					});
				}
			});
		}
	},
	del : function(f){
		jConfirm("删除主题后将无法恢复.<br />确定这样做吗?", "确认对话框", function(r){
			if (r){
				$.getJSON(
					"../pjblog.logic/control/log_template.asp?action=DelTheme&s=" + Math.random(), 
					{
						fo : f
					},
					function(data){
						if (data.Suc){
							jConfirm("删除主题成功, 是否要刷新数据?", "确认对话框", function(r){
								if (r) window.location.reload();										   
							})
						}else{
							jAlert(data.Info, "错误对话框");
						}
					}
				);
			}
		});
	}
}