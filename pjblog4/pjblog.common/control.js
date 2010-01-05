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
	}
}

conMain.initialize = {
	Get : function(){
		$("input").attr("disabled", true);
		if ($("#s1").attr("checked")) this.s1();
		if ($("#s2").attr("checked")) this.s2();
		if ($("#s3").attr("checked")) this.s3();
		if ($("#s4").attr("checked")) this.s4();
	},
	s1 : function(){conMain.ReBuild.s1();},
	s2 : function(){conMain.ReBuild.s2();},
	s3 : function(){conMain.ReBuild.s3();},
	s4 : function(){conMain.ReBuild.s4();},
	conWidth : function(obj, delay, obj2){
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
			setTimeout("conMain.initialize.conWidth('" + obj + "', " + delay + ", '" + obj2 + "')", delay);
		}
	}
}

conMain.ReBuild = {}

conMain.ReBuild.s1 = function(){
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
				conMain.initialize.conWidth("#n1", 10, "#t1");
			}
		}
	);
}

conMain.ReBuild.s2 = function(){
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
				conMain.initialize.conWidth("#n2", 10, "#t2");
			}
		}
	);
}

conMain.ReBuild.s4 = function(){
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
	}
}