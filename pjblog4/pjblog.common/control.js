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