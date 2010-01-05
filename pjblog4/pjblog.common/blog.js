// JavaScript Document
var blog = {}

blog.post = function(obj){
	var _this = this, _obj = obj;
	$.get(
		"../pjblog.express/post.asp?s=" + Math.random(),
		function(data){
			_this.add(_obj, data);
		}
	);
}

function active(){
	$("input[type='text']").focus(function(){
		$(this).css("border", "1px solid #fc5d71");
	});	
}

blog.close = function(){
	try{$("#cateSelectBox").remove();}catch(e){}
	$('#blog').eBoxClose();
}

blog.add = function(obj, html){
	$.eBoxClick(obj, {
		css : {
			width : "600px"
		},
		html 	: 	html,
		fix 	: 	true,
		id 		: 	"blog"
	});
	$("#blog").draggable({handle : ".box-title"});
	var t = $(".xml > .item").size();
	if ($(".xml > .item").size() > 0){
		$("#cate-select").mouseover(function(){
			try{$("#cateSelectBox").remove();}catch(e){}
			var c = "<div class=\"cate-doTop\"></div><div class=\"cate-doContent\"><ul class=\"cate-doSel\">";
			$(".xml > .item").each(function(){
				c += "<li><div class=\"cate-li\">" + $(this).find(".icon").html() + $(this).find(".txt").html() + "</div></li>";
			})
				c += "</ul></div><div class=\"cate-doBottom\"></div>";
			$.eBoxClick(this, {
				css : {
					width : "151px"
				},
				n 	: 	5,
				html 	: 	c,
				fix 	: 	false,
				id 	: 	"cateSelectBox",
				oft 	: 	0,
				ofl 	: 	1
			});								 
		});
	}
}