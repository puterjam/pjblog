// JavaScript Document
String.prototype.json = function(){
	try{
		eval("var jsonStr = (" + this.toString() + ")");
	}catch(ex){
		var jsonStr = null;
	}
	return jsonStr;
};
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
			width : "610px",
			border : "1px solid #7BBCF6",
			background : "#fff",
			padding : "10px 15px"
		},
		html 	: 	html,
		fix 	: 	true,
		id 		: 	"cateAdd",
		oft 	: 	0,
		ofl 	: 	0,
		complete : function(){
			$('#cateAdd').draggable({handle : "#blogPost"});
			$("#blogSlide").tabs("#blogSlideContent");
			$("textarea[name='Message']").ubbEditor();
		}
	});
}