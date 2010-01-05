// JavaScript Document
/*
$.eBoxClick(obj, {
	css : {
		background : "#ff0000",
		width:"100px",
		height:"100px",
		"z-index" : 99
	},
	html:"evio",
	fix : true,
	-------------- clear -------------
	n : 0,
	ofl : 0,
	oft : 0
});
*/	
	
;(function($){
	$.eBoxClick = function(obj, options){

		var div = document.createElement("div");
		$("body").append(div);
		
		options = $.extend({
			id		: 	null,
			n		: 	null,
			ofl 	: 	null,
			oft 	: 	null,
			css 	: 	null,
			html 	: 	null,
			fix 	: 	true
		}, options || {});
		
		$(div).html(options.html).css(options.css);

		if (options.id != null) $(div).attr("id", options.id);
		
		var off = $(obj).offset();
		
		var c = {
			n 	:	options.n,
			ol 	: 	off.left,
			ot 	: 	off.top,
			ow 	: 	$(obj).outerWidth(),
			oh 	: 	$(obj).outerHeight(),
			dw 	: 	$(div).outerWidth(),
			dh 	: 	$(div).outerHeight(),
			ofl : 	options.ofl,
			oft : 	options.oft
		}

		var pos = getPostion(c);

		var left = pos.left;
		var top = pos.top;

		$(div).css({
			left : left + "px",
			top : top + "px"
		});
		
		if (options.fix){
			$(div).css("position", "fixed");
		}else{
			$(div).css("position", "absolute");
		}
		//$(div).fadeOut(0).fadeIn(600);
	}
/*
		option : options
		{
			n :  0 - 16,
			ol : object left,
			ot : object top,
			ow : object width,
			oh : object height,
			dw : div width,
			dh : div height,
			ofl : div offset left,
			oft : div offset top
		}
*/
	function getPostion(option){
		var left, top;
		var n 	= 	option.n;
		var ol 	= 	option.ol;
		var ot 	= 	option.ot;
		var ow 	= 	option.ow;
		var oh 	= 	option.oh;
		var dw 	= 	option.dw;
		var dh 	= 	option.dh;
		var ofl = 	option.ofl;
		var oft = 	option.oft;
		if (ofl == null) ofl = 0;
		if (oft == null) oft = 0;
		switch (n){
			case 0 :
				left = ol + ofl;
				top = ot + oft;
				break;
			case 1 :
				left = ol + ofl;
				top = ot - dh + oft;
				break;
			case 2 :
				left = ol + parseInt((ow - dw) / 2) + ofl;
				top = ot - dh + oft;
				break;
			case 3 :
				left = ol + (ow - dw) + ofl;
				top = ot - dh + oft;
				break;
			case 4 :
				left = ol + ow + ofl;
				top = ot - dh + oft;
				break;
			case 5 : 
				left = ol + ow + ofl;
				top = ot + oft;
				break;
			case 6 :
				left = ol + ow + ofl;
				top = ot + parseInt((oh - dh) / 2) + oft;
				break;
			case 7 :
				left = ol + ow + ofl;
				top = ot + (oh - dh) + oft;
				break;
			case 8 :
				left = ol + ow + ofl;
				top = ot + oh + oft;
				break;
			case 9 :
				left = ol + (ow - dw) + ofl;
				top = ot + oh + oft;
				break;
			case 10 :
				left = ol + parseInt((ow - dw) / 2) + ofl;
				top = ot + oh + oft;
				break;
			case 11 :
				left = ol + ofl;
				top = ot + oh + oft;
				break;
			case 12 :
				left = ol - dw + ofl;
				top = ot + oh + oft;
				break;
			case 13 :
				left = ol - dw + ofl;
				top = ot + (oh - dw) + oft;
				break;
			case 14 :
				left = ol - dw + ofl;
				top = ot + parseInt((oh - dw) / 2) + oft;
				break;
			case 15 :
			 	left = ol - dw + ofl;
				top = ot + oft;
				break;
			case 16 :
				left = ol - dw + ofl;
				top = ot - dh + oft;
				break;
			default : 
				left = parseInt(parseInt($(window).width()/ 2) - parseInt(dw / 2));
				top = parseInt(parseInt($(window).height()/ 2) - parseInt(dh / 2));
		}
		
		return {left : left, top : top};
	}
	
	$.fn.eBoxClose = function(){
		this.remove();
	}
	
	$.fn.rePostion = function(){
		var width = this.outerWidth();
		var height = this.outerHeight();
		var ps = getPostion({
			dw : width,
			dh : height
		});
		this.css({
			left : ps.left + "px",
			top : ps.top + "px"
		});
	}
})(jQuery);