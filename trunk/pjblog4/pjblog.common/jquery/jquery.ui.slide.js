// JavaScript Document
/*	
@	Dom :
	<ul id="t1">
    	<li>测试1</li>
        <li>测试2</li>
        <li>测试3</li>
    </ul>
    <ul id="t2">
    	<li>1<b>dsaf</b></li>
        <li>2</li>
        <li>3</li>
    </ul>

@	How To Use :
	$(function(){
		$("#t1").tabs("#t2");   
	});
*/
;(function($){
	$.fn.tabs = function(div, options){
		var _this = this;
		var tabsArray = {}
		
		options = $.extend({
			current : "current"
		}, options || {});
		
		_this.find("li:first").addClass(options.current);
		$(div).find("li").each(function(i){
			$(this).attr("id", "e_" + i);
		});
		$(div).find("li:not(:first)").hide();
		_this.find("li").click(function(){
			tabsArray.show(this);							   
		});
		tabsArray.show = function(obj){
			var c = _this.find("li").index($(obj));
			$(div).find("#e_" + c).siblings().hide();
         	$(div).find("#e_" + c).show();
			$(obj).addClass(options.current).siblings().removeClass(options.current);
		}
	}   
})(jQuery);