// JavaScript Document
;jQuery.fn.extend({
	setDomEdit : function(){$(this).attr({designMode : "on", contentEditable : true});},
	setDomUnEdit : function(){$(this).attr({designMode : "off", contentEditable : false});},
	setAllDomEdit : function(){$(this).find("*[name]").setDomEdit();},
	setAllDomUnEdit : function(){$(this).find("*[name]").setDomUnEdit();},
	setDomFormClose : function(){$(this).unbind("click")},
	setDomForm : function(options){
		options = $.extend({
			dataType : "json",
			target : null,
			beforeSubmit : function(){return true},
			success : function(){},
			submitDomObject : null,
			AltTitle : null
		}, options || {});
		var _this = this;
		$(this).find("*[name]").setDomEdit();
		$(this).find("*[name]").mouseover(function(){
			$(this).attr("title", options.AltTitle)
			.focus();								   
		});
		options.submitDomObject.bind("click", function(){
			if (options.beforeSubmit()){
				var c = new Array(), s = new Array(), UrlStr = "";
				var left, right;
				$(_this).find("*[name]").each(function(){
					left = $.trim($(this).attr("name"));
					right = $.trim($(this).text());
					if (left.length > 0){c.push(left); s.push(right);}
				});
				for (var i = 0 ; i < c.length ; i++){
					UrlStr += (c[i] + "=" + s[i] + "&");
				}
				UrlStr = UrlStr.substring(0, UrlStr.lastIndexOf("&"));
				$.ajax({
				   type : "post",
				   url : options.target,
				   dataType : options.dataType,
				   data : UrlStr,
				   success : options.success
				});
			}						   
		});
	}
})