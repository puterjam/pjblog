// JavaScript Document
// author : evio http://www.evio.name
var NodeForEvioBject;

;(function($){
	$.fn.ubbEditor = function(options){
		
		// 被绑定元素 $(this) || this
			NodeForEvioBject = $(this)
		// 继承元素属性
		options = $.extend({
			 toolBar : "1,2,3,4,5,6",
			 offsetLeft : 0,
			 iframeOffsetTop : 0
		}, options || {});
		
		var objects = {
			objWidth : 0,
			objHeight : 0,
			objOffsetLeft : 0,
			objOffsetTop : 0
		}
		
		// 获取被当定元素的长宽数据
		objects.thisData = function(o){
			this.objWidth = $(o).outerWidth();
			this.objHeight = $(o).outerHeight();
			this.objOffsetLeft = $(o).offset().left;
			this.objOffsetLeft = $(o).offset().top;
		}
		
		// font size ||  color left center right blod underline oblique delete ||
		// url mail list quote code html down || edk mp3 real || smile || about
		// insert : options.toolBar
		objects.toolBarItem = function(str){
			var c = "";
			var arr = str.split(",");
			for (var i = 0 ; i < arr.length ; i++){
				c += getItem(parseInt(arr[i]));
			}
			return c;
		}
		
		// insert : options.toolBar
		// obj : 被绑定的元素 $(this)
		objects.load = function(obj, str){
			objects.thisData(obj);
			var c = this.toolBarItem(str)
			var div = document.createElement("div");
			$(div).attr("class", "UBBeditor").html(c).css({
				width	:	"100%"
			}).insertBefore($(obj)).append("<div class=\"clear\"></div>");
			
			//  加载iframe
			var iframe = document.createElement("iframe");
			$(iframe)
			.attr({
				marginWidth : 0,
				marginHeight : 0,
				id : "HTMLEditor",
				src : ""
			})
			.css({
				width : this.objWidth + "px",
				height : this.objHeight + "px"
			}).insertAfter($(obj));
			
			$(iframe).ready(function(){
				var iObj = document.getElementById("HTMLEditor").contentWindow; 
				iObj.document.designMode = "on"; 
				iObj.document.contentEditable = true; 
				iObj.document.open(); 
				iObj.document.writeln("<html><head>");
				iObj.document.writeln("</head><body contenteditable='true'></body></html>"); 
				iObj.document.close();
				//$("#HTMLEditor", parent.document.body).attr("contenteditable", "true")
			}) 

			//if ($(div).outerWidth() > this.objWidth) $(div).css("width", (objects.objWidth + options.offsetLeft)+"px");
		}	
/*
		处理代码
*/
		objects.load(this, options.toolBar);
		$(this).setCaret();
		$(".UBBeditor > .bar:last").css("background", "none");
		$(this).click(function(){
			try{
				$(".float").hide('fast');
			}catch(e){}
		});
		
		//objects.thisData(this);
		//alert(options.toolBar)
		
		function getItem(str){
			var c = "";
			switch (str){
				case 1 :
					c = "<div class=\"child_1 bar\">"
					  + 	"<div class=\"font baritem\" title=\"字体\"></div>"
					  + 	"<div class=\"size baritem\" title=\"文字大小\"></div>"
					  + 	"<div class=\"clear\"></div>"
					  + "</div>";
					break;
				case 2 :
					c = "<div class=\"child_2 bar\">"
					  + 	"<div class=\"color baritem\" title=\"颜色\">"
					  +			"<div onclick=\"$(this).ubbEditorClick('color');\" class=\"itemClick\" title=\"颜色\"></div>"
					  + 	"</div>"
					  + 	"<div class=\"left baritem\" onclick=\"$(this).ubbEditorClick('left');\" title=\"居左\"></div>"
					  + 	"<div class=\"center baritem\" onclick=\"$(this).ubbEditorClick('center');\" title=\"居中\"></div>"
					  + 	"<div class=\"right baritem\" onclick=\"$(this).ubbEditorClick('right');\" title=\"居右\"></div>"
					  + 	"<div class=\"blod baritem\" onclick=\"$(this).ubbEditorClick('font');\" title=\"加粗\"></div>"
					  + 	"<div class=\"underline baritem\" onclick=\"$(this).ubbEditorClick('underline');\" title=\"下划线\"></div>"
					  + 	"<div class=\"oblique baritem\" onclick=\"$(this).ubbEditorClick('oblique');\" title=\"斜体\"></div>"
					  + 	"<div class=\"delete baritem\" onclick=\"$(this).ubbEditorClick('deleteline');\" title=\"删除线\"></div>"
					  + 	"<div class=\"clear\"></div>"
					  + "</div>";
					break;
				case 3 :
					c = "<div class=\"child_3 bar\">"
					  + 	"<div class=\"url baritem\">"
					  +			"<div onclick=\"$(this).ubbEditorClick('url');\" class=\"itemClick\" title=\"插链接\"></div>"
					  +		"</div>"
					  + 	"<div class=\"mail baritem\">"
					  +			"<div onclick=\"$(this).ubbEditorClick('mail');\" class=\"itemClick\" title=\"邮件\"></div>"
					  +		"</div>"
					  + 	"<div class=\"list baritem\"></div>"
					  + 	"<div class=\"quote baritem\">"
					  +			"<div onclick=\"$(this).ubbEditorClick('quote');\" class=\"itemClick\" title=\"引用\"></div>"
					  +		"</div>"
					  + 	"<div class=\"code baritem\">"
					  +			"<div onclick=\"$(this).ubbEditorClick('code');\" class=\"itemClick\" title=\"插入代码\"></div>"
					  +		"</div>"
					  + 	"<div class=\"html baritem\">"
					  +			"<div onclick=\"$(this).ubbEditorClick('html');\" class=\"itemClick\" title=\"运行html代码\"></div>"
					  +		"</div>"
					  + 	"<div class=\"down baritem\">"
					  +			"<div onclick=\"$(this).ubbEditorClick('down');\" class=\"itemClick\" title=\"插入待下载代码\"></div>"
					  +		"</div>"
					  + 	"<div class=\"clear\"></div>"
					  + "</div>";
					break;
				case 4 :
					c = "<div class=\"child_4 bar\">"
					  + 	"<div class=\"edk baritem\"></div>"
					  + 	"<div class=\"mp3 baritem\"></div>"
					  + 	"<div class=\"real baritem\"></div>"
					  + 	"<div class=\"clear\"></div>"
					  + "</div>";
					break;
				case 5 :
					c = "<div class=\"child_5 bar\">" 
					  + 	"<div class=\"smile baritem\"></div>"
					  + 	"<div class=\"clear\"></div>"
					  + "</div>";
					break;
				case 6 :
					c = "<div class=\"child_6 bar\">"
					  + 	"<div class=\"about baritem\"></div>"
					  + 	"<div class=\"clear\"></div>"
					  + "</div>";
					break;
			}
			return c;
		}
	}
	
	$.fn.ubbEditorClick = function(fn){
		switch (fn){
			case "font" :
				outTextRange.font(); break;
			case "left" :
				outTextRange.left(); break;
			case "center" :
				outTextRange.center(); break;
			case "right" :
				outTextRange.right(); break;
			case "underline" :
				outTextRange.underline(); break;
			case "oblique" :
				outTextRange.oblique(); break;
			case "deleteline" :
				outTextRange.deleteline(); break;
			case "url" :
				outTextRange.url($(this)); break;
			case "mail" :
				outTextRange.mail($(this)); break;
			case "quote" :
				outTextRange.quote($(this)); break;	
			case "code" :
				outTextRange.code($(this)); break;
			case "html" :
				outTextRange.html($(this)); break;
			case "down" :
				outTextRange.down($(this)); break;
			case "color" :
				outTextRange.color($(this)); break;
			default :
				""
		}
	}
	
	var outTextRange = {
		ovelayBox : function(str){
			var c = "<div class=\"ubbOverlay_wrap\">"
				  + 	"<div class=\"ubbOverlay_inner\">"
				  +		"<div class=\"outWrap\">" + str + "</div>"
				  + 	"</div>"
				  +	"</div>";
				  return c;
		},
		font : function(){
			if (NodeForEvioBject.cloneTextRange().length > 0){NodeForEvioBject.insertAtCaret("[b]{$}[/b]", true);}
		},
		left : function(){
			if (NodeForEvioBject.cloneTextRange().length > 0){NodeForEvioBject.insertAtCaret("[align=left]{$}[/align]", true);}
		},
		center : function(){
			if (NodeForEvioBject.cloneTextRange().length > 0){NodeForEvioBject.insertAtCaret("[align=center]{$}[/align]", true);}
		},
		right : function(){
			if (NodeForEvioBject.cloneTextRange().length > 0){NodeForEvioBject.insertAtCaret("[align=right]{$}[/align]", true);}
		},
		underline : function(){
			if (NodeForEvioBject.cloneTextRange().length > 0){NodeForEvioBject.insertAtCaret("[u]{$}[/u]", true);}
		},
		oblique : function(){
			if (NodeForEvioBject.cloneTextRange().length > 0){NodeForEvioBject.insertAtCaret("[i]{$}[/i]", true);}
		},
		deleteline : function(){
			if (NodeForEvioBject.cloneTextRange().length > 0){NodeForEvioBject.insertAtCaret("[s]{$}[/s]", true);}
		},
		url : function(obj){
			if (NodeForEvioBject.cloneTextRange().length > 0){
				NodeForEvioBject.insertAtCaret("[url]{$}[/url]", true);
			}else{
				try{$(".float").remove();}catch(e){}
				var div = "<div class=\"float\">" + this.ovelayBox(textOutString.url()) + "</div>";
				$(div).insertAfter(obj);
				$(".float").slideToggle('fast');
				$("#url_bnt").click(function(){
					$(this).attr("disabled", true);
					var urlname = $("#urlname").val();
					var url = $("#url").val();
					var c;
					if (urlname.length > 0){
						c = "[url=" + url + "]" + urlname + "[/url]";
					}else{
						c = "[url]" + url + "[/url]";
					}
					$(this).attr("disabled", false);
					NodeForEvioBject.insertAtCaret(c, false);
					$(".float").hide('fast');
				});
			}
		},
		mail : function(obj){
			if (NodeForEvioBject.cloneTextRange().length > 0){
				NodeForEvioBject.insertAtCaret("[mail]{$}[/mail]", true);
			}else{
				try{$(".float").remove();}catch(e){}
				var div = "<div class=\"float\">" + this.ovelayBox(textOutString.mail()) + "</div>";
				$(div).insertAfter(obj);
				$(".float").slideToggle('fast');
				$("#mail_bnt").click(function(){
					$(this).attr("disabled", true);
					var mailname = $("#mailname").val();
					var mail = $("#mail").val();
					var c;
					if (mailname.length > 0){
						c = "[mail=" + mail + "]" + mailname + "[/mail]";
					}else{
						c = "[mail]" + mail + "[/mail]";
					}
					$(this).attr("disabled", false);
					NodeForEvioBject.insertAtCaret(c, false);
					$(".float").hide('fast');
				});
			}
		},
		quote : function(obj){
			if (NodeForEvioBject.cloneTextRange().length > 0){
				NodeForEvioBject.insertAtCaret("[quote]{$}[/quote]", true);
			}else{
				try{$(".float").remove();}catch(e){}
				var div = "<div class=\"float\">" + this.ovelayBox(textOutString.quote()) + "</div>";
				$(div).insertAfter(obj);
				$(".float").slideToggle('fast');
				$("#quote_bnt").click(function(){
					$(this).attr("disabled", true);
					var quotename = $("#quotename").val();
					var quote = $("#quote").val();
					var c;
					if (quotename.length > 0){
						c = "[quote=" + quotename + "]" + quote + "[/quote]";
					}else{
						c = "[quote]" + quote + "[/quote]";
					}
					$(this).attr("disabled", false);
					NodeForEvioBject.insertAtCaret(c, false);
					$(".float").hide('fast');
				});
			}
		},
		code : function(obj){
			if (NodeForEvioBject.cloneTextRange().length > 0){
				NodeForEvioBject.insertAtCaret("[code]{$}[/code]", true);
			}else{
				try{$(".float").remove();}catch(e){}
				var div = "<div class=\"float\">" + this.ovelayBox(textOutString.code()) + "</div>";
				$(div).insertAfter(obj);
				$(".float").slideToggle('fast');
				$("#code_bnt").click(function(){
					$(this).attr("disabled", true);
					var codename = $("#codename").val();
					var code = $("#code").val();
					var c;
					if (codename.length > 0){
						c = "[code=" + codename + "]" + code + "[/code]";
					}else{
						c = "[code]" + code + "[/code]";
					}
					$(this).attr("disabled", false);
					NodeForEvioBject.insertAtCaret(c, false);
					$(".float").hide('fast');
				});
			}
		},
		html : function(obj){
			if (NodeForEvioBject.cloneTextRange().length > 0){
				NodeForEvioBject.insertAtCaret("[html]{$}[/html]", true);
			}else{
				try{$(".float").remove();}catch(e){}
				var div = "<div class=\"float\">" + this.ovelayBox(textOutString.html()) + "</div>";
				$(div).insertAfter(obj);
				$(".float").slideToggle('fast');
				$("#html_bnt").click(function(){
					$(this).attr("disabled", true);
					var html = $("#html").val();
					var c;
					if (html.length > 0){
						c = "[html]" + html + "[/html]";
					}
					$(this).attr("disabled", false);
					NodeForEvioBject.insertAtCaret(c, false);
					$(".float").hide('fast');
				});
			}
		},
		down : function(obj){
			if (NodeForEvioBject.cloneTextRange().length > 0){
				jConfirm("是否是会员才能下载?", "确认对话框", function(r){
					if (r){
						NodeForEvioBject.insertAtCaret("[mdown]{$}[/mdown]", true);
					}else{
						NodeForEvioBject.insertAtCaret("[down]{$}[/down]", true);
					}									 
				})
			}else{
				try{$(".float").remove();}catch(e){}
				var div = "<div class=\"float\">" + this.ovelayBox(textOutString.down()) + "</div>";
				$(div).insertAfter(obj);
				$(".float").slideToggle('fast');
				$("#down_bnt").click(function(){
					$(this).attr("disabled", true);
					var downTitle = $("#downTitle").val();
					var downUrl = $("#downUrl").val();
					var downType = 0;
					$("input[name=\"downType\"]").each(function(){
						if ($(this).attr("checked")){downType = $(this).val()}
					});
					var c;
					if (downTitle.length > 0){
						if (downType == 0){
							c = "[down=" + downUrl + "]" + downTitle + "[/down]";
						}else{
							c = "[mdown=" + downUrl + "]" + downTitle + "[/mdown]";
						}
					}else{
						if (downType == 0){
							c = "[down]" + downUrl + "[/down]";
						}else{
							c = "[mdown]" + downUrl + "[/mdown]";
						}
					}
					$(this).attr("disabled", false);
					NodeForEvioBject.insertAtCaret(c, false);
					$(".float").hide('fast');
				});
			}
		},
		color : function(obj){
			if (NodeForEvioBject.cloneTextRange().length > 0){
				try{$(".float").remove();}catch(e){}
				var div = "<div class=\"float\">" + this.ovelayBox(textOutString.color()) + "</div>";
				$(div).insertAfter(obj);
				$(".float").slideToggle('fast');
				$(obj).next().find(".ubbOverlay_wrap > .ubbOverlay_inner > .outWrap").css({
					"padding-left" : "1px",
					"padding-right" : "0px",
					"padding-top" : "1px",
					"padding-bottom" : "1px"
				});
				$(".colorBlock").each(function(){
					$(this).bind("click", 
								 
						function(){
							var c = $(this).attr("rel");
							NodeForEvioBject.insertAtCaret("[color=#" + c + "]{$}[/color]", true);
							$(".float").hide('fast');
						}			 
								 
					)						   
				});
			}else{
				jAlert("需要选中文字", "确认对话框");
			}
		}
	}
	
	//UBB到HTML的转换方法
	$.fn.UBBTOHTML = function(obj){
		var Text = $(this).val();
		for (var i = 0 ; i < UBBASHTML.length ; i++){
			Text = Text.replace(UBBASHTML[i].reg, UBBASHTML[i].value);
		}
		$(obj).html(Text);
	}
	
	//HTML到UBB的转换方法
	$.fn.HTMLTOUBB = function(obj){
		var Text = $(this).html();
		for (var i = 0 ; i < HTMLASUBB.length ; i++){
			Text = Text.replace(HTMLASUBB[i].reg, HTMLASUBB[i].value);
		}
		$(obj).val(Text);
	}
	
	
})(jQuery);

var textOutString = {
	getColor : function(){
		// 色值范围
			var color = "FFFFFF,FFCCFF,FF99FF,FF66FF,FF33FF,FF00FF,CCFFFF,CCCCFF,CC99FF,CC66FF,CC33FF,CC00FF,99FFFF,99CCFF,9999FF,9966FF,9933FF,9900FF,66FFFF,66CCFF,6699FF,6666FF,6633FF,6600FF,33FFFF,33CCFF,3399FF,3366FF,3333FF,3300FF,00FFFF,00CCFF,0099FF,0066FF,0033FF,0000FF,FFFFCC,FFCCCC,FF99CC,FF66CC,FF33CC,FF00CC,CCFFCC,CCCCCC,CC99CC,CC66CC,CC33CC,CC00CC,99FFCC,99CCCC,9999CC,9966CC,9933CC,9900CC,66FFCC,66CCCC,6699CC,6666CC,6633CC,6600CC,33FFCC,33CCCC,3399CC,3366CC,3333CC,3300CC,00FFCC,00CCCC,0099CC,0066CC,0033CC,0000CC,FFFF99,FFCC99,FF9999,FF6699,FF3399,FF0099,CCFF99,CCCC99,CC9999,CC6699,CC3399,CC0099,99FF99,99CC99,999999,996699,993399,990099,66FF99,66CC99,669999,666699,663399,660099,33FF99,33CC99,339999,336699,333399,330099,00FF99,00CC99,009999,006699,003399,000099,FFFF66,FFCC66,FF9966,FF6666,FF3366,FF0066,CCFF66,CCCC66,CC9966,CC6666,CC3366,CC0066,99FF66,99CC66,999966,996666,993366,990066,66FF66,66CC66,669966,666666,663366,660066,33FF66,33CC66,339966,336666,333366,330066,00FF66,00CC66,009966,006666,003366,000066,FFFF33,FFCC33,FF9933,FF6633,FF3333,FF0033,CCFF33,CCCC33,CC9933,CC6633,CC3333,CC0033,99FF33,99CC33,999933,996633,993333,990033,66FF33,66CC33,669933,666633,663333,660033,33FF33,33CC33,339933,336633,333333,330033,00FF33,00CC33,009933,006633,003333,000033,FFFF00,FFCC00,FF9900,FF6600,FF3300,FF0000,CCFF00,CCCC00,CC9900,CC6600,CC3300,CC0000,99FF00,99CC00,999900,996600,993300,990000,66FF00,66CC00,669900,666600,663300,660000,33FF00,33CC00,339900,336600,333300,330000,00FF00,00CC00,009900,006600,003300,000000";
			var colorArray = color.split(",");
			var colorText = "";
			for (var colori = 0 ; colori < colorArray.length ; colori++){
				colorText += "<div class=\"colorBlock\" style=\"background:#" + colorArray[colori] + "\" rel=\"" + colorArray[colori] + "\" title=\"#" + colorArray[colori] + "\"></div>";
			}
			colorText += "<div class=\"clear\"></div>";
			return colorText;
	},
	url : function(){
		var c = "<div style=\"width:190px;\">"
		+			"<div>链 接 名 <input type=\"text\" class=\"text\" id=\"urlname\"></div>"
		+			"<div>链接地址 <input type=\"text\" class=\"text\" id=\"url\"></div>"
		+			"<div><input type=\"button\" value=\"插入\" class=\"button\" id=\"url_bnt\"></div>"
		+		"</div>"
		return c;
	},
	mail : function(){
		var c = "<div style=\"width:190px;\">"
		+			"<div>收 件 人 <input type=\"text\" class=\"text\" id=\"mailname\"></div>"
		+			"<div>邮件地址 <input type=\"text\" class=\"text\" id=\"mail\"></div>"
		+			"<div><input type=\"button\" value=\"插入\" class=\"button\" id=\"mail_bnt\"></div>"
		+		"</div>"
		return c;
	},
	quote : function(){
		var c = "<div style=\"width:190px;\">"
		+			"<div>引用来自 <input type=\"text\" class=\"text\" id=\"quotename\"></div>"
		+			"<div>引用内容 <textarea id=\"quote\" style=\"width:180px; height:50px;\" class=\"text\"></textarea></div>"
		+			"<div style=\"margin-top:10px;\"><input type=\"button\" value=\"插入\" class=\"button\" id=\"quote_bnt\"></div>"
		+		"</div>"
		return c;
	},
	code : function(){
		var c = "<div style=\"width:190px;\">"
		+			"<div>代码类型 <input type=\"text\" class=\"text\" id=\"codename\"></div>"
		+			"<div>代码内容 <textarea id=\"code\" style=\"width:180px; height:50px;\" class=\"text\"></textarea></div>"
		+			"<div style=\"margin-top:10px;\"><input type=\"button\" value=\"插入\" class=\"button\" id=\"code_bnt\"></div>"
		+		"</div>"
		return c;
	},
	html : function(){
		var c = "<div style=\"width:190px;\">"
		+			"<div>HTML代码内容 <textarea id=\"html\" style=\"width:180px; height:50px;\" class=\"text\"></textarea></div>"
		+			"<div style=\"margin-top:10px;\"><input type=\"button\" value=\"插入\" class=\"button\" id=\"html_bnt\"></div>"
		+		"</div>"
		return c;
	},
	down : function(){
		var c = "<div style=\"width:220px;\">"
		+			"<div>下载标题 <input type=\"text\" class=\"text\" id=\"downTitle\"></div>"
		+			"<div><input type=\"radio\" name=\"downType\" value=\"0\">  普通下载 <input type=\"radio\" name=\"downType\" value=\"1\" checked>  会员下载</div>"
		+			"<div>下载链接 <input type=\"text\" class=\"text\" value=\"\" style=\"width:150px;\" id=\"downUrl\"></div>"
		+			"<div style=\"margin-top:10px;\"><input type=\"button\" value=\"插入\" class=\"button\" id=\"down_bnt\"></div>"
		+		"</div>"
		return c;
	},
	color : function(){
		var c = "<div style=\"width:236px;\">"
		+			"<div style=\"width:234px; border-top:1px solid #000; border-left:1px solid #000\">" + this.getColor() + "</div>"
		+		"</div>"
		return c;
	}
}

var UBBASHTML = [
	{
		reg : /\[color=(#\w{3,10}|\w{3,10})\]([^\r]*?)\[\/color\]/igm, 
		value : "<span style=\"color:$1\">$2</span>"
	}
];

var HTMLASUBB = [
	{
		reg : /\<span\sstyle\=\"color\:(.*?)\"\>([\s\S]*?)\<\/span\>/igm,
		value : "[color=$1]$2[/color]"
	}			 
];

//1.var n = 0;   
//2.var hex = new Array('FF', 'CC', '99', '66', '33', '00');   
//3.  
//4.function colorPanel(){   
//5.    for (var i = 0; i < 6; i++) {   
//6.        for (var j = 0; j < 6; j++) {   
//7.            for (var k = 0; k < 6; k++) {   
//8.                n++;   
//9.                var color = hex[j] + hex[k] + hex[i];   
//10.                document.write(color + '<br />');   
//11.            }   
//12.        }   
//13.    }   
//14.}  