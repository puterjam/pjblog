// JavaScript Document
function Check(){
	init();
	function init(){
		
	}
	// 用户
	this.User = {
		GetCode : function(Str, id){
			$(id).innerHTML = "<img id=\"vcodeImg\" src=\"about:blank\" onerror=\"this.onerror=null;this.src='" + Str + "?s='+Math.random();\" alt=\"" + lang.Set.Asp(4) + "\" title=\"" + lang.Set.Asp(5) + "\" style=\"margin-right:40px;cursor:pointer;width:40px;height:18px;margin-bottom:-4px;margin-top:3px;\" onclick=\"src='" + Str + "?s='+Math.random()\"/>";
		},
		CheckCode : function(FormObj, ToObj){
			var CheckValue = $(FormObj).value.trim();
			if (CheckValue.length >= 4){
				Ajax({
				  url : "../pjblog.logic/log_Ajax.asp?action=CheckCode&s=" + Math.random(),
				  method : "GET",
				  content : "",
				  oncomplete : function(obj){
						var bvalue = obj.responseText.trim();
						if (bvalue == CheckValue){
							$(ToObj).innerHTML = "<img src=\"../images/check_right.gif\" />";
						}else{
							$(ToObj).innerHTML = "<img src=\"../images/check_error.gif\" />";
						}
				  },
				  ononexception:function(obj){
					  alert(obj.state);
				  }
			   	});
			}
		}
	}
	// 分类
	this.Category = {
		Add : function(obj){
			_obj = obj;
			obj.innerHTML = "正在提交...";
			obj.disabled = true;
			var new_order = $("new_order").value;
			var new_icon = $("new_icon").options[$("new_icon").options.selectedIndex].value;
			var new_name = $("new_name").value;
			var new_Intro = $("new_Intro").value;
			var new_Part = $("new_Part").value;
			var new_URL = $("new_URL").value;
			var new_local = $("new_local").options[$("new_local").options.selectedIndex].value;
			var new_Secret = $("new_Secret").options[$("new_Secret").options.selectedIndex].value;
			
			if ((new_order.legnth < 1) || (!/^\d+$/.test(new_order))){alert("您所填写的排序格式不正确!");$("new_order").select();obj.innerHTML = "保存新分类" ; obj.disabled = false; return;}
			if (new_icon.length < 1){alert("您所选择的图标不正确!"); obj.innerHTML = "保存新分类" ; obj.disabled = false; return;}
			if (new_name.length < 2){alert("标题格式不正确或为空, 标题应大于2位字符!"); $("new_name").select(); obj.innerHTML = "保存新分类" ; obj.disabled = false; return;}
			Ajax({
				url : "../pjblog.logic/control/log_category.asp?action=add&s=" + Math.random(),
				method : "POST",
				content : "cate_Order=" + escape(new_order) + "&cate_icon=" + escape(new_icon) + "&cate_Name=" + escape(new_name) + "&cate_Intro=" + escape(new_Intro) + "&cate_Folder=" + escape(new_Part) + "&cate_URL=" + escape(new_URL) + "&cate_local=" + escape(new_local) + "&cate_Secret=" + escape(new_Secret),
				oncomplete : function(obj){
					var json = obj.responseText.json();
					if (json.Suc){
						Box.selfWidth = true;
						Box.selefHeight = true;
						var layer = Box.FollowBox($("AddRowMark_del"), $("AddRowMark_del").offsetWidth, $("AddRowMark_del").offsetHeight, 0, "保存新分类成功!");
						layer.className = "opacity";
						layer.style.cssText += "; margin:0 auto; text-align:center; line-height:30px; font-size:11px; color:#000000; font-weight:bolder";
						layer.id = "layerTip";
						setTimeout("$('layerTip').parentNode.removeChild($('layerTip'))", 2000);
						
						$("AddRowMark_del").style.backgroundColor = "#ffffff";
						_obj.disabled = false;
						// --------------------------------------------------
						var checkbox = document.createElement("div");
						checkbox.innerHTML = "<input type=\"checkbox\" value=\"" + json.Info.trim() + "\" name=\"SelectID\" /><input type=\"hidden\" value=\"" + json.Info.trim() + "\" name=\"Cate_ID\" />";
						$("new_selectid").parentNode.replaceChild(checkbox, $("new_selectid"));
						// --------------------------------------------------
						var _div = document.createElement("div");
						_div.innerHTML = "<input type=\"text\" class=\"text\" name=\"cate_count\" value=\"0\" size=\"2\" readonly=\"readonly\" style=\"background:#ffe\"/> 篇";
						_obj.parentNode.replaceChild(_div, _obj);
						
						if (new_URL.length > 0) $("new_Secret").disabled = true;
						$("AddRowMark_del").id = "";
						$("new_order").id = "";
						$("new_icon").id = "";
						$("new_name").id = "";
						$("new_Intro").id = "";
						$("new_Part").id = "";
						$("new_URL").id = "";
						$("new_local").id = "";
						$("new_Secret").id = "";
						try{$("Addbutton").disabled = false}catch(e){}
					}else{alert(json.Info);_obj.disabled = false; _obj.innerHTML = "保存新分类" ;}
				},
				ononexception:function(obj){
					alert(obj.state);
				}
			});
		}
	}
}
var CheckForm = new Check();