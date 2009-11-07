// JavaScript Document
function Check(){
	init(
		 
	);
	function init(){
		
	}
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
		},
		login : function(){
			
		}
	}
}
var CheckForm = new Check();