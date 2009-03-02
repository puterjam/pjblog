//PJBlog 3 Ajax Action File
//Author:evio
var GetFile = ["Action.asp?action="]
var GetAction = ["Hidden&", "checkname&"]

// 关于 [Hidden] 标签的修复代码
function Hidden(i){
	var TempStr;
	var ajax = new AJAXRequest;
	ajax.get(
			 GetFile[0]+GetAction[0]+"TimeStamp="+new Date().getTime(),
			 function(obj) {
				 TempStr = obj.responseText;
				 if ( TempStr == "1" ){
					 $("hidden1_"+i).style.display = "";
					 $("hidden2_"+i).style.display = "none";
				 }else{
					 $("hidden1_"+i).style.display = "none";
					 $("hidden2_"+i).style.display = "";
				 }
			 }
	 );
}

//关于注册页面的用户名判断
function CheckName(){
	var TempStr, StrHtml, Bool;
	$("CheckName").innerHTML = "<font color='#ccc'>检测用户名...</font>"
	var UserName = document.forms["frm"].username.value;
	var HoldValue = $("PostBack_UserName").value;
	var ajax = new AJAXRequest;
	ajax.get(
			 GetFile[0]+GetAction[1]+"usename="+escape(UserName)+"&TimeStamp="+new Date().getTime(),
			 function(obj) {
				 TempStr = obj.responseText;
				 StrHtml = TempStr.split("|$|")[0];
				 Bool = TempStr.split("|$|")[1];
				 if ( Bool == "True" ){
					 $("PostBack_UserName").value = "True|$|" + HoldValue.split("|$|")[1];
					 $("CheckName").innerHTML = StrHtml;
				 }else{
					 $("PostBack_UserName").value = "False|$|" + HoldValue.split("|$|")[1];
					 $("CheckName").innerHTML = StrHtml;
				 }
			 }
	 );
}
function IsPost(){
	if ($("PostBack_UserName").value == "True|$|True"){
		document.forms["frm"].submit();
	}else{
		alert("你填写的资料不正确，或者用户名已被注册！")
	}
}

//关于密码的判断
function CheckPwd(){
	$("CheckPwds").innerHTML = "正在检测密码";
	var Pwd = $("cpassword").value;
	var SePwd = $("cConfirmpassword").value;
	var HoldValue = $("PostBack_UserName").value;
	if (Pwd != SePwd){
		$("PostBack_UserName").value = HoldValue.split("|$|")[0] + "|$|False";
		$("CheckPwds").innerHTML = "&nbsp;&nbsp;<u><font color=red>两次输入的密码不同</font></u>";
	}else{
		$("PostBack_UserName").value = HoldValue.split("|$|")[0] + "|$|True";
		$("CheckPwds").innerHTML = "&nbsp;&nbsp;<u><font color='blue'>两次输入的密码相同！</font></u>"
	}
}