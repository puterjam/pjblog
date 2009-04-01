//PJBlog 3 Ajax Action File
//Author:evio
var GetFile = ["Action.asp?action="];
var GetAction = ["Hidden&", "checkname&", "PostSave&", "UpdateSave&"];

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
	$("CheckName").innerHTML = "<font color='#ccc'>检测用户名...</font>";
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
		alert("你填写的资料不正确，或者用户名已被注册！");
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

// Ajax草稿保存  
function OutTime(){    
    var loop = time;    
    $("AjaxTimeSave").innerHTML = loop;    
    setTimeout('goTime('+loop+');',1000);    
}    
function goTime(i){    
    i = i - 1;    
    if ( i != 0 ){    
        $("AjaxTimeSave").innerHTML = i;    
        setTimeout("goTime("+i+");",1000);    
    }else{    
        PostSave();    
        setTimeout('goTime('+(time+1)+')',3000);    
    }    
}    
   
//发表时的保存    
function PostSave(){    
    var TempStr, left, right, ToWhere, postId;    
    var ajax = new AJAXRequest;    
    var str = ReadCode();    
    var FirstPost = document.forms["frm"].FirstPost.value;    
    if (FirstPost == 1){    
        var zpt =document.forms["frm"].postbackId.value;    
        ToWhere = GetAction[3] + "postId=" + escape(zpt) + "&";    
    }else{    
        ToWhere = GetAction[2];    
    }   
    ajax.get(    
             GetFile[0] + ToWhere + str + "TimeStamp=" + new Date().getTime(),    
             function(obj) {    
                 TempStr = obj.responseText;    
                 left = TempStr.split("|$|")[0];    
                 right = TempStr.split("|$|")[1];    
                 postId = TempStr.split("|$|")[2];    
                 $("AjaxTimeSave").innerHTML = left;    
                 if ( right == 1 ){    
                     document.forms["frm"].FirstPost.value = 1;    
                     document.forms["frm"].postbackId.value = postId;    
                 }    
             }    
     );    
}    
   
// 从表单中读取数据    
function ReadCode(){ 
	var mCateID, str, cname, ctype, logweather, logLevel;
	
	//title
    var title = document.forms["frm"].title.value;  
	
	//cname
    try {
		cname = document.forms["frm"].cname.value;
	} catch(e){
		if (e.description != "" ){
			cname = "";
		} 
	} 
	
	//content
    var Message = document.forms["frm"].Message.value; 
	
	//tag
	var Tags = document.forms["frm"].tags.value; 
	
	//cate
	try{
		mCateID = $("select2").options[$("select2").options.selectedIndex].value;	
	}catch(e){
		if (e.description != "" ){
			try{
				mCateID = $("log_CateID").value;
			}catch(e){
				alert(e.description);
			}
		}
	}
	
	//ctype
	try{
		ctype = document.forms["frm"].ctype.options[document.forms["frm"].ctype.options.selectedIndex].value;	
	}catch(e){
		if (e.description != "" ){
			ctype = "1";
		}
	}
	
	//logweather
	logweather = select_model("logweather", "sunny");
	
	//logLevel
	logLevel = select_model("logLevel", "level3");
	
    str = "title=" + escape(title) + "&cname=" + escape(cname) + "&ctype=" + escape(ctype) + "&logweather=" + escape(logweather) + "&logLevel=" + escape(logLevel) + "&tags=" + escape(Tags) + "&cateid=" + escape(mCateID) + "&Message=" + escape(Message) + "&";     
    return str;    
}

//select 选择器
function select_model(A, B){
	var c;
	try{
		c = $(A).options[$(A).options.selectedIndex].value;
	}catch(e){
		if (e.description != ""){
			c = B;
		}
	}
	return c;
}