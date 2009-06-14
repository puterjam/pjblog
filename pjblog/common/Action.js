//PJBlog 3 Ajax Action File
//Author:evio
var GetFile = ["Action.asp?action="];
var GetAction = ["Hidden&", "checkname&", "PostSave&", "UpdateSave&", "AddNewCate&", "GetPassReturnInfo&", "UpdatePass&", "CheckPostName&", "updatepassto&", "IndexAudit&", "ReadArticleComentByCommentID&", "GetHttpURL&"];

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
		$("CheckPwds").innerHTML = "&nbsp;&nbsp;<font color=red>两次输入的密码不同</font>";
	}else{
		$("PostBack_UserName").value = HoldValue.split("|$|")[0] + "|$|True";
		$("CheckPwds").innerHTML = "&nbsp;&nbsp;<font color='blue'>两次输入的密码相同！</font>"
	}
}

// Ajax草稿保存
var e_event = null;
function UBB_AjaxLogSave(){
	var obj = $("AjaxTimeSave");
	if (obj.style.display == "none"){ obj.style.display = ""; }
	if (e_event == null){
		OutTime();
	}else{
		clearTimeout(e_event);
		e_event = null;
		obj.innerHTML = (obj.innerHTML + "&nbsp;&nbsp;&nbsp;&nbsp;<span style='font-size:9px;text-decoration:none'><a href='javascript:void(0);' onclick=$('AjaxTimeSave').style.display='none'><b>Close</b></a></span>");
	}
}
function OutTime(){    
    var loop = time;    
    $("AjaxTimeSave").innerHTML = loop + " 秒后自动保存为草稿!";    
    e_event = setTimeout('goTime('+loop+');',2000);    
}    
function goTime(i){    
    i = i - 1;    
    if ( i != 0 ){    
        $("AjaxTimeSave").innerHTML = i + " 秒后自动保存为草稿!";    
        e_event = setTimeout("goTime("+i+");",1000);    
    }else{    
        PostSave();   
		try{
        	e_event = setTimeout('goTime('+(time+1)+')',3000); 
		}catch(e){
			if (e.description.length > 0){
				e_event = setTimeout('goTime('+(time+1)+')',3000);
			}
		}
    }    
}    
   
//发表时的保存    
function PostSave(){    
    var TempStr, left, right, ToWhere, postId;    
    var ajax = new AJAXRequest;    
    //var str = ReadCode();  	
    var FirstPost = document.forms["frm"].FirstPost.value;    
    if (FirstPost == 1){    
        //var zpt =document.forms["frm"].postbackId.value;    
        ToWhere = GetAction[3];    
    }else{    
        ToWhere = GetAction[2];    
    }   
	//alert(ReadCode());
	//alert(document.forms["frm"].postbackId.value)
    ajax.post(    
             GetFile[0] + ToWhere + "TimeStamp=" + new Date().getTime(),
			 ReadCode(),
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
	
	//评论倒序
	logcomorder = checkbox_model("label");
	
	//禁止评论
	logDisComment = checkbox_model("label2");
	
	//日志置顶
	logIsTop = checkbox_model("label3");
	
	//隐私和META
	//logIsHidden = checkbox_model("Secret");
	logMeta = checkbox_model("Meta");
	
	//来源网址
	logFrom = $("log_From").value;
	logFromURL = $("log_FromURL").value;
	
	//禁止显示图片
	logdisImg = checkbox_model("label4");
	
	//禁止表情转换
	logDisSM = checkbox_model("label5");
	
	//禁止自动转换链接
	logDisURL = checkbox_model("label6");
	
	//禁止自动转换关键字
	logDisKey = checkbox_model("label7");
	
	//引用通告
	logQuote = $("logQuote").value;
	
    //str = "title=" + escape(title) + "&cname=" + escape(cname) + "&ctype=" + escape(ctype) + "&logweather=" + escape(logweather) + "&logLevel=" + escape(logLevel) + "&logcomorder=" + escape(logcomorder) + "&logDisComment=" + escape(logDisComment) + "&logIsTop=" + escape(logIsTop) + "&logMeta=" + escape(logMeta) + "&logFrom=" + escape(logFrom) + "&logFromURL=" + escape(logFromURL) + "&logdisImg=" + escape(logdisImg) + "&logDisSM=" + escape(logDisSM) + "&logDisURL=" + escape(logDisURL) + "&logDisKey=" + escape(logDisKey) + "&logQuote=" + escape(logQuote) + "&tags=" + escape(Tags) + "&cateid=" + escape(mCateID) + "&Message=" + escape(Message) + "&";
	
	str = escape(title) + "|$|" + escape(cname) + "|$|" + escape(ctype) + "|$|" + escape(logweather) + "|$|" + escape(logLevel) + "|$|" + escape(logcomorder) + "|$|" + escape(logDisComment) + "|$|" + escape(logIsTop) + "|$|" + escape(logMeta) + "|$|" + escape(logFrom) + "|$|" + escape(logFromURL) + "|$|" + escape(logdisImg) + "|$|" + escape(logDisSM) + "|$|" + escape(logDisURL) + "|$|" + escape(logDisKey) + "|$|" + escape(logQuote) + "|$|" + escape(Tags) + "|$|" + escape(mCateID) + "|$|" + escape(Message);
	 if (document.forms["frm"].FirstPost.value == 1){
		 var zpt =document.forms["frm"].postbackId.value;    
         str = str + "|$|" + escape(zpt);
	 }  
	
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

// checkbox 选择器
function checkbox_model(A){
	var temp;
	if ($(A).checked){
		temp = $(A).value;
	}else{
		temp = "0";
	}
	return temp;
}



function GoToCateAdd(){
	var CateAddText = $("log_NewCate").value;
	var ajax = new AJAXRequest;
	ajax.get(
			 GetFile[0]+GetAction[4]+"newcate="+escape(CateAddText)+"&TimeStamp="+new Date().getTime(),
			 function(obj) {
				 TempStr = parseInt(obj.responseText);
				 try{
				 	document.forms[0].log_CateID.add(new Option(CateAddText, TempStr, false, true));
					hidePopup();
				 }catch(e){
					if (e.description.length > 0){alert(e.description);hidePopup();}
				 }
			 }
	 );
	
	//alert(document.forms[0].log_CateID.options[document.forms[0].log_CateID.options.selectedIndex].value);
}


function GoToPassCheck(Name, i){
	var GetPass = $("c_Answer").value;
	var ajax = new AJAXRequest;
	ajax.get(
			 GetFile[0]+GetAction[5]+"password="+escape(GetPass) + "&name=" + escape(Name) +"&TimeStamp="+new Date().getTime(),
			 function(obj) {
				 TempStr = obj.responseText;
				 if (TempStr == "none"){
					 $("passContent").innerHTML = "<div style='text-align:center'><b><font color=red>没有该用户信息</font></b></div>";
				 }else if(TempStr == "wrong"){
					 $("passContent").innerHTML = "<div style='text-align:center'><b><font color=red>密码保护答案错误</font></b></div>";
				 }else{
					 if (i == 0){
					 	$("passContent").innerHTML = ModiyStr2(TempStr);
					 }else{
						$("passContent").innerHTML = ModiyStr3(TempStr);
					 }
				 } 
			 }
	 );
}

function GoToPassCheck2(id){
	var e_Question = $("c_Question").value;
	var e_Answer = $("c_Answer").value;
	var ajax = new AJAXRequest;
	ajax.get(
			 GetFile[0]+GetAction[6]+"id="+escape(id) + "&q=" + escape(e_Question) + "&a=" + escape(e_Answer) +"&TimeStamp="+new Date().getTime(),
			 function(obj) {
				 TempStr = obj.responseText;
				 if (TempStr == "1"){
					$("passContent").innerHTML = "<div style='text-align:center'><b><font color=blue>操作成功</font></b></div>";
				 }else{
					$("passContent").innerHTML = "<div style='text-align:center'><b><font color=red>操作失败</font></b></div>";
				 }
			 }
	 );
}

function PostPName(){
	var name = $("c_Name").value;
	var ajax = new AJAXRequest;
	ajax.get(
			 GetFile[0]+GetAction[7]+"name="+escape(name) + "&TimeStamp="+new Date().getTime(),
			 function(obj) {
				 TempStr = obj.responseText;
				 if (TempStr == "0"){
					$("passContent").innerHTML = "<div style='text-align:center'><b>没有该用户的信息</b></div>"; 
				 }else{
					 var d = TempStr.split("|$|");
					$("passContent").innerHTML =  ModiyStr(d[1], d[0], 1);
				 }
			 }
	 );
}

function Gotoupdatepass(id){
	var c_pass = $("c_Password").value;
	var c_repass = $("c_RePassword").value;
	var ajax = new AJAXRequest;
	ajax.get(
			 GetFile[0]+GetAction[8]+"id="+escape(id) + "&pass=" + escape(c_pass) + "&repass=" + escape(c_repass) + "&TimeStamp="+new Date().getTime(),
			 function(obj) {
				 TempStr = obj.responseText;
				 if (TempStr == "1"){
					 $("passContent").innerHTML = "<div style='text-align:center'><b><font color=blue>操作成功</font></b></div>";
				 }else{
					 $("passContent").innerHTML = "<div style='text-align:center'><b><font color=red>操作失败</font></b></div>";
				 }
			 }
	 );
}

function GetNoPassInforMation(id){
	var str = ModiyStr2(id);
	$("passContent").innerHTML = str;
}

function IndexAudit(id, i, Thisobj, BlogID){
	var c;
	var ajax = new AJAXRequest;
	if (i == 0){
		LoadInformation("正在通过审核ID为" + id + "的评论,请稍后...");
		c = GetFile[0]+GetAction[9]+"type=0&id=" + escape(id) + "&blogid=" + escape(BlogID) + "&TimeStamp="+new Date().getTime();
	}else if (i == 1){
		LoadInformation("正在取消审核ID为" + id + "的评论,请稍后...");
		c = GetFile[0]+GetAction[9]+"type=1&id=" + escape(id) + "&blogid=" + escape(BlogID) + "&TimeStamp="+new Date().getTime();
	}else{
		c = 0;
	}
	if (c != 0){
		ajax.get(
			 	c,
			 	function(obj) {
				 	TempStr = obj.responseText;
				 	TempStr = parseInt(TempStr);
					if (TempStr == 0){
						Thisobj.innerHTML = "取消审核";
						$("commcontent_" + id).innerHTML = $("commcontent_" + id).innerHTML.replace("[未审核评论,仅管理员可见]:&nbsp;", "");
						Thisobj.onclick = function(){IndexAudit(id, 1, Thisobj, BlogID);}
					}else if (TempStr == 1){
						Thisobj.innerHTML = "通过审核"; 
						$("commcontent_" + id).innerHTML = "[未审核评论,仅管理员可见]:&nbsp;" + $("commcontent_" + id).innerHTML;
						Thisobj.onclick = function(){IndexAudit(id, 0, Thisobj, BlogID);}
					}
					LoadInformation();
			 	}
	 	);
	}
}

function ReadArticleComentByCommentID(CommentID){
	var t = CommentID.split("|$|");
	var ajax = new AJAXRequest;
	ajax.get(
			 GetFile[0]+GetAction[10]+"CommentIDStr="+escape(CommentID) + "&TimeStamp="+new Date().getTime(),
			 function(obj) {
				 TempStr = obj.responseText;
				 var row = unescape(TempStr).split("|$|");
				 for (var i = 0 ; i < row.length ; i++){
					 try{$("commcontent_" + t[i]).innerHTML = row[i];}catch(e){}
				 }
			 }
	 );
}

function ReturnPage(URL, Obj){
	var TempStr;
	LoadInformation("正在读取信息...");
	var ajax = new AJAXRequest;
	ajax.get(
			 GetFile[0]+GetAction[11]+"url="+escape(URL) + "&TimeStamp="+new Date().getTime(),
			 function(obj) {
				 TempStr = obj.responseText;
				 $(Obj).innerHTML = TempStr;
				 LoadInformation();
			 }
	 );
}

//后台Ajax分段静态化
	 
function CreateHtml(){
	if ($('AjaxRebuildButton').disabled == false) $('AjaxRebuildButton').disabled = true;
	if(Lists.length<=0){
		$("msgbox").innerHTML="没有文章需要静态化";
	}else if(CurrentIndex == Lists.length){
		$("msgbox").innerHTML = "静态化完毕";
		$('AjaxRebuildButton').disabled = false;
	}else{
		if (IsStop == false) {
			$("msgbox").innerHTML = "<font color='#0000ff'>正在静态化第 " + Lists[CurrentIndex] + " 篇日志 ...</font>" ;
		}
		var ajax=new AJAXRequest();
		ajax.get("action.asp?action=ReBuildArticle&id=" + Lists[CurrentIndex] + "&TimeStamp="+new Date().getTime(),		
			function(obj){
					var msg=eval("("+obj.responseText+")");	
					if(msg.suc){
						CurrentIndex++;
						window.setTimeout("CreateHtml()",10);	
					}else{
						$("msgbox").innerHTML="静态化过程出现错误，已静态化" + CurrentIndex + "篇文章!";
						return;
					}
			}
		);
	}
}

function StopHtml(){
	if ($('AjaxRebuildButton').disabled == true)
	{
		$('AjaxRebuildButton').disabled = false;
		IsStop = true;
	}
	if(Lists.length <= 0){
		$("msgbox").innerHTML="无效操作";
	}else{
		CurrentIndex = Lists.length;
		$("msgbox").innerHTML = "静态化停止";
	}
}

function StartHTML(){
	IsStop = false;
	CreateHtml();
}
