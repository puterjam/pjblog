// JavaScript Document
var cookie={
    SET	: function(name, value, days) {var expires = "";if (days) {var d = new Date();d.setTime(d.getTime() + days * 24 * 60 * 60 * 1000);expires = "; expires=" + d.toGMTString();}document.cookie = name + "=" + value + expires + "; path=/";},
	GET	: function (name) {var re = new RegExp("(\;|^)[^;]*(" + name + ")\=([^;]*)(;|$)");var res = re.exec(document.cookie);return res != null ? res[3] : null;}
};

/************************************************************/
// 这里的必须按照你站上的设置修改,由于是JS,无法从文件调用这些参数,所以请手动设置.
//var cookieName = "evio", AccessPath = "blogDB/PBLog3.asp";
var cookieName = cookie.GET("InstallCookie"); // cookie 名称
var AccessPath = cookie.GET("InstallAccess"); // 数据库地址,按需要修改.
//alert(cookieName + "|" + AccessPath)
/************************************************************/

String.prototype.json = function(){
	try{eval("var jsonStr = (" + this.toString() + ")");}catch(ex){var jsonStr = null;}
	return jsonStr;
};

/***************************************
	定义$方法
****************************************/
function $(){ 
	var elements = new Array(); 
	for (var i = 0; i < arguments.length; i++){ 
		var element = arguments[i]; 
		if (typeof element == 'string') 
			element = document.getElementById(element); 
		if (element) {} else {element = null;}
		if (arguments.length == 1) {return element;} else {elements.push(element);}
	} 
	return elements; 
}

/////////////////////////////////////////////////////////////////////////////
// 	数据库更新字段
/////////////////////////////////////////////////////////////////////////////

var NewArray_0 = [
	"create table blog_book",
	"create table blog_Category",
	"create table blog_Comment",
	"create table blog_Content",
	"create table blog_Counter",
	"create table blog_Files",
	"create table blog_Info",
	"create table blog_Keywords",
	"create table blog_Links",
	"create table blog_Member",
	"create table blog_ModSetting",
	"create table blog_module",
	"create table blog_Notdownload",
	"create table blog_Smilies",
	"create table blog_status",
	"create table blog_tag",
	"create table blog_Trackback"
];

var NewArray_1 = [
	"ALTER TABLE [blog_book] ADD COLUMN book_ID COUNTER(1, 1) PRIMARY KEY",
	"ALTER TABLE [blog_book] ADD COLUMN book_Messager varchar(20)",
	"ALTER TABLE [blog_book] ADD COLUMN book_QQ varchar(20)",
	"ALTER TABLE [blog_book] ADD COLUMN book_Url varchar(255)",
	"ALTER TABLE [blog_book] ADD COLUMN book_Mail varchar(255)",
	"ALTER TABLE [blog_book] ADD COLUMN book_face varchar(50)",
	"ALTER TABLE [blog_book] ADD COLUMN book_IP varchar(20)",
	"ALTER TABLE [blog_book] ADD COLUMN book_Content ntext",
	"ALTER TABLE [blog_book] ADD COLUMN book_PostTime smalldatetime default now()",
	"ALTER TABLE [blog_book] ADD COLUMN book_reply ntext",
	"ALTER TABLE [blog_book] ADD COLUMN book_HiddenReply bit default 1",
	"ALTER TABLE [blog_book] ADD COLUMN book_replyTime smalldatetime",
	"ALTER TABLE [blog_book] ADD COLUMN book_replyAuthor varchar(20)"
];

var NewArray_2 = [
	"ALTER TABLE [blog_Category] ADD COLUMN cate_ID COUNTER(1, 1) PRIMARY KEY",
	"ALTER TABLE [blog_Category] ADD COLUMN cate_Name varchar(50)",
	"ALTER TABLE [blog_Category] ADD COLUMN cate_Order Integer default 0",
	"ALTER TABLE [blog_Category] ADD COLUMN cate_Intro varchar(100)",
	"ALTER TABLE [blog_Category] ADD COLUMN cate_OutLink bit default 0",
	"ALTER TABLE [blog_Category] ADD COLUMN cate_URL varchar(255)",
	"ALTER TABLE [blog_Category] ADD COLUMN cate_Lock bit default 0",
	"ALTER TABLE [blog_Category] ADD COLUMN cate_icon varchar(255)",
	"ALTER TABLE [blog_Category] ADD COLUMN cate_count Integer default 0",
	"ALTER TABLE [blog_Category] ADD COLUMN cate_local Integer default 0",
	"ALTER TABLE [blog_Category] ADD COLUMN cate_Secret bit default 0",
	"ALTER TABLE [blog_Category] ADD COLUMN cate_Part ntext"
];

var NewArray_3 = [
	"ALTER TABLE [blog_Comment] ADD COLUMN comm_ID COUNTER(1, 1) PRIMARY KEY",
	"ALTER TABLE [blog_Comment] ADD COLUMN blog_ID Integer default 0",
	"ALTER TABLE [blog_Comment] ADD COLUMN comm_Content ntext",
	"ALTER TABLE [blog_Comment] ADD COLUMN comm_Author varchar(24)",
	"ALTER TABLE [blog_Comment] ADD COLUMN comm_PostTime smalldatetime default Now()",
	"ALTER TABLE [blog_Comment] ADD COLUMN comm_DisSM Integer default 0",
	"ALTER TABLE [blog_Comment] ADD COLUMN comm_DisUBB Integer default 0",
	"ALTER TABLE [blog_Comment] ADD COLUMN comm_DisIMG Integer default 0",
	"ALTER TABLE [blog_Comment] ADD COLUMN comm_AutoURL Integer default 1",
	"ALTER TABLE [blog_Comment] ADD COLUMN comm_PostIP varchar(20)",
	"ALTER TABLE [blog_Comment] ADD COLUMN comm_AutoKEY Integer default 0",
	"ALTER TABLE [blog_Comment] ADD COLUMN comm_IsAudit bit default False"
];

var NewArray_4 = [
	"ALTER TABLE [blog_Content] ADD COLUMN log_ID COUNTER(1, 1) PRIMARY KEY",
	"ALTER TABLE [blog_Content] ADD COLUMN log_CateID Integer default 0",
	"ALTER TABLE [blog_Content] ADD COLUMN log_Title varchar(100)",
	"ALTER TABLE [blog_Content] ADD COLUMN log_Intro ntext",
	"ALTER TABLE [blog_Content] ADD COLUMN log_Author varchar(24)",
	"ALTER TABLE [blog_Content] ADD COLUMN log_Modify varchar(50)",
	"ALTER TABLE [blog_Content] ADD COLUMN log_From varchar(20)",
	"ALTER TABLE [blog_Content] ADD COLUMN log_FromUrl varchar(255)",
	"ALTER TABLE [blog_Content] ADD COLUMN log_Quote varchar(255)",
	"ALTER TABLE [blog_Content] ADD COLUMN log_Content ntext",
	"ALTER TABLE [blog_Content] ADD COLUMN log_PostTime smalldatetime default Now()",
	"ALTER TABLE [blog_Content] ADD COLUMN log_CommNums Integer default 0",
	"ALTER TABLE [blog_Content] ADD COLUMN log_ViewNums Integer default 0",
	"ALTER TABLE [blog_Content] ADD COLUMN log_QuoteNums Integer default 0",
	"ALTER TABLE [blog_Content] ADD COLUMN log_IsShow bit default True",
	"ALTER TABLE [blog_Content] ADD COLUMN log_DisComment bit default False",
	"ALTER TABLE [blog_Content] ADD COLUMN log_IsTop bit default False",
	"ALTER TABLE [blog_Content] ADD COLUMN log_weather varchar(10) default sunny",
	"ALTER TABLE [blog_Content] ADD COLUMN log_level varchar(10) default level3",
	"ALTER TABLE [blog_Content] ADD COLUMN log_edittype Integer default 1",
	"ALTER TABLE [blog_Content] ADD COLUMN log_comorder bit default 1",
	"ALTER TABLE [blog_Content] ADD COLUMN log_mf varchar(50)",
	"ALTER TABLE [blog_Content] ADD COLUMN log_ubbFlags varchar(10) default 110",
	"ALTER TABLE [blog_Content] ADD COLUMN log_IsDraft bit default 0",
	"ALTER TABLE [blog_Content] ADD COLUMN log_tag varchar(255)",
	"ALTER TABLE [blog_Content] ADD COLUMN log_Readpw varchar(50)",
	"ALTER TABLE [blog_Content] ADD COLUMN log_Pwtips varchar(150)",
	"ALTER TABLE [blog_Content] ADD COLUMN log_cname ntext",
	"ALTER TABLE [blog_Content] ADD COLUMN log_ctype ntext",
	"ALTER TABLE [blog_Content] ADD COLUMN log_KeyWords ntext",
	"ALTER TABLE [blog_Content] ADD COLUMN log_Description ntext",
	"ALTER TABLE [blog_Content] ADD COLUMN log_Meta Integer",
	"ALTER TABLE [blog_Content] ADD COLUMN log_Pwtitle bit default False",
	"ALTER TABLE [blog_Content] ADD COLUMN log_Pwcomm bit default False"
];

var NewArray_5 = [
	"ALTER TABLE [blog_Counter] ADD COLUMN coun_ID COUNTER(1, 1) PRIMARY KEY",
	"ALTER TABLE [blog_Counter] ADD COLUMN coun_IP varchar(20)",
	"ALTER TABLE [blog_Counter] ADD COLUMN coun_OS varchar(50)",
	"ALTER TABLE [blog_Counter] ADD COLUMN coun_Browser varchar(50)",
	"ALTER TABLE [blog_Counter] ADD COLUMN coun_Time smalldatetime default Now()",
	"ALTER TABLE [blog_Counter] ADD COLUMN coun_Referer varchar(255)"
];

var NewArray_6 = [
	"ALTER TABLE [blog_Files] ADD COLUMN ID COUNTER(1, 1) PRIMARY KEY",
	"ALTER TABLE [blog_Files] ADD COLUMN FilesPath varchar(255)",
	"ALTER TABLE [blog_Files] ADD COLUMN FilesCounts Integer default 0"
];

var NewArray_7 = [
	"ALTER TABLE [blog_Info] ADD COLUMN blog_ID COUNTER(1, 1) PRIMARY KEY",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_Name varchar(50) default Blog名字",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_Title varchar(100)",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_URL varchar(150) default Blog的地址",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_LogNums Integer default 0",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_CommNums Integer default 0",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_colNums Integer default 0",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_MessageNums Integer default 0",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_MemNums Integer default 0",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_VisitNums Integer default 0",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_tbNums Integer default 0",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_PerPage Integer default 8",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_commPage Integer default 10",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_BookPage Integer default 10",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_colPage Integer default 20",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_affiche ntext",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_about ntext",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_showtotal bit default 1",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_FilterName ntext",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_FilterIP ntext",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_commTimerout Integer default 0",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_commUBB Integer default 0",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_commImg Integer default 0",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_DefaultSkin varchar(50)",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_SkinName varchar(50)",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_SplitType bit",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_introChar Integer default 4000",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_introLine Integer default 4",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_validate bit",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_ImgLink bit",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_commLength Integer default 1000",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_postFile Byte",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_postCalendar bit default -1",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_downLocal bit default -1",		
	"ALTER TABLE [blog_Info] ADD COLUMN blog_DisMod bit default -1",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_Disregister bit default 0",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_master varchar(10) default admin",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_email varchar(100) default your@mail.com",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_CountNum Integer default 500",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_wap bit default 0",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_wapNum Integer default 8",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_wapImg bit default 0",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_wapURL bit default 0",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_wapHTML bit default 0",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_wapLogin bit default 0",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_wapComment bit default 0",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_KeyWords ntext default http://www.pjhome.net/",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_Description ntext default http://www.pjhome.net/",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_SaveTime Integer default 10",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_UpLoadSet ntext",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_PasswordProtection bit default False",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_AuditOpen bit default False"	
];

var NewArray_8 = [
	"ALTER TABLE [blog_Keywords] ADD COLUMN key_ID COUNTER(1, 1) PRIMARY KEY",
	"ALTER TABLE [blog_Keywords] ADD COLUMN key_Text varchar(100)",	
	"ALTER TABLE [blog_Keywords] ADD COLUMN key_URL varchar(255)",
	"ALTER TABLE [blog_Keywords] ADD COLUMN key_Image varchar(50)"
];

var NewArray_9 = [
	"ALTER TABLE [blog_Links] ADD COLUMN link_ID COUNTER(1, 1) PRIMARY KEY",
	"ALTER TABLE [blog_Links] ADD COLUMN link_Name varchar(50)",
	"ALTER TABLE [blog_Links] ADD COLUMN link_URL varchar(200)",
	"ALTER TABLE [blog_Links] ADD COLUMN link_Image varchar(250)",
	"ALTER TABLE [blog_Links] ADD COLUMN link_Order Short default 0",
	"ALTER TABLE [blog_Links] ADD COLUMN link_IsShow bit default False",
	"ALTER TABLE [blog_Links] ADD COLUMN link_IsMain bit default False"
];

var NewArray_10 = [
	"ALTER TABLE [blog_Member] ADD COLUMN mem_ID COUNTER(1, 1) PRIMARY KEY",
	"ALTER TABLE [blog_Member] ADD COLUMN mem_Name varchar(24)",		
	"ALTER TABLE [blog_Member] ADD COLUMN mem_Password varchar(40)",
	"ALTER TABLE [blog_Member] ADD COLUMN mem_salt varchar(6) default \s",
	"ALTER TABLE [blog_Member] ADD COLUMN mem_Sex byte default 0",	
	"ALTER TABLE [blog_Member] ADD COLUMN mem_Email varchar(50)",
	"ALTER TABLE [blog_Member] ADD COLUMN mem_HideEmail bit default False",
	"ALTER TABLE [blog_Member] ADD COLUMN mem_QQ varchar(12)",
	"ALTER TABLE [blog_Member] ADD COLUMN mem_HomePage varchar(50)",
	"ALTER TABLE [blog_Member] ADD COLUMN mem_Local varchar(50)",
	"ALTER TABLE [blog_Member] ADD COLUMN mem_RegTime smalldatetime default Now()",
	"ALTER TABLE [blog_Member] ADD COLUMN mem_RegIP varchar(20)",
	"ALTER TABLE [blog_Member] ADD COLUMN mem_Status varchar(10) default Member",
	"ALTER TABLE [blog_Member] ADD COLUMN mem_PostLogs Integer default 0",
	"ALTER TABLE [blog_Member] ADD COLUMN mem_PostComms Integer default 0",
	"ALTER TABLE [blog_Member] ADD COLUMN mem_PostMessageNums Integer default 0",
	"ALTER TABLE [blog_Member] ADD COLUMN mem_Intro ntext",
	"ALTER TABLE [blog_Member] ADD COLUMN mem_lastVisit smalldatetime default Now()",
	"ALTER TABLE [blog_Member] ADD COLUMN mem_LastIP varchar(25)",
	"ALTER TABLE [blog_Member] ADD COLUMN mem_hashKey varchar(40)",
	"ALTER TABLE [blog_Member] ADD COLUMN mem_Question ntext",
	"ALTER TABLE [blog_Member] ADD COLUMN mem_Answer ntext"
];

var NewArray_11 = [
	"ALTER TABLE [blog_ModSetting] ADD COLUMN set_id COUNTER(1, 1) PRIMARY KEY",
	"ALTER TABLE [blog_ModSetting] ADD COLUMN set_ModName varchar(50)",
	"ALTER TABLE [blog_ModSetting] ADD COLUMN set_KeyName varchar(50)",
	"ALTER TABLE [blog_ModSetting] ADD COLUMN set_KeyValue ntext"
];

var NewArray_12 = [
	"ALTER TABLE [blog_module] ADD COLUMN id COUNTER(1, 1) PRIMARY KEY",
	"ALTER TABLE [blog_module] ADD COLUMN name varchar(50)",		
	"ALTER TABLE [blog_module] ADD COLUMN title varchar(50)",
	"ALTER TABLE [blog_module] ADD COLUMN type varchar(50) default sidebar",
	"ALTER TABLE [blog_module] ADD COLUMN IsHidden bit default 0",
	"ALTER TABLE [blog_module] ADD COLUMN IndexOnly bit default False",
	"ALTER TABLE [blog_module] ADD COLUMN SortID Integer default 0",
	"ALTER TABLE [blog_module] ADD COLUMN IsSystem bit default 0",
	"ALTER TABLE [blog_module] ADD COLUMN HtmlCode ntext",
	"ALTER TABLE [blog_module] ADD COLUMN PluginPath varchar(255)",
	"ALTER TABLE [blog_module] ADD COLUMN IsInstall bit default 0",
	"ALTER TABLE [blog_module] ADD COLUMN InstallFolder varchar(255)",
	"ALTER TABLE [blog_module] ADD COLUMN InstallDate smalldatetime default Now()",
	"ALTER TABLE [blog_module] ADD COLUMN SettingXML varchar(255)",
	"ALTER TABLE [blog_module] ADD COLUMN CateID Integer",
	"ALTER TABLE [blog_module] ADD COLUMN ConfigPath varchar(255)"
];

var NewArray_13 = [
	"ALTER TABLE [blog_Notdownload] ADD COLUMN blog_Nodownload OLEObject"			   
];

var NewArray_14 = [
	"ALTER TABLE [blog_Smilies] ADD COLUMN sm_ID COUNTER(1, 1) PRIMARY KEY",
	"ALTER TABLE [blog_Smilies] ADD COLUMN sm_Image varchar(25)",
	"ALTER TABLE [blog_Smilies] ADD COLUMN sm_Text varchar(25)"
];

var NewArray_15 = [
	"ALTER TABLE [blog_status] ADD COLUMN stat_id COUNTER(1, 1) PRIMARY KEY",
	"ALTER TABLE [blog_status] ADD COLUMN stat_name varchar(50)",
	"ALTER TABLE [blog_status] ADD COLUMN stat_title varchar(255)",
	"ALTER TABLE [blog_status] ADD COLUMN stat_Code varchar(50)",
	"ALTER TABLE [blog_status] ADD COLUMN stat_attSize Integer default 0",
	"ALTER TABLE [blog_status] ADD COLUMN stat_attType varchar(255)"
];

var NewArray_16 = [
	"ALTER TABLE [blog_tag] ADD COLUMN tag_id COUNTER(1, 1) PRIMARY KEY",
	"ALTER TABLE [blog_tag] ADD COLUMN tag_name varchar(20)",
	"ALTER TABLE [blog_tag] ADD COLUMN tag_count Integer default 0"
];

var NewArray_17 = [
	"ALTER TABLE [blog_Trackback] ADD COLUMN tb_ID COUNTER(1, 1) PRIMARY KEY",
	"ALTER TABLE [blog_Trackback] ADD COLUMN blog_ID Integer default 0",
	"ALTER TABLE [blog_Trackback] ADD COLUMN tb_URL varchar(255)",
	"ALTER TABLE [blog_Trackback] ADD COLUMN tb_Title varchar(100)",
	"ALTER TABLE [blog_Trackback] ADD COLUMN tb_Site varchar(50)",
	"ALTER TABLE [blog_Trackback] ADD COLUMN tb_SiteURL varchar(100)",
	"ALTER TABLE [blog_Trackback] ADD COLUMN tb_Intro ntext",
	"ALTER TABLE [blog_Trackback] ADD COLUMN tb_PostTime smalldatetime default Now()"
];

var Arrays = [
	// step one for pjblog, detection System Environment!
	[
	 	[1, "开始检测系统环境."],
	 	[
			"ADOX.Catalog", 
			"Adodb.RecordSet",
			"Adodb.Stream",
			"Microsoft.XMLHTTP",
			"Scripting.FileSystemObject",
			"Microsoft.XMLDOM",
			"Scripting.Dictionary",
			"Msxml2.ServerXMLHTTP.5.0",
			"MSXML2.ServerXMLHTTP",
			"Msxml2.DOMDocument.5.0",
			"FileUp.upload",
			"JMail.SMTPMail",
			"GflAx190.GflAx",
			"easymail.Mailsend"
		]
	],
	
	// step two for pjblog, Create Database!
	[
		[2, "开始创建数据库"],
		[
		 	// datebase path for pjblog
		 	"../" + AccessPath // 数据库地址,按需要修改.
		]
	],
	
	// step three for pjblog, update datebase!
	[
	 	[3, "开始操作执行升级数据库"],
		[NewArray_0, NewArray_1, NewArray_2, NewArray_3, NewArray_4, NewArray_5, NewArray_6, NewArray_7, NewArray_8, NewArray_9, NewArray_10, NewArray_11, NewArray_12, NewArray_13, NewArray_14, NewArray_15, NewArray_16, NewArray_17]
	],
	
	// step four for pjblog, clear Application
	[
	 	[4, "开始清理全局缓存"],
		[
		 	cookieName + "_blog_Infos",
			cookieName + "_blog_rights",
			cookieName + "_blog_Category",
			cookieName + "_blog_archive",
			cookieName + "_blog_Comment",
			cookieName + "_blog_Tags",
			cookieName + "_blog_Smilies",
			cookieName + "_blog_Keywords",
			cookieName + "_blog_Bloglinks",
			cookieName + "_blog_ReBuildID",
			cookieName + "_blog_module"
		]
	]
];

var clong;
var canSart = true;

var install = {
	
	// 开始执行的方法
	Start : function(obj){
		if (!canSart) {try{$("button").disabled = false;return}catch(e){}};
		$(obj).innerHTML = "";
		var installArray = this.Step;
		var count, titles, smallArray, DataArray, GArray, Dength;
		var Hore = [];
		var a, b, c, d;
		for (a = 0 ; a < installArray.length ; a++){
			// 得到步数和标题
			count = installArray[a][0][0];
			titles = installArray[a][0][1];
			smallArray = installArray[a][1];
			this.insTit(obj, count, titles);
			for (b = 0 ; b < smallArray.length ; b++){
				if (count == 3){
					DataArray = smallArray[b];
					for (c = 0 ; c < DataArray.length ; c++){
						GArray = [count, DataArray[c]];
						Hore.push(GArray);
					}
				}else{
					GArray = [count, smallArray[b]];
					Hore.push(GArray);
				}	
			}
		}
		clong = Hore;
		Dength = clong.length
		if (Dength > 0){this.SecAjax(clong[0][0], clong[0][1], 0, Dength)}
	},
	
	// 步骤数组
	Step : Arrays,

	// Ajax加载方法
	SecAjax : function(step, info, i, Len){
		Ajax({
			async : true,
			url : "Action.asp?TimeStamp=" + new Date().getTime(),
			method : "POST",
			content : "step=" + step + "&info=" + escape(info),
			oncomplete : function(obj){
				var str = obj.responseText.json();
				if (str.suc){
					install.insCom(step, str.info);
					$("Ajax").scrollTop = parseInt($("Ajax").scrollHeight);
				}else{install.insCom(step, str.info);$("Ajax").scrollTop = parseInt($("Ajax").scrollHeight);}
				i += 1;
				var pwidth = $("processBar").offsetWidth - 6;
				var percent = ((i + 1) / Len) * 100;
				var percentWidth = ((i + 1) / Len) * pwidth;
				$("number").innerHTML = parseInt(percent) + "%";
				$("process").style.width = percentWidth + "px";
				if (i <= Len) install.SecAjax(clong[i][0], clong[i][1], i, Len)
				if (i == (Len - 1)) {
					$("button").disabled = false;
					$("button").innerHTML = "下一步 - 清理临时文件";
					$("button").href = "?misc=complate";
					canSart = false;
					$("button").onclick = function(){return true;}	
				}
			},
			ononexception:function(obj){
			}
		});
	},
	
	// 插入标题
	insTit : function(obj, step, title){
		try{
			var _ul = document.createElement("ul");
			$(obj).appendChild(_ul);
			_ul.id = "step_" + step;
			_ul.innerHTML = title;
		}catch(e){}
	},
	
	//插入内容
	insCom : function(step, info){
		var obj = "step_" + step;
		try{
			var _li = document.createElement("li");
			$(obj).appendChild(_li);
			_li.innerHTML = info;
		}catch(e){}
	},
	
	//我同意
	Agree : function(obj){
		if (obj.checked){
			$("next").disabled = false;
			$("next").className = "Textbutton";
			$("next").onclick = function(){
				cookie.SET("InstallCookie", $("PostCookie").Install_Cookie.value, 1);
				cookie.SET("InstallAccess", $("PostCookie").Install_Access.value, 1);
				return true;
			}
		}else{
			$("next").disabled = true;
			$("next").className = "Textbuttons";
			$("next").onclick = function(){return false;}
		}
	}
}


var File = {
	Start : function(){
		this.insTit("Ajax"); // 插入标题
		var clen = xml;
		this.SecAjax(0, clen);		
	},
	
	// Ajax加载方法
	SecAjax : function(i, Len){
		Ajax({
			async : true,
			url : "Action.asp?TimeStamp=" + new Date().getTime(),
			method : "POST",
			content : "step=0&id=" + i,
			oncomplete : function(obj){
				var str = obj.responseText.json();
				if (str.suc){
					File.insCom("SolveFiles", str.info);
					$("Ajax").scrollTop = parseInt($("Ajax").scrollHeight);
				}else{File.insCom("SolveFiles", str.info);$("Ajax").scrollTop = parseInt($("Ajax").scrollHeight);}
				i += 1;
				var pwidth = $("processBar").offsetWidth - 6;
				var percent = ((i + 1) / Len) * 100;
				var percentWidth = ((i + 1) / Len) * pwidth;
				$("number").innerHTML = parseInt(percent) + "%";
				$("process").style.width = percentWidth + "px";
				if (i <= Len) File.SecAjax(i, Len)
				if (i == Len) {
					$("button").disabled = false;
					$("button").innerHTML = "下一步 - 安装升级数据库";
					$("button").href = "?misc=install";
					canSart = false;
					$("button").onclick = function(){return true;}
				}
			},
			ononexception:function(obj){
			}
		});
	},
	
	//插入标题
	insTit : function(obj){
		var _ul = document.createElement("ul");
		$(obj).appendChild(_ul);
		_ul.id = "SolveFiles";
		_ul.innerHTML = "开始解压程序文件....";
	},
	
	//插入内容
	insCom : function(obj, html){
		var _div = document.createElement("li");
		$(obj).appendChild(_div);
		_div.innerHTML = html;
	}
}

/***************************************
	绝对坐标
****************************************/
  function ABS(a){
	  a = $(a);
	  var b = { x: a.offsetLeft, y: a.offsetTop};
	  a = a.offsetParent;
	  while (a) {
		  b.x += a.offsetLeft;
		  b.y += a.offsetTop;
		  a = a.offsetParent;
	  }
	  return b;
  } 
  
  
  function FollowBox(obj, obj2){
	  var Box = $(obj);
	  var Ps = ABS(Box);
	  var objWidth = Box.offsetWidth;
	  var objHeight = Box.offsetHeight;
	  $(obj2).style.cssText += ";left:" + Ps.x + "px; top:" + (objHeight - 168) + "px;";
  }