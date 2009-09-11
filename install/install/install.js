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

/*****************************************************************/
// 	系统字段开始
/*****************************************************************/

var Tables = [
	"create table blog_book",
	"create table blog_Category",
	"create table blog_Comment",
	"create table blog_Content",
	"create table blog_Counter",
	"create table blog_Files",
	"create table blog_Info",
	"create table blog_Keywords",
	"create table blog_LinkClass",
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

var blog_book = [
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
	"ALTER TABLE [blog_book] ADD COLUMN book_replyAuthor varchar(20)",
	"ALTER TABLE [blog_book] ADD COLUMN email varchar(50)",
	"ALTER TABLE [blog_book] ADD COLUMN siteurl varchar(50)"
];

var blog_Category = [
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

var blog_Comment = [
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
	"ALTER TABLE [blog_Comment] ADD COLUMN comm_IsAudit bit default False",
	"ALTER TABLE [blog_Comment] ADD COLUMN comm_Email varchar(255)",
	"ALTER TABLE [blog_Comment] ADD COLUMN comm_WebSite varchar(255)"
];

var blog_Content = [
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

var blog_Counter = [
	"ALTER TABLE [blog_Counter] ADD COLUMN coun_ID COUNTER(1, 1) PRIMARY KEY",
	"ALTER TABLE [blog_Counter] ADD COLUMN coun_IP varchar(20)",
	"ALTER TABLE [blog_Counter] ADD COLUMN coun_OS varchar(50)",
	"ALTER TABLE [blog_Counter] ADD COLUMN coun_Browser varchar(50)",
	"ALTER TABLE [blog_Counter] ADD COLUMN coun_Time smalldatetime default Now()",
	"ALTER TABLE [blog_Counter] ADD COLUMN coun_Referer varchar(255)"
];

var blog_Files = [
	"ALTER TABLE [blog_Files] ADD COLUMN ID COUNTER(1, 1) PRIMARY KEY",
	"ALTER TABLE [blog_Files] ADD COLUMN FilesPath varchar(255)",
	"ALTER TABLE [blog_Files] ADD COLUMN FilesCounts Integer default 0"
];

var blog_Info = [
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
	"ALTER TABLE [blog_Info] ADD COLUMN blog_AuditOpen bit default False",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_GravatarOpen bit default False",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_Isjmail bit default False",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_reply_Isjmail bit default False",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_smtp varchar(50)",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_smtpuser varchar(50)",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_smtppassword varchar(50)",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_jmail varchar(50)",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_smtpmail varchar(50)",
	"ALTER TABLE [blog_Info] ADD COLUMN blog_html ntext"
];

var blog_Keywords = [
	"ALTER TABLE [blog_Keywords] ADD COLUMN key_ID COUNTER(1, 1) PRIMARY KEY",
	"ALTER TABLE [blog_Keywords] ADD COLUMN key_Text varchar(100)",	
	"ALTER TABLE [blog_Keywords] ADD COLUMN key_URL varchar(255)",
	"ALTER TABLE [blog_Keywords] ADD COLUMN key_Image varchar(50)"
];

var blog_LinkClass = [
	"ALTER TABLE [blog_LinkClass] ADD COLUMN LinkClass_ID COUNTER(1, 1) PRIMARY KEY",
	"ALTER TABLE [blog_LinkClass] ADD COLUMN LinkClass_Name varchar(50)",		
	"ALTER TABLE [blog_LinkClass] ADD COLUMN LinkClass_Title varchar(255)",
	"ALTER TABLE [blog_LinkClass] ADD COLUMN LinkClass_Order Integer default 0"
];

var blog_Links = [
	"ALTER TABLE [blog_Links] ADD COLUMN link_ID COUNTER(1, 1) PRIMARY KEY",
	"ALTER TABLE [blog_Links] ADD COLUMN link_Name varchar(50)",
	"ALTER TABLE [blog_Links] ADD COLUMN link_URL varchar(200)",
	"ALTER TABLE [blog_Links] ADD COLUMN link_Image varchar(250)",
	"ALTER TABLE [blog_Links] ADD COLUMN link_Order Short default 0",
	"ALTER TABLE [blog_Links] ADD COLUMN link_IsShow bit default False",
	"ALTER TABLE [blog_Links] ADD COLUMN link_IsMain bit default False",
	"ALTER TABLE [blog_Links] ADD COLUMN Link_ClassID Integer default 0"
];

var blog_Member = [
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

var blog_ModSetting = [
	"ALTER TABLE [blog_ModSetting] ADD COLUMN set_id COUNTER(1, 1) PRIMARY KEY",
	"ALTER TABLE [blog_ModSetting] ADD COLUMN set_ModName varchar(50)",
	"ALTER TABLE [blog_ModSetting] ADD COLUMN set_KeyName varchar(50)",
	"ALTER TABLE [blog_ModSetting] ADD COLUMN set_KeyValue ntext"
];

var blog_module = [
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

var blog_Notdownload = [
	"ALTER TABLE [blog_Notdownload] ADD COLUMN blog_Nodownload OLEObject"			   
];

var blog_Smilies = [
	"ALTER TABLE [blog_Smilies] ADD COLUMN sm_ID COUNTER(1, 1) PRIMARY KEY",
	"ALTER TABLE [blog_Smilies] ADD COLUMN sm_Image varchar(25)",
	"ALTER TABLE [blog_Smilies] ADD COLUMN sm_Text varchar(25)"
];

var blog_status = [
	"ALTER TABLE [blog_status] ADD COLUMN stat_id COUNTER(1, 1) PRIMARY KEY",
	"ALTER TABLE [blog_status] ADD COLUMN stat_name varchar(50)",
	"ALTER TABLE [blog_status] ADD COLUMN stat_title varchar(255)",
	"ALTER TABLE [blog_status] ADD COLUMN stat_Code varchar(50)",
	"ALTER TABLE [blog_status] ADD COLUMN stat_attSize Integer default 0",
	"ALTER TABLE [blog_status] ADD COLUMN stat_attType varchar(255)"
];

var blog_tag = [
	"ALTER TABLE [blog_tag] ADD COLUMN tag_id COUNTER(1, 1) PRIMARY KEY",
	"ALTER TABLE [blog_tag] ADD COLUMN tag_name varchar(20)",
	"ALTER TABLE [blog_tag] ADD COLUMN tag_count Integer default 0"
];

var blog_Trackback = [
	"ALTER TABLE [blog_Trackback] ADD COLUMN tb_ID COUNTER(1, 1) PRIMARY KEY",
	"ALTER TABLE [blog_Trackback] ADD COLUMN blog_ID Integer default 0",
	"ALTER TABLE [blog_Trackback] ADD COLUMN tb_URL varchar(255)",
	"ALTER TABLE [blog_Trackback] ADD COLUMN tb_Title varchar(100)",
	"ALTER TABLE [blog_Trackback] ADD COLUMN tb_Site varchar(50)",
	"ALTER TABLE [blog_Trackback] ADD COLUMN tb_SiteURL varchar(100)",
	"ALTER TABLE [blog_Trackback] ADD COLUMN tb_Intro ntext",
	"ALTER TABLE [blog_Trackback] ADD COLUMN tb_PostTime smalldatetime default Now()"
];

/*****************************************************************/
// 	系统字段结束
/*****************************************************************/


/*****************************************************************/
// 	插入系统字段内容开始
/*****************************************************************/

var insSysTem = [
	// ------------------------------- blog_Category -----------------------------------
	"insert into blog_Category (cate_Name, cate_Order, cate_Intro, cate_OutLink, cate_URL, cate_Lock, cate_icon) values ('Index', -99, '日志首页', True, 'default.asp', True, 'images/icons/22.gif')",
	"insert into blog_Category (cate_Name, cate_OutLink, cate_URL, cate_Lock, cate_icon) values ('TagsCloud', True, 'tag.asp', False, 'images/icons/10.gif')",
	// ------------------------------- blog_info -----------------------------------
	"insert into blog_Info (blog_Name, blog_Title, blog_URL, blog_MemNums, blog_showtotal, blog_FilterName, blog_commTimerout, blog_commImg, blog_SplitType, blog_introChar, blog_postFile, blog_DisMod, blog_master, blog_CountNum, blog_wap, blog_wapNum, blog_wapHTML, blog_wapLogin, blog_UpLoadSet, blog_html) values ('PJBlog3', '创造机会的人是勇者；等待机会的人是愚者', 'http://www.pjhome.net/', 1, 0, '游客|客人|Admin|SupAdmin|Fuck|', 30, 1, True, 2000, 2, 0, 'YourName', 0, -1, 5, -1, -1, '1|4|0|PJBlog|PJBlog|9|0|10|10|FFFFFF|0|10|10|0.5|images/WaterMaker.png|280|45|www.pjhome.net|FFFFFF|18|宋体|1|0|000000|0|0', 'html,htm,shtml,xhtml,do')",
	// ------------------------------- blog_Keywords -----------------------------------
	"insert into blog_Keywords (key_Text, key_URL) values ('puterjam', 'http://www.pjhome.net')",
	"insert into blog_Keywords (key_Text, key_URL) values ('PJBlog', 'http://www.pjhome.net')",
	"insert into blog_Keywords (key_Text, key_URL) values ('pj', 'http://www.pjhome.net')",
	// ------------------------------- blog_LinkClass -----------------------------------
	"insert into blog_LinkClass (LinkClass_Name, LinkClass_Title, LinkClass_Order) values ('pjblog', 'pjblog', 1)",
	// ------------------------------- blog_Member -----------------------------------
	"insert into blog_Member (mem_Name, mem_Password, mem_salt, mem_HideEmail, mem_RegTime, mem_Status, mem_LastIP, mem_hashKey) values ('admin', '9d787f1f9cd95ec1b97c13644e6a15c4a0fcac8c', '2gg905', True, '2004/11/8 15:37:53', 'SupAdmin', '127.0.0.1', 'ece268b25ba022d553acbaa18c282bb05133ac40')",
	// ------------------------------- blog_Member -----------------------------------
	"insert into blog_status (stat_name, stat_title, stat_Code, stat_attSize, stat_attType) values ('SupAdmin', '超级管理员', '111111111111', 1024000000, 'RAR|ZIP|SWF|JPG|PNG|GIF|DOC|TXT|CHM|PDF|ACE|JPG|MP3|WMA|WMV|MIDI|AVI|RM|RA|RMVB|MOV|TORRENT|XLS')",
	"insert into blog_status (stat_name, stat_title, stat_Code, stat_attSize, stat_attType) values ('Member', '普通会员', '000000110000', 250000, 'JPG|PNG|GIF|JPEG')",
	"insert into blog_status (stat_name, stat_title, stat_Code, stat_attSize, stat_attType) values ('Guest', '游客', '000000110000', 5000, 'JPG|PNG|GIF|JPEG')",
	// ------------------------------- blog_Smilies -----------------------------------
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_01.gif', '[face01]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_02.gif', '[face02]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_03.gif', '[face03]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_04.gif', '[face04]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_05.gif', '[face05]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_06.gif', '[face06]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_07.gif', '[face07]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_08.gif', '[face08]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_09.gif', '[face09]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_10.gif', '[face10]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_11.gif', '[face11]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_12.gif', '[face12]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_13.gif', '[face13]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_14.gif', '[face14]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_15.gif', '[face15]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_16.gif', '[face16]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_17.gif', '[face17]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_18.gif', '[face18]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_19.gif', '[face19]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_20.gif', '[face20]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_21.gif', '[face21]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_22.gif', '[face22]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_23.gif', '[face23]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_24.gif', '[face24]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_25.gif', '[face25]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_26.gif', '[face26]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_27.gif', '[face27]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_28.gif', '[face28]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_29.gif', '[face29]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_30.gif', '[face30]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_31.gif', '[face31]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_32.gif', '[face32]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_33.gif', '[face33]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_34.gif', '[face34]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_35.gif', '[face35]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_36.gif', '[face36]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_37.gif', '[face37]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_38.gif', '[face38]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_39.gif', '[face39]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_40.gif', '[face40]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_41.gif', '[face41]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_42.gif', '[face42]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_43.gif', '[face43]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_44.gif', '[face44]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_45.gif', '[face45]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_46.gif', '[face46]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_47.gif', '[face47]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_48.gif', '[face48]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_49.gif', '[face49]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_50.gif', '[face50]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_51.gif', '[face51]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_52.gif', '[face52]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_53.gif', '[face53]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_54.gif', '[face54]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_55.gif', '[face55]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_56.gif', '[face56]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_57.gif', '[face57]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_58.gif', '[face58]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_59.gif', '[face59]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_60.gif', '[face60]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_61.gif', '[face61]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_62.gif', '[face62]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_63.gif', '[face63]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_64.gif', '[face64]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_65.gif', '[face65]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_66.gif', '[face66]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_67.gif', '[face67]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_68.gif', '[face68]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_69.gif', '[face69]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_70.gif', '[face70]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_71.gif', '[face71]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_72.gif', '[face72]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_73.gif', '[face73]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_74.gif', '[face74]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_75.gif', '[face75]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_76.gif', '[face76]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_77.gif', '[face77]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_78.gif', '[face78]')",
	"insert into blog_Smilies (sm_Image, sm_Text) values ('Face_79.gif', '[face79]')",
	// ------------------------------- blog_Notdownload -----------------------------------
	"insert into blog_Notdownload",
	// ------------------------------- blog_module -----------------------------------
	"insert into blog_module (name, title, type, IndexOnly, SortID, IsSystem, HtmlCode) values ('User', 'User Panel', 'sidebar', False, 3, True, '$user_code$')",
	"insert into blog_module (name, title, type, IndexOnly, SortID, IsSystem, HtmlCode) values ('BlogInfo', 'Statistics', 'sidebar', True, 4, True, '\"日志: <a href=\"\"default.asp\"\"><b>$blog_LogNums$</b> 篇</a><br/>评论: <a href=\"\"search.asp?searchType=Comments\"\"><b>$blog_CommNums$</b> 个</a><br/>引用: <a href=\"\"search.asp?searchType=trackback\"\"><b>$blog_TbCount$</b> 个</a><br/>留言: <b>$blog_MessageNums$</b> 个<br/>会员: <a href=\"\"member.asp\"\"><b>$blog_MemNums$</b> 人</a><br/>访问: <b>$blog_VisitNums$</b> 次<br/>在线: <b>$blog_OnlineNums$</b> 人<br/>建站时间: <strong>2005-06-20</strong>\"')",
	"insert into blog_module (name, title, type, IndexOnly, SortID, IsSystem, HtmlCode) values ('Calendar', 'Calendar', 'sidebar', False, 1, True, '$calendar_code$')",
	"insert into blog_module (name, title, type, IndexOnly, SortID, IsSystem, HtmlCode) values ('Comment', 'Comment', 'sidebar', False, 5, True, '\"<div class=\"\"commentTable\"\">$comment_code$</div>\"')",
	"insert into blog_module (name, title, type, IndexOnly, SortID, IsSystem, HtmlCode) values ('Search', 'Search', 'sidebar', True, 7, True, '\"<form style=\"\"MARGIN: 0px\"\" onsubmit=\"\"if (this.SearchContent.value.length&lt;3) {alert(&#39;关键字不能少于3个字&#39;);this.SearchContent.focus();return false}\"\" action=\"\"search.asp\"\">关键字 <input class=\"\"userpass\"\" name=\"\"SearchContent\"\"/><div style=\"\"OVERFLOW: hidden; HEIGHT: 3px\"\">&nbsp;</div> 类　型 <sel&#101;ct name=\"\"searchType\"\"><option value=\"\"title\"\" sel&#101;cted=\"\"sel&#101;cted\"\">日志标题</option> <option value=\"\"Content\"\">日志内容</option><option value=\"\"Comments\"\">日志评论</option><option value=\"\"trackback\"\">引用通告</option> </sel&#101;ct> <input class=\"\"userbutton\"\" type=\"\"submit\"\" value=\"\"查 找\"\"/></form>\"')",
	"insert into blog_module (name, title, type, IndexOnly, SortID, IsSystem, HtmlCode) values ('Links', 'Links', 'sidebar', True, 8, True, '\"<div class=\"\"LinkTable\"\">$Link_Code$ </div><div align=\"\"right\"\"><a href=\"\"bloglink.asp\"\">更多链接… </a></div>\"')",
	"insert into blog_module (name, title, type, IndexOnly, SortID, IsSystem, HtmlCode) values ('Support', 'Support', 'sidebar', False, 9, False, '<div style=&#34;PADDING-RIGHT: 4px; PADDING-LEFT: 4px; PADDING-BOTTOM: 4px; PADDING-TOP: 4px; TEXT-ALIGN: left&#34;><a href=&#34;http://validator.w3.org/check/referer&#34; target=&#34;_blank&#34;><img alt=&#34;XHTML 1.0 Transitional&#34; src=&#34;images/xhtml.png&#34; border=&#34;0&#34; /></a> <a href=&#34;http://jigsaw.w3.org/css-validator/validator-uri.html&#34; target=&#34;_blank&#34;><img alt=&#34;Css Validator&#34; src=&#34;images/css.png&#34; border=&#34;0&#34; /></a> <a href=&#34;feed.asp&#34; target=&#34;_blank&#34;><img alt=&#34;RSS 2.0&#34; src=&#34;images/rss2.png&#34; border=&#34;0&#34; /></a>&nbsp;<a href=&#34;atom.asp&#34; target=&#34;_blank&#34;><img alt=&#34;Atom 1.0&#34; src=&#34;images/atom.png&#34; border=&#34;0&#34; /></a> <a href=&#34;http://www.mozilla.org/products/firefox/&#34; target=&#34;_blank&#34;><img alt=&#34;Get firefox&#34; src=&#34;images/firefox.gif&#34; border=&#34;0&#34; /></a> <a href=&#34;http://www.creativecommons.cn/licenses/by-nc-sa/1.0/&#34; target=&#34;_blank&#34;><img alt=&#34;Creative Commons&#34; src=&#34;images/cc.png&#34; border=&#34;0&#34; /></a> </div>')",
	"insert into blog_module (name, title, type, IndexOnly, SortID, IsSystem, HtmlCode) values ('Archive', 'Archive', 'sidebar', False, 6, True, '$archive_code$')",
	"insert into blog_module (name, title, type, IndexOnly, SortID, IsSystem, HtmlCode) values ('Category', 'Category', 'sidebar', False, 0, True, '$Category_code$')",
	"insert into blog_module (name, title, type, IndexOnly, SortID, IsSystem, HtmlCode) values ('ContentList', 'ContentList', 'content', False, 0, True, '$Content_code$')",
	"insert into blog_module (name, title, type, IndexOnly, SortID, IsSystem, HtmlCode) values ('Welcome', 'Welcome', 'content', True, -1, False, '\"<p style=&#34;line-height: 140%&#34; align=&#34;left&#34;><strong>致Blogger:<img style=&#34;width: 96px; height: 96px&#34; alt=&#34;&#34; align=&#34;right&#34; src=&#34;images/welcome.gif&#34; /></strong><br /><br />感谢你选择 <strong><font color=&#34;#006600&#34;>PJBlog3</font></strong> 。第一次使用，请用初始管理员账号登陆系统，然后到系统管理，设置<font color=&#34;#0000ff&#34;><strong>站点基本信息</strong></font>。<br />初次使用还需要到后台建立你的日志的分类。<br />建站时间可以到 <font color=&#34;#0000ff&#34;><strong>界面与插件-设置模块</strong></font> 编辑模块标识为 <strong>BlogInfo</strong> 的模块即可。<br /><br />祝你使用愉快。最后，请到 <font color=&#34;#0000ff&#34;><strong>界面与插件-设置模块</strong></font> 把模块标识为 <strong>Welcome </strong>模块删除。</p><p style=&#34;line-height: 140%&#34; align=&#34;left&#34;>PJBLOG基本使用教程：<a href=&#34;http://bbs.pjhome.net/thread-30520-1-1.html&#34;>http://bbs.pjhome.net/thread-30520-1-1.html</a><br />PJBLOG常见问题集：<a href=&#34;http://bbs.pjhome.net/forum-35-1.html&#34;>http://bbs.pjhome.net/forum-35-1.html</a></p><hr /><p>&nbsp;</p>\"')"
];

/*****************************************************************/
// 	插入系统字段内容结束
/*****************************************************************/

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
		[Tables, blog_book, blog_Category, blog_Comment, blog_Content, blog_Counter, blog_Files, blog_Info, blog_Keywords, blog_LinkClass, blog_Links, blog_Member, blog_ModSetting, blog_module, blog_Notdownload, blog_Smilies, blog_status, blog_tag, blog_Trackback]
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