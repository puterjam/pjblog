<!--#include file="conCommon.asp" -->
<!--#include file="common/function.asp" -->
<!--#include file="common/library.asp" -->
<!--#include file="common/cache.asp" -->
<!--#include file="common/checkUser.asp" -->
<%
'==================================
'  后台管理菜单部分
'    更新时间: 2005-7-3
'==================================
If session(CookieName&"_System") = True And memName<>Empty And stat_Admin = True Then
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="UTF-8">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<meta http-equiv="Content-Language" content="UTF-8" />
<meta name="author" content="puter2001@21cn.com,PuterJam" />
<meta name="Copyright" content="PL-Blog 2 CopyRight 2005" />
<meta name="keywords" content="PuterJam,Blog,ASP,designing,with,web,standards,xhtml,css,graphic,design,layout,usability,ccessibility,w3c,w3,w3cn" />
<meta name="description" content="PuterJam's BLOG" />
<link rel="stylesheet" rev="stylesheet" href="common/control.css" type="text/css" media="all" />
<title>后台管理-菜单</title>
<script>
var LastItem=null
function MenuClick(obj,url){
 if (LastItem!=null){
  LastItem.className="menuA"
 }
 obj.className="menuAS"
 LastItem=obj
 obj.blur()
 if (url.length>0) parent.MainContent.location=url;
}
</script>
<style>
html{
	overflow:hidden;
}
</style>
</head>
<body class="menuBody" onLoad="MenuClick(document.getElementById('index'),'ConContent.asp?Fmenu=welcome')">

 <ul class="menuUL">
   <li><a id="index" href="javascript:void(0)" class="menuA" onClick="MenuClick(this,'ConContent.asp?Fmenu=welcome')"><img src="images/Control/icon/b.gif" alt="" border="0" class="MenuIcon icon0"/>后台首页</a></li>
   <li><a href="javascript:void(0)" class="menuA" onClick="MenuClick(this,'ConContent.asp?Fmenu=General&Smenu=')"><img src="images/Control/icon/b.gif" alt="" border="0" class="MenuIcon icon2"/>站点基本设置</a></li>
   <li><a href="javascript:void(0)" class="menuA" onClick="MenuClick(this,'ConContent.asp?Fmenu=Categories&Smenu=')"><img src="images/Control/icon/b.gif" alt="" border="0" class="MenuIcon icon4"/>日志分类管理</a></li>
   <li><a href="javascript:void(0)" class="menuA" onClick="MenuClick(this,'ConContent.asp?Fmenu=Article&Smenu=')"><img src="images/Control/icon/b.gif" alt="" border="0" class="MenuIcon icon11"/>日志管理</a></li>
   <li><a href="javascript:void(0)" class="menuA" onClick="MenuClick(this,'ConContent.asp?Fmenu=Comment&Smenu=')"><img src="images/Control/icon/b.gif" alt="" border="0" class="MenuIcon icon10"/>评论留言管理</a></li>
   <li><a href="javascript:void(0)" class="menuA" onClick="MenuClick(this,'ConContent.asp?Fmenu=Skins&Smenu=')"><img src="images/Control/icon/b.gif" alt="" border="0" class="MenuIcon icon5"/>界面与插件</a></li>
   <li><a href="javascript:void(0)" class="menuA" onClick="MenuClick(this,'ConContent.asp?Fmenu=SQLFile&Smenu=')"><img src="images/Control/icon/b.gif" alt="" border="0" class="MenuIcon icon3"/>数据库与附件</a></li>
   <li><a href="javascript:void(0)" class="menuA" onClick="MenuClick(this,'ConContent.asp?Fmenu=Members&Smenu=')"><img src="images/Control/icon/b.gif" alt="" border="0" class="MenuIcon icon6"/>帐户与权限</a></li>
   <li><a href="javascript:void(0)" class="menuA" onClick="MenuClick(this,'ConContent.asp?Fmenu=Link&Smenu=')"><img src="images/Control/icon/b.gif" alt="" border="0" class="MenuIcon icon7"/>友情链接管理</a></li>
   <li><a href="javascript:void(0)" class="menuA" onClick="MenuClick(this,'ConContent.asp?Fmenu=smilies&Smenu=')"><img src="images/Control/icon/b.gif" alt="" border="0" class="MenuIcon icon8"/>表情与关键字</a></li>
   <li><a href="javascript:void(0)" class="menuA" onClick="MenuClick(this,'ConContent.asp?Fmenu=Status&Smenu=')"><img src="images/Control/icon/b.gif" alt="" border="0" class="MenuIcon icon1"/>服务器信息</a></li>
   <!--<li><a href="javascript:void(0)" class="menuA" onclick="MenuClick(this,'ConContent.asp?Fmenu=Logout&Smenu=')"><img src="images/Control/icon/icon9.gif" alt="" border="0" class="MenuIcon"/>退出</a></li>-->
</ul>
 </body>
</html>
<%
Else
    RedirectUrl("default.asp")
End If
%>
