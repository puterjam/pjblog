<!--#include file="conCommon.asp" -->
<!--#include file="common/function.asp" -->
<!--#include file="common/library.asp" -->
<!--#include file="common/cache.asp" -->
<!--#include file="common/checkUser.asp" -->
<!--#include file="common/ubbcode.asp" -->
<!--#include file="common/XML.asp" -->
<!--#include file="common/ModSet.asp" -->
<!--#include file="class/cls_logAction.asp" -->
<!--#include file="class/cls_article.asp" -->
<!--#include file="class/cls_control.asp" -->
<!--#include file="FCKeditor/fckeditor.asp" -->
<%
'***************PJblog3 后台管理页面*******************
' PJblog3 Copyright 2008
' Update:2007-7-13
'**************************************************
'开始加载后台需要的各个模块

%>
<!--#include file="control/c_welcome.asp" -->
<!--#include file="control/c_general.asp" -->
<!--#include file="control/c_categories.asp" -->
<!--#include file="control/c_comment.asp" -->
<!--#include file="control/c_skins.asp" -->
<!--#include file="control/c_SQLFile.asp" -->
<!--#include file="control/c_members.asp" -->
<!--#include file="control/c_link.asp" -->
<!--#include file="control/c_smilies.asp" -->
<!--#include file="control/c_status.asp" -->
<!--#include file="control/Action.asp" -->
<%
'定义变量
Dim i

'判断是否有权限访问
If Not ChkPost() OR session(CookieName&"_System") <> True OR isEmpty(memName) Or stat_Admin <> True Then
	Call c_Logout
    Response.end
End If
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="UTF-8">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
	<meta http-equiv="Content-Language" content="UTF-8" />
	<meta name="author" content="PuterJam" />
	<meta name="Copyright" content="PJBlog3 CopyRight 2008" />
	<meta name="keywords" content="PuterJam,Blog,ASP,designing,with,web,standards,xhtml,css,graphic,design,layout,usability,ccessibility,w3c,w3,w3cn" />
	<meta name="description" content="PuterJam's BLOG" />
	<link rel="stylesheet" rev="stylesheet" href="common/control.css" type="text/css" media="all" />
	<script type="text/javascript" src="common/control.js"></script>
	<title>后台管理-内容</title>
	<style type="text/css">
		<!--
		.style1 {color: #999}
		-->
	</style>
</head>
<body class="ContentBody">
  <div class="MainDiv">
	<%
	If Request.QueryString("Fmenu") = "General" Then ' 站点基本设置
		Call c_ceneral
	ElseIf Request.QueryString("Fmenu") = "Categories" Then ' 日志分类管理
		Call c_categories
	ElseIf Request.QueryString("Fmenu") = "Comment" Then ' 评论留言管理
		Call c_comment
	ElseIf Request.QueryString("Fmenu") = "Skins" Then ' 界面设置
		Call c_skins
	ElseIf Request.QueryString("Fmenu") = "SQLFile" Then ' 数据库与文件
		Call c_SQLFile
	ElseIf Request.QueryString("Fmenu") = "Members" Then ' 帐户和权限设置
		Call c_members
	ElseIf Request.QueryString("Fmenu") = "Link" Then ' 友情链接管理
		Call c_link
	ElseIf Request.QueryString("Fmenu") = "smilies" Then ' 表情和关键字
		Call c_smilies
	ElseIf Request.QueryString("Fmenu") = "Status" Then ' 服务器配置信息
		Call c_status
	ElseIf Request.QueryString("Fmenu") = "Logout" Then ' 退出后台管理
		Call c_Logout
	ElseIf Request.QueryString("Fmenu") = "welcome" Then '欢迎页面
		Call c_welcome
	Else
		Call doAction
	End If
	%>
  </div>
 </body>
</html>
