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

<!--#include file="FCKeditor/fckeditor.asp" -->
<%
'***************PJblog3 后台管理页面*******************
' PJblog3 Copyright 2008
' Update:2007-7-13
'**************************************************
'开始加载后台需要的各个模块

%>
<!--#include file="control/f_control.asp" -->
<!--#include file="control/c_welcome.asp" -->
<!--#include file="control/c_general.asp" -->
<!--#include file="control/c_categories.asp" -->
<!--#include file="control/c_comment.asp" -->
<!--#include file="control/c_skins.asp" -->
<!--#include file="control/c_SQLFile.asp" -->
<!--#include file="control/c_members.asp" -->
<!--#include file="control/c_article.asp" -->
<!--#include file="control/c_link.asp" -->
<!--#include file="control/c_smilies.asp" -->
<!--#include file="control/c_status.asp" -->
<!--#include file="control/c_codeEditor.asp" -->
<!--#include file="control/Action.asp" -->
<%
'定义变量
Dim i

Dim blog_module
Dim PluginsFolders, PluginsFolder, Bmodules, Bmodule, tempB, SubItemLen, tempI
Dim PluginsXML, DBXML, TypeArray

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
      Select Case Request.QueryString("Fmenu") 
           Case "welcome"  Call c_welcome '欢迎页面
           Case "General"  Call c_ceneral  ' 站点基本设置
           Case "Categories" Call c_categories ' 日志分类管理
           Case "Article"  Call c_article '日志管理
           Case "Comment"  Call c_comment ' 评论留言管理
           Case "Skins"  Call c_skins ' 界面设置
           Case "SQLFile"  Call c_SQLFile ' 数据库与文件
           Case "Members"  Call c_members ' 帐户和权限设置
           Case "Link"  Call c_link ' 友情链接管理
           Case "smilies"  Call c_smilies ' 表情和关键字
           Case "Status"  Call c_status ' 服务器配置信息
           Case "CodeEditor"  Call c_codeEditor '代码编辑器
           Case "Logout"  Call c_Logout ' 退出后台管理
           Case Else Call doAction '开始执行保存设置
      End Select
	%>
  </div>
 </body>
</html>
