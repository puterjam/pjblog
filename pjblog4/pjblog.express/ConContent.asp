<!--#include file="../include.asp" -->
<%
'*************** PJblog4 基本设置 *******************
' PJblog4 Copyright 2009
' Update: 2009-11-03
' Author: evio
' Owner : PuterJam
' Mode  : Class
' MoName: control
'***************************************************
'判断是否有权限访问
If Not Asp.ChkPost() OR session(Sys.CookieName&"_System") <> True OR isEmpty(memName) Or stat_Admin <> True Then
	Call Init.c_Logout
    Response.end
End If
'开始加载后台需要的各个模块
%>
<!--#include file = "log_control.asp" -->
<!--#include file = "control/c_welcome.asp" -->
<!--#include file = "control/c_general.asp" -->
<!--#include file = "control/c_categories.asp" -->
<!--#include file = "control/c_skins.asp" -->

<%
' 加载功能类
%>
<!--#include file = "../pjblog.model/cls_fso.asp" -->
<!--#include file = "../pjblog.model/cls_stream.asp" -->
<!--#include file = "../pjblog.model/cls_xml.asp" -->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="UTF-8">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
	<meta http-equiv="Content-Language" content="UTF-8" />
	<meta name="author" content="PuterJam" />
	<meta name="Copyright" content="PJBlog3 CopyRight 2008" />
	<meta name="keywords" content="PuterJam,Blog,ASP,designing,with,web,standards,xhtml,css,graphic,design,layout,usability,ccessibility,w3c,w3,w3cn" />
	<meta name="description" content="PuterJam's BLOG" />
	<link rel="stylesheet" rev="stylesheet" href="../pjblog.common/control.css" type="text/css" media="all" />
    <link rel="stylesheet" rev="stylesheet" href="../pjblog.common/common.css" type="text/css" media="all" />
    <script language="javascript" type="text/javascript" src="../pjblog.common/language.js"></script>
    <script language="javascript" type="text/javascript" src="../pjblog.common/jquery.js"></script>
    <script language="javascript" type="text/javascript" src="../pjblog.common/jqueryform.js"></script>
    <script language="javascript" type="text/javascript" src="../pjblog.common/common.js"></script>
    <script language="javascript" type="text/javascript" src="../pjblog.model/base64.js"></script>
	<title>后台管理 - PJBlog4 v<%=Sys.version%></title>
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
