<!--#include file="../include.asp" --><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="Content-Language" content="UTF-8" />
<meta name="author" content="PuterJam evio" />
<meta name="Copyright" content="PJBlog3 CopyRight 2009 2010" />
<meta name="keywords" content="PuterJam,evio,Blog,ASP,designing,with,web,standards,xhtml,css,graphic,design,layout,usability,ccessibility,w3c,w3,w3cn" />
<meta name="description" content="PuterJam's BLOG" />
<script language="javascript" type="text/javascript" src="../pjblog.common/language.js"></script>
<script language="javascript" type="text/javascript" src="../pjblog.common/jquery/jquery.js"></script>
<script language="javascript" type="text/javascript" src="../pjblog.common/jquery/jquery.form.js"></script>
<script language="javascript" type="text/javascript" src="../pjblog.common/jquery/jquery.alerts.js"></script>
<script language="javascript" type="text/javascript" src="../pjblog.common/jquery/jquery.ui.draggable.js"></script>
<script language="javascript" type="text/javascript" src="../pjblog.common/jquery/jquery.ui.ebox.js"></script>
<script language="javascript" type="text/javascript" src="../pjblog.common/jquery/jquery.ui.slide.js"></script>
<script language="javascript" type="text/javascript" src="../pjblog.common/jquery/jquery.ui.ubbeditor.js"></script>
<script language="javascript" type="text/javascript" src="../pjblog.common/jquery/jquery.ui.selectContent.js"></script>
<script language="javascript" type="text/javascript" src="../pjblog.common/jquery/jquery.ui.domedit.js"></script>
<script language="javascript" type="text/javascript" src="../pjblog.common/jquery/jquery.ui.fonteffect.js"></script>
<script language="javascript" type="text/javascript" src="../pjblog.common/jquery/jquery.ui.scrolling.js"></script>
<script language="javascript" type="text/javascript" src="../pjblog.common/control.js"></script>
<script language="javascript" type="text/javascript" src="../pjblog.common/blog.js"></script>
<script language="javascript" type="text/javascript" src="../pjblog.model/base64.js"></script>
<link rel="stylesheet" rev="stylesheet" href="../pjblog.common/jquery/ubbeditor.css" type="text/css" media="all" />
<link rel="stylesheet" rev="stylesheet" href="../pjblog.common/control.css" type="text/css" media="all" />
<link rel="stylesheet" rev="stylesheet" href="../pjblog.common/jquery/jquery.alerts.css" type="text/css" media="all" />
<link rel="stylesheet" rev="stylesheet" href="../pjblog.common/blog.css" type="text/css" media="all" />
<title>后台管理 - PJBlog4 v<%=Sys.version%></title>
</head>
<!--#include file = "log_control.asp" -->
<!--#include file = "control/c_welcome.asp" -->
<!--#include file = "control/c_general.asp" -->
<!--#include file = "control/c_categories.asp" -->
<!--#include file = "control/c_link.asp" -->
<!--#include file = "control/c_comment.asp" -->
<!--#include file = "control/c_skins.asp" -->
<!--#include file = "control/c_skins_set.asp" -->
<!--#include file = "control/c_skins_net.asp" -->
<!--#include file = "control/c_skins_upd.asp" -->
<!--#include file = "control/c_plugins.asp" -->

<!--#include file = "../pjblog.model/cls_fso.asp" -->
<!--#include file = "../pjblog.model/cls_stream.asp" -->
<!--#include file = "../pjblog.model/cls_xml.asp" -->
<!--#include file = "../pjblog.common/md5.asp" --> 

<%
Dim Menu, SubMenu
	Menu = Request.QueryString("m")
	SubMenu = Request.QueryString("s")
%>
<body>
<div id="ovlayinfo" style=" display:none;">
	<div id="lay_0">这里设置网站的基本信息和你的基本信息</div>
    <div id="lay_1">这里将退出登入,返回首页.</div>
</div>
<div class="wrap">
	<!--顶部导航 开始-->
	<div class="ConMenu">
    	<div class="nav">
<!--        	<ul>
            	<li><a href="">预览首页</a></li>
                <li><a href="?m=default">后台首页</a></li>
                <li><a href="">服务器信息</a></li>
            </ul>-->
           	PJBlog 4 Control Pannel  -  New Offical Studio
        </div>
        <div class="action">
        	<div class="l f">
                <div class="photo-s">
                    <img src="http://www.gravatar.com/avatar/6bab7c91a03d47fe1aa5b5b6b6f8cc55?s=80&r=r&d=http%3a%2f%2fwww.gravatar.com%2favatar%2fad516503a11cd5ca435acc9bb6523536%3fs%3d35">
                </div>
            </div>
            <div class="r f">
            	<span class="mail"><%=blog_master%>(<%=blog_email%>)</span><br />
                <span class="active">任务</span>
                <span class="active"><a href="" class="lay_0">设置</a></span>
                <span class="active"><a href="" class="lay_1">退出</a></span>
            </div>
            <div class="clear"></div>
        </div>
        <div class="clear"></div>
    </div>
    <!--顶部导航 结束-->
    <!--内容部分 开始-->
    <div class="main">
    	<!--内容左边部分 开始-->
    	<div class="menu">
        	<div class="menu_top"></div>
            <div class="menu_mid">
            	<div class="menuList">
                	<ul>
                    	<li><div class="icon"><img src="../images/Control/Icon/zhengzhan.gif"></div><div class="title"><a href="?m=general">整站优化</a></div><div class="do"></div><div class="clear"></div></li>
                        <li><div class="icon"><img src="../images/Control/Icon/toppiao.gif"></div><div class="title"><a href="?m=category">分类管理</a></div><div class="do"><a href="javascript:;" onClick="conMain.category.add(this)">新增</a></div><div class="clear"></div></li>
                        <li><div class="icon"><img src="../images/Control/Icon/rizhi.gif"></div><div class="title"><a href="">日志管理</a></div><div class="do"><a href="javascript:;" onClick="">发表</a></div><div class="clear"></div></li>
                        <li><div class="icon"><img src="../images/Control/Icon/pinglun.gif"></div><div class="title"><a href="">评论管理</a></div><div class="do">[1]</div><div class="clear"></div></li>
                        <li><div class="icon"><img src="../images/Control/Icon/waiguan.gif"></div><div class="title"><a href="?m=skin">外观设置</a></div><div class="do"></div><div class="clear"></div></li>
                        <li><div class="icon"><img src="../images/Control/Icon/chajian.gif"></div><div class="title"><a href="?m=plus">插件管理</a></div><div class="do"></div><div class="clear"></div></li>
                        <li><div class="icon"><img src="../images/Control/Icon/fujian.gif"></div><div class="title"><a href="">附件管理</a></div><div class="do"></div><div class="clear"></div></li>
                        <li><div class="icon"><img src="../images/Control/Icon/zhanghu.gif"></div><div class="title"><a href="">账户权限</a></div><div class="do"></div><div class="clear"></div></li>
                        <li><div class="icon"><img src="../images/Control/Icon/link.gif"></div><div class="title"><a href="?m=link">友情链接</a></div><div class="do"><a href="javascript:;" onClick="conMain.links.addNewLink(this, 6)">新增</a></div><div class="clear"></div></li>
                        <li><div class="icon"><img src="../images/Control/Icon/qita.gif"></div><div class="title"><a href="">其他管理</a></div><div class="do"></div><div class="clear"></div></li>
                        <li style="border-top:1px dashed #8ec5f7; margin-top:10px; padding-top:10px;"><div class="icon"><img src="../images/Control/Icon/home.gif"></div><div class="title"><a href="">前台预览</a></div><div class="do"></div><div class="clear"></div></li>
                        <li><div class="icon"><img src="../images/Control/Icon/yulan.gif"></div><div class="title"><a href="?m=default">后台首页</a></div><div class="do"></div><div class="clear"></div></li>
                        <li><div class="icon"><img src="../images/Control/Icon/fuwuqi.gif"></div><div class="title"><a href="">服 务 器</a></div><div class="do"></div><div class="clear"></div></li>
                        <li><div class="icon"><img src="../images/Control/Icon/fangke.gif"></div><div class="title"><a href="">访客记录</a></div><div class="do"></div><div class="clear"></div></li>
                    </ul>
                </div>
            </div>
            <div class="menu_bom"></div>
        </div>
       <!-- 内容左边部分 结束
        内容右边部分 开始-->
        <div class="content">
        <%
			Select Case Menu
				Case "default" Call c_welcome
				Case "general" Call c_ceneral
				Case "category" Call c_categories
				Case "skin" 
					Select Case SubMenu
						Case "set" Call Skin_Set
						Case "net" Call Skin_Net
						Case "upd" Call Skin_Upload
						Case Else Call c_skins
					End Select
				Case "plus"
					Select Case SubMenu
						Case Else Call c_plugins
					End Select
				Case "link" Call c_link
				Case Else Call c_welcome
			End Select
		%>
        </div>
        <!--内容右边部分 结束-->
        <div class="clear"></div>
    </div>
    <!--内容部分 结束-->
</div>
</body>
</html>
<%
Set Sys = Nothing
Public Function yellowBox(ByVal Str)
	yellowBox = ""
	yellowBox = yellowBox & "<div class=""yellowbox"" style=""display:none"">"
	yellowBox = yellowBox & "<div class=""rtxt"">" & Str & "</div>"
	yellowBox = yellowBox & "</div>"
End Function
%>