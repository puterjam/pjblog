<!--#include file = "../include.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="UTF-8">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
	<meta http-equiv="Content-Language" content="UTF-8" />
	<meta name="robots" content="all" />
	<meta name="author" content="<%=blog_email%>, <%=blog_master%>" />
	<meta name="Copyright" content="PJBlog3 CopyRight 2008" />
	<meta name="keywords" content="<%=blog_KeyWords%>" />
	<meta name="description" content="<%=blog_Description%>" />
	<meta name="generator" content="PJBlog3" />
	<link rel="service.post" type="application/atom+xml" title="<%=SiteName%> - Atom" href="<%=SiteURL%>xmlrpc.asp" />
	<link rel="EditURI" type="application/rsd+xml" title="RSD" href="<%=SiteURL%>rsd.asp" />
	<title><%=SiteName%> - <%=blog_Title%></title>
	<link rel="alternate" type="application/rss+xml" href="<%=SiteURL%>feed.asp" title="订阅 I Miss You 所有文章(rss2)" />
	<link rel="alternate" type="application/atom+xml" href="<%=SiteURL%>atom.asp" title="订阅 I Miss You 所有文章(atom)" />
	<link rel="stylesheet" rev="stylesheet" href="../pjblog.template/<%=blog_DefaultSkin%>/global.css" type="text/css" media="all" /><!--全局样式表-->
	<link rel="stylesheet" rev="stylesheet" href="../pjblog.template/<%=blog_DefaultSkin%>/layout.css" type="text/css" media="all" /><!--层次样式表-->
	<link rel="stylesheet" rev="stylesheet" href="../pjblog.template/<%=blog_DefaultSkin%>/typography.css" type="text/css" media="all" /><!--局部样式表-->
	<link rel="stylesheet" rev="stylesheet" href="../pjblog.template/<%=blog_DefaultSkin%>/link.css" type="text/css" media="all" /><!--超链接样式表-->
	<link rel="stylesheet" rev="stylesheet" href="../pjblog.template/<%=blog_DefaultSkin%>/UBB/editor.css" type="text/css" media="all" /><!--UBB编辑器代码-->
	<link rel="stylesheet" rev="stylesheet" href="../FCKeditor/editor/css/Dphighlighter.css" type="text/css" media="all" /><!--FCK块引用&代码样式-->
	<link rel="icon" href="favicon.ico" type="image/x-icon" />
	<link rel="shortcut icon" href="../favicon.ico" type="image/x-icon" />
    <script language="javascript" type="text/javascript" src="../pjblog.common/common.js"></script>
	<script language="javascript" type="text/javascript" src="../pjblog.common/Ajax.js"></script>
	<script language="javascript" type="text/javascript" src="../pjblog.common/language.js"></script>
	<script language="javascript" type="text/javascript" src="../pjblog.common/checkform.js"></script>
</head>
<body>
<a href="default.asp" accesskey="i"></a>
<a href="javascript:history.go(-1)" accesskey="z"></a>
<div id="container">
 <!--顶部-->
    <div id="header">
        <div id="blogname"><%=SiteName%><div id="blogTitle"><%=blog_Title%></div></div>
        <div id="menu">
            <div id="Left"></div><div id="Right"></div>
            <%Init.HeadNav%>
        </div>
 	</div>
<!--内容-->
 <div id="Tbody">
<%
'==================================
'  用户登录页面
'    更新时间: 2006-5-29
'==================================
dim Referer_Url
If Request.QueryString("action") = "logout" Then
    Init.logout(True)
    Referer_Url = Cstr(Request.ServerVariables("HTTP_REFERER"))
    If len(Referer_Url) < 8 Then Referer_Url= "http://" & Request.ServerVariables("HTTP_HOST")
%><br/><br/>
   <div style="text-align:center;">
    <div id="MsgContent" style="width:300px">
      <div id="MsgHead"><%=lang.Set.Asp(99)%></div>
      <div id="MsgBody">
		 <div class="MessageIcon"></div>
        <div class="MessageText"><b><%=lang.Set.Asp(100)%></b><br/>
         <a href="default.asp"><%=lang.Set.Asp(15)%></a>&nbsp;|&nbsp;<a href="<%=Referer_Url%>"><%=lang.Set.Asp(101)%></a>
         <br/><%=lang.Set.Asp(102)%></div>
         <meta http-equiv="refresh" content="3;url=<%=Referer_Url%>"/>
		</div>
	  </div>
	</div>
<br/><br/>
 <%
ElseIf Request.Form("action") = "login" Then
    Dim loginUser
    Referer_Url = Session(CookieName & "_Login_Referer_Url")
    If len(Referer_Url) < 8 Then Referer_Url = Cstr(Request.ServerVariables("HTTP_REFERER"))
    If len(Referer_Url) < 8 Then Referer_Url = "http://" & Request.ServerVariables("HTTP_HOST")
    loginUser = login(Request.Form("UserName"), Request.Form("Password"))
%><br/><br/>
   <div style="text-align:center;">
    <div id="MsgContent" style="width:300px">
      <div id="MsgHead"><%=loginUser(0)%></div>
      <div id="MsgBody">
	   <div class="<%=loginUser(2)%>"></div>
       <div class="MessageText"><%=Replace(Replace(loginUser(1),"default.asp",Referer_Url), lang.Tip.SysTem(4) & "</a>", lang.Tip.SysTem(9) & "</a>&nbsp;|&nbsp;<a href=""default.asp"">" & lang.Tip.SysTem(4) & "</a><br/>" & lang.Action.LoginSuc(2))%></div>
	  </div>
	</div>
  </div><br/><br/>
<%
Else
%><br/><br/>
   <div style="text-align:center;">
  <form id="checkUser" action="../pjblog.logic/log_User.asp?action=login" method="post">
    <div id="MsgContent">
      <div id="MsgHead"><%=lang.Set.Asp(17)%></div>
      <div id="MsgBody" style="text-align:left; padding-left:150px;">
	   <label><input name="username" type="text" size="18" class="userpass" maxlength="24"/> <%=lang.Set.Asp(18)%> </label><br/>
	   <label><input name="password" type="password" size="18" class="userpass" onfocus="this.select()"/> <%=lang.Set.Asp(19)%> </label><br/>
	   <label><input name="validate" type="text" size="4" class="userpass" maxlength="4" id="validate" onfocus="CheckForm.User.GetCode('../pjblog.common/GetCode.asp', 'checkcode')" onkeyup="CheckForm.User.CheckCode('validate', 'isok_checkcode')"/> <span id="checkcode"><label style="cursor:pointer;" onClick="CheckForm.User.GetCode('../pjblog.common/GetCode.asp', 'checkcode')"><%=lang.Set.Asp(4)%></label></span> <span id="isok_checkcode"></span></label><br/>
	   　　<label><input name="KeepLogin" type="checkbox" value="1"/><%=lang.Set.Asp(20)%></label><br/>
	   <input type="submit" value="<%=lang.Set.Asp(21)%>" class="userbutton"/>　<input type="button" value="<%=lang.Set.Asp(22)%>" class="userbutton" onclick="location='register.asp'"/>
	   </div>
	</div>
  </form>
  </div><br/><br/>
 <%End if%>
 </div>
<!--底部-->
  <div id="foot">
    <p>Powered By <a href="http://www.pjhome.net" target="_blank"><strong>PJBlog3</strong></a> <a href="http://www.pjhome.net" target="_blank"><strong>V<%=Sys.version%></strong></a> CopyRight 2005 - 2009, <strong><%=SiteName%></strong> <a href="http://validator.w3.org/check/referer" target="_blank">xhtml</a> | <a href="http://jigsaw.w3.org/css-validator/validator-uri.html">css</a></p>
    <p style="font-size:11px;">Processed in <b><%=FormatNumber(Timer()-StartTime,6,-1)%></b> second(s) , <b><%'=SQLQueryNums%></b> queries<%'=SkinInfo%>
    <br/><a href="http://www.miibeian.gov.cn" style="font-size:12px" target="_blank"><b><%=blogabout%></b></a>
    </p>
   </div>
</div>
</body>
</html><%Sys.Close%>