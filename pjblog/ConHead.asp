<!--#include file="conCommon.asp" -->
<!--#include file="common/function.asp" -->
<!--#include file="common/library.asp" -->
<!--#include file="common/cache.asp" -->
<!--#include file="common/checkUser.asp" -->
<%
'==================================
'  后台管理顶部
'    更新时间: 2005-6-24
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
<title>后台管理-顶部</title>
</head>
<body class="headbody">
 <div class="headmain">
   <div style="height:70px;background:url('images/Control/Pic2.jpg') no-repeat;">
   <div style="padding-top:53px;padding-left:70px;font-size:11px;font-family:verdana;font-weight:bold;color:#fff;">PJBlog2 v<%=blog_version%> - <%=DateToStr(blog_UpdateDate,"mdy")%></div>
   </div>
 </div>
</body>
</html>
<%
Else
    RedirectUrl("default.asp")
End If
%>
