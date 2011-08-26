<!--#include file="conCommon.asp" -->
<!--#include file="common/function.asp" -->
<!--#include file="common/library.asp" -->
<!--#include file="common/cache.asp" -->
<!--#include file="common/checkUser.asp" -->
<!--#include file="common/sha1.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="UTF-8">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="Content-Language" content="UTF-8" />
<meta name="author" content="puter2001@21cn.com,PuterJam" />
<meta name="Copyright" content="PL-Blog 2 CopyRight 2005" />
<meta name="keywords" content="PuterJam,Blog,ASP,designing,with,web,standards,xhtml,css,graphic,design,layout,usability,ccessibility,w3c,w3,w3cn" />
<meta name="description" content="PuterJam's BLOG" />

<title>后台管理</title>
</head>
<%
If memName = Empty Or stat_Admin<>True Then
    RedirectUrl("default.asp")
Else
    If session(CookieName&"_System") = True Then

%>
    <frameset rows="31,*" framespacing="0" border="0" frameborder="0">
    <frame src="ConHead.asp" scrolling="no" name="Head" noresize>
      <frameset  id="ContentSet" rows="80,*">
      <frame src="ConMenu.asp" name="Menu">
       <frame src="ConContent.asp" name="MainContent" scrolling="yes" noresize>
     </frameset>
    </frameset><noframes></noframes>
    <%
Else

%>
  <body style="background:#E8EDF4">
  <form action="control.asp" method="post" style="margin:2px;">
  <input type="hidden" name="action" value="login"/>
  <div style="font-size:12px;text-align:center">
   <div style="margin:auto;border:1px solid #999;padding:2px;background:#fff;width:350px;">

    <div style="border:1px solid #e5e5e5;padding:2px;">
	    <div style="border-bottom:1px solid #e5e5e5;padding:5px;text-align:left"><img src="images/Control/logo.gif"/></div>
	    <div style="padding:16px;padding-top:40px"><b style="margin-left:-146px;font-size:14px;">管理员密码: </b><br/><br/><input name="adpass" type="password" size="20" style="border:1px solid #999;font-size:18px"/></div>
	    <input type="submit" value=" 登 录 " style="background:#fff;border:1px solid #999;padding:2px 2px 0px 2px;margin:4px;border-width:1px 3px 1px 3px"/>
	    <div style="padding:8px;height:22px;color:#f00;font-weight:bold"><%=session(CookieName&"_ShowError")%></div><%session(CookieName&"_ShowError")=""%>
	    <div style="padding:2px;font-family:arial;color:#666;font-size:9px;text-align:right"><b>PJBlog3 v<%=blog_version%></b></div>
   </div>
   </div>
  </div>
  </form></body>
  <%
Dim action
action = CheckStr(Request.Form("action"))
If action = "login" Then
    Dim chUser, getPass
    getPass = CheckStr(Request.Form("adpass"))
    Set chUser = conn.Execute("SELECT Top 1 mem_Name,mem_Password,mem_salt,mem_Status,mem_LastIP,mem_lastVisit,mem_hashKey FROM blog_member WHERE mem_Name='"&memName&"'")
    If chUser.EOF Or chUser.bof Then
        session(CookieName&"_ShowError") = "管理员密码错误!"
        RedirectUrl("control.asp")
    Else
        If chUser("mem_Password")<>SHA1(getPass&chUser("mem_salt")) Then
            session(CookieName&"_ShowError") = "管理员密码错误!"
            RedirectUrl("control.asp")
        Else
            session(CookieName&"_System") = True
            RedirectUrl("control.asp")
        End If
    End If
End If
End If
End If
%>
</html>
