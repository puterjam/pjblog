<!--#include file="common/function.asp" -->
<!--#include file="common/ubbcode.asp" -->
<!--#include file="common/library.asp" -->
<!--#include file="common/cache.asp" -->
<!--#include file="common/checkUser.asp" -->
<!--#include file="common/XML.asp" -->
<%
'==================================
'  Blog顶部
'    更新时间: 2005-10-23
'==================================

'=========================Funciton In Head=============================
'处理标题
 Dim BlogTitle
 BlogTitle=siteName & "-" & blog_Title 
 if inStr(Replace(Lcase(Request.ServerVariables("URL")),"\","/"),"/default.asp")<>0 then
 	dim Tid
	If CheckStr(Request.QueryString("id"))<>Empty Then
		Tid=CheckStr(Request.QueryString("id"))
	End If
	if len(Tid)>0 then Response.Redirect ("article.asp?id="&Tid)
 end if
 
 if inStr(Replace(Lcase(Request.ServerVariables("URL")),"\","/"),"/article.asp")=0 then
  	getBlogHead BlogTitle,"",-1
 end if

'输出文件头
sub getBlogHead(Title,CateTitle,CateID)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="UTF-8">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
	<meta http-equiv="Content-Language" content="UTF-8" />
	<meta name="robots" content="all" />
	<meta name="author" content="<%=blog_email%>,<%=blog_master%>" />
	<meta name="Copyright" content="PJBlog2 CopyRight 2008" />
	<meta name="keywords" content="PuterJam,Blog,ASP,designing,with,web,standards,xhtml,css,graphic,design,layout,usability,accessibility,w3c,w3,w3cn" />
	<meta name="description" content="<%=SiteName%> - <%=blog_Title%>" />
	<title><%=Title%></title>
	<%if len(CateTitle)>0 and CateID>0 then %>
	<link rel="alternate" type="application/rss+xml" href="<%=siteURL%>feed.asp?cateID=<%=CateID%>" title="订阅 <%=siteName%> - <%=CateTitle%> 所有文章(rss2)" />
	<link rel="alternate" type="application/atom+xml" href="<%=siteURL%>atom.asp?cateID=<%=CateID%>"  title="订阅 <%=siteName%> - <%=CateTitle%> 所有文章(atom)"  />
	<%else%>
	<link rel="alternate" type="application/rss+xml" href="<%=siteURL%>feed.asp" title="订阅 <%=siteName%> 所有文章(rss2)" />
	<link rel="alternate" type="application/atom+xml" href="<%=siteURL%>atom.asp" title="订阅 <%=siteName%> 所有文章(atom)" />
	<%end if%>
	<link rel="stylesheet" rev="stylesheet" href="skins/<%=Skins%>/global.css" type="text/css" media="all" /><!--全局样式表-->
	<link rel="stylesheet" rev="stylesheet" href="skins/<%=Skins%>/layout.css" type="text/css" media="all" /><!--层次样式表-->
	<link rel="stylesheet" rev="stylesheet" href="skins/<%=Skins%>/typography.css" type="text/css" media="all" /><!--局部样式表-->
	<link rel="stylesheet" rev="stylesheet" href="skins/<%=Skins%>/link.css" type="text/css" media="all" /><!--超链接样式表-->
	<link rel="stylesheet" rev="stylesheet" href="skins/<%=Skins%>/UBB/editor.css" type="text/css" media="all" /><!--UBB编辑器代码-->
	<link rel="stylesheet" rev="stylesheet" href="FCKeditor/editor/css/Dphighlighter.css" type="text/css" media="all" /><!--FCK块引用&代码样式-->
	<link rel="icon" href="favicon.ico" type="image/x-icon" />
	<link rel="shortcut icon" href="favicon.ico" type="image/x-icon" />
	<script type="text/javascript" src="common/common.js"></script>
	<!--<script type="text/javascript" src="common/nicetitle.js"></script>-->
</head>
<body onload="initJS()" onkeydown="PressKey()">
<a href="default.asp" accesskey="i"></a>
<a href="javascript:history.go(-1)" accesskey="z"></a>
<%getSkinFlash%>
 <div id="container">
 <!--顶部-->
  <div id="header">
   <div id="blogname"><%=siteName%>
    <div id="blogTitle"><%=blog_Title%></div>
   </div>
   		<%CategoryList(0)%>
  </div>
<%
end sub

'读取Flash导航条
dim SkinInfo
sub getSkinFlash
 if CheckObjInstalled(getXMLDOM()) then
	  dim SkinXML
	  set SkinXML=new PXML
	  SkinInfo=""
	  SkinXML.XmlPath="skins/"&Skins&"/skin.xml"
	  SkinXML.open
	  if SkinXML.getError=0 then
		    SkinInfo=" , <a href=""" & SkinXML.SelectXmlNodeText("DesignerURL") & """ target=""_blank""><strong>" & SkinXML.SelectXmlNodeText("SkinName") & "</strong></a> Design By <a href=""mailto:" & SkinXML.SelectXmlNodeText("DesignerMail") & """ target=""_blank""><strong>" & SkinXML.SelectXmlNodeText("SkinDesigner") & "</strong></a>"
		     dim useFlash
		     useFlash = SkinXML.SelectXmlNodeText("Flash/UseFlash")
		     if useFlash = "" then useFlash = "false"
		     if CBool(useFlash) then 
			      %>
			       <div id="FlashHead" style="text-align:<%=SkinXML.SelectXmlNodeText("Flash/FlashAlign")%>;top:<%=SkinXML.SelectXmlNodeText("Flash/FlashTop")%>px"></div> 
			       <script type="text/javascript">WriteHeadFlash('<%="skins/"&Skins&"/"&SkinXML.SelectXmlNodeText("Flash/FlashPath")%>','<%=SkinXML.SelectXmlNodeText("Flash/FlashWidth")%>','<%=SkinXML.SelectXmlNodeText("Flash/FlashHeight")%>',<%=Lcase(CBool(SkinXML.SelectXmlNodeText("Flash/FlashTransparent")))%>)</script>
			      <%
		       end if
		      SkinXML.CloseXml
		      set SkinXML=nothing
	    end if
 end if
end sub
%>