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
BlogTitle = siteName & "-" & blog_Title
If InStr(Replace(LCase(Request.ServerVariables("URL")), "\", "/"), "/default.asp")<>0 Then

'备用做304优化
'	Dim clientEtag, serverEtag
'	serverEtag = getEtag
'	clientEtag = Request.ServerVariables("HTTP_IF_NONE_MATCH")
'	Response.AddHeader "ETag", getEtag
	
'	if serverEtag = clientEtag then
'		Response.Status = "304 Not Modified"
'		Session.CodePage = 936
'		Call CloseDB
'		Response.end
'	end if
	
    Dim Tid
    If CheckStr(Request.QueryString("id"))<>Empty Then
        Tid = CheckStr(Request.QueryString("id"))
    End If
    If Len(Tid)>0 Then 
    	Dim rUrl
        If blog_postFile = 2 Then
        	rUrl = "article/" & Tid & ".htm"
	    else
		 	rUrl = "article.asp?id=" & Tid
	    end if 
	    RedirectUrl (rUrl)
	    Response.end
    End If
End If

If InStr(Replace(LCase(Request.ServerVariables("URL")), "\", "/"), "/article.asp") = 0 Then
    getBlogHead BlogTitle, "", -1
End If

'输出文件头

Sub getBlogHead(Title, CateTitle, CateID)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="UTF-8">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
	<meta http-equiv="Content-Language" content="UTF-8" />
	<meta name="robots" content="all" />
	<meta name="author" content="<%=blog_email%>,<%=blog_master%>" />
	<meta name="Copyright" content="PJBlog3 CopyRight 2008" />
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
   		<%=CategoryList(0)%>
  </div>
<%
End Sub

'读取Flash导航条
Dim SkinInfo

Sub getSkinFlash
    If CheckObjInstalled(getXMLDOM()) Then
        Dim SkinXML
        Set SkinXML = New PXML
        SkinInfo = ""
        SkinXML.XmlPath = "skins/"&Skins&"/skin.xml"
        SkinXML.Open
        If SkinXML.getError = 0 Then
            SkinInfo = " , <a href=""" & SkinXML.SelectXmlNodeText("DesignerURL") & """ target=""_blank""><strong>" & SkinXML.SelectXmlNodeText("SkinName") & "</strong></a> Design By <a href=""mailto:" & SkinXML.SelectXmlNodeText("DesignerMail") & """ target=""_blank""><strong>" & SkinXML.SelectXmlNodeText("SkinDesigner") & "</strong></a>"
            Dim useFlash
            useFlash = SkinXML.SelectXmlNodeText("Flash/UseFlash")
            If useFlash = "" Then useFlash = "false"
            If CBool(useFlash) Then

%>
			       <div id="FlashHead" style="text-align:<%=SkinXML.SelectXmlNodeText("Flash/FlashAlign")%>;top:<%=SkinXML.SelectXmlNodeText("Flash/FlashTop")%>px"></div>
			       <script type="text/javascript">WriteHeadFlash('<%="skins/"&Skins&"/"&SkinXML.SelectXmlNodeText("Flash/FlashPath")%>','<%=SkinXML.SelectXmlNodeText("Flash/FlashWidth")%>','<%=SkinXML.SelectXmlNodeText("Flash/FlashHeight")%>',<%=Lcase(CBool(SkinXML.SelectXmlNodeText("Flash/FlashTransparent")))%>)</script>
			      <%
End If
SkinXML.CloseXml
Set SkinXML = Nothing
End If
End If
End Sub
%>
