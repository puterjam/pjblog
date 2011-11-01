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
        	rUrl = caload(Tid)
	    else
		 	rUrl = "article.asp?id=" & Tid
	    end if
		If rUrl = "" Then
		 	rUrl = "article.asp?id=" & Tid
		End If
	    RedirectUrl (rUrl)
	    Response.end
    End If
End If

If InStr(Replace(LCase(Request.ServerVariables("URL")), "\", "/"), "/article.asp") = 0 Then
    getBlogHead BlogTitle, "", -1, "", ""
End If


'输出文件头

Sub getBlogHead(Title, CateTitle, CategoryID, KeyWords, s_Description)
If not isblank(KeyWords) then
	blog_KeyWords = trim(KeyWords)&","
End If
If not isblank(s_Description) then
	blog_Description = trim(s_Description)&","
End If
'高亮分类for首页
If IsInteger(cateID) = True Then
    blog_currentCategoryID = cateID
End If

'高亮分类for日志单篇
If IsInteger(CategoryID) = True  and CategoryID<>-1 Then
    blog_currentCategoryID = CategoryID
End If

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="UTF-8">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
	<meta http-equiv="Content-Language" content="UTF-8" />
	<meta name="robots" content="all" />
	<meta name="author" content="<%=blog_email%>,<%=blog_master%>" />
	<meta name="Copyright" content="PJBlog3 CopyRight 2011" />
	<meta name="keywords" content="<%=blog_KeyWords%>" />
	<meta name="description" content="<%=blog_Description%> - <%=SiteName%> - <%=blog_Title%>" />
	<meta name="generator" content="PJBlog3" />
	<link rel="service.post" type="application/atom+xml" title="<%=SiteName%> - Atom" href="<%=siteURL%>xmlrpc.asp" />
	<link rel="EditURI" type="application/rsd+xml" title="RSD" href="<%=siteURL%>rsd.asp" />
	<title><%=Title%></title>

	<%if len(CateTitle)>0 and CategoryID>0 then %>
	<link rel="alternate" type="application/rss+xml" href="<%=siteURL%>feed.asp?cateID=<%=CategoryID%>" title="订阅 <%=siteName%> - <%=CateTitle%> 所有文章(rss2)" />
	<link rel="alternate" type="application/atom+xml" href="<%=siteURL%>atom.asp?cateID=<%=CategoryID%>"  title="订阅 <%=siteName%> - <%=CateTitle%> 所有文章(atom)"  />
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
	<script type="text/javascript" src="common/AjaxRequest.js"></script>
	<script type="text/javascript" src="common/Action.js"></script>
    <script language="javascript" type="text/javascript">
		var time = <%=blog_SaveTime%>;
	</script>
	<!--<script type="text/javascript" src="common/nicetitle.js"></script>-->
	<style type="text/css"> 
		.ownerClassLog{display:none}
		.ownerClassComment{display:<%if stat_Admin <> True then response.write("none")%>;}
		/*全局的提示框样式*/
		h5.tips{
			background:#FFCC00;border:1px solid #FFCC00;background-image:url(images/bg_tips.png);color:#990000;padding:3px;margin:0;font-size:13px
		}
		h5.tips img{
			border:0;background:transparent none !important;
		}
		.tips_body{
			background:#FFFFCC;border:1px solid #FFCC00;padding:4px;
		}
		.tips_body form{
			margin:0;
		}
		.tips_body a:link,.tips_body a:visited,.tips_body a:hover{
			color:#990000;
		}
		.tips_body .input{
			height:17px;
			border:1px solid #BD5B21
		}
		.tips_body .hints{
				margin:3px 0 0 3px;
				padding:2px 2px 2px 18px;
				background:url(images/notify.gif) no-repeat left 4px
		}
		.tips_body .error{
			border:1px solid #CC0033;
			padding:1px 3px 1px 21px;
			color:#990000;
			margin-bottom:2px;
			background:#FF9F88 url(images/tips.gif) no-repeat 3px 4px
		}
	</style>
</head>
<body onLoad="initJS()" onkeydown="PressKey()">
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
'合作信息
	SkinInfo = SkinInfo & " <a href=""http://www.codefense.cn"" target=""_blank""><img border=""0"" src=""images/detect.gif"" alt=""Code Detection By Codefense"" style=""margin-bottom:-2px;height:14px;width:12px""/></a>"
End If
End If
End Sub
%>
