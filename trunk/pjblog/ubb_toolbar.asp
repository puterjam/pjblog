<!--#include file="const.asp" -->
﻿<!--#include file="common/function.asp" -->
<!--#include file="common/cache.asp" -->
<!--#include file="common/UBBconfig.asp" -->
<%
'====
'自动展开编辑器的优化
' 2008-6-29
'====
getInfo 0 

Skins = blog_DefaultSkin

Response.Expires = 99999 '让服务器缓存这个文件更加长时间
Response.ContentType = "application/x-javascript"

UBB_Tools_Items="bold,italic,underline,deleteline"
UBB_Tools_Items=UBB_Tools_Items&"||image,link,mail,quote,smiley"
             	   
dim toolbarCode, smileCode
toolbarCode = replace(ToolsToCode,"'","\'")
smileCode = replace(showSmilie,"'","\'")
Response.Clear
Response.write "var ubbTools='"&toolbarCode&"';"
Response.write "var ubbSmile='"&smileCode&"';"
Response.write "showUBB();"
%>