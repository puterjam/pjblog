<!--#include file="const.asp" -->
﻿<!--#include file="common/function.asp" -->
<!--#include file="common/cache.asp" -->

<%
	Dim id, tKey, clientEtag, serverEtag
	id = 0
	serverEtag = getEtag
	clientEtag = Request.ServerVariables("HTTP_IF_NONE_MATCH")
	

	Response.AddHeader "ETag", getEtag
	'Response.CacheControl = "no-cache"
	'Response.Expires = 10
	'Response.AddHeader "Last-Modified", getNow()
	Response.ContentType = "application/x-javascript" 

	if serverEtag = clientEtag then
		Response.Status = "304 Not Modified"
	else
		Server.Transfer "static_js_mod.asp"
		'Server.Transfer("static_js_mod.asp");
		'Side_Module_Replace '处理系统侧栏模块信息
		'outputSideBar side_html_static
	end if
%>