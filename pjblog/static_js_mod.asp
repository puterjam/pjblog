<!--#include file="BlogCommon.asp" -->
﻿<!--#include file="common/function.asp" -->
<!--#include file="common/library.asp" -->
<!--#include file="common/cache.asp" -->
<!--#include file="common/checkUser.asp" -->
<!--#include file="common/ModSet.asp" -->
<!--#include file="class/cls_article.asp" -->
<!--#include file="Plugins.asp" -->
<%

	'Response.AddHeader "ETag", "test"
	'Response.CacheControl = "no-cache"
	'Response.Expires = 10
	'Response.AddHeader "Last-Modified", getNow()
	
	'Response.Status = "304"
	'Response.Clear
	'log_module(2)
	
	Side_Module_Replace '处理系统侧栏模块信息
	outputSideBar side_html_static
%>
<script runat="server" Language="javascript">
	function outputSideBar(html){
		html = html.replace(/[\n\r]/g,"");
		html = html.replace(/\\/g,"\\\\");
		html = html.replace(/\'/g,"\\'");
		Response.Clear();
		Response.Write ("﻿document.write('" + html + "');")
	}
</script>
<%
Session.CodePage = 936
Call CloseDB
%>
