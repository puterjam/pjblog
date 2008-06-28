<!--#include file="BlogCommon.asp" -->
﻿<!--#include file="common/function.asp" -->
<!--#include file="common/library.asp" -->
<!--#include file="common/cache.asp" -->
<!--#include file="common/checkUser.asp" -->
<!--#include file="common/ModSet.asp" -->
<!--#include file="plugins.asp" -->
<%
	Response.AddHeader "pragma", "no-cache"
	Response.AddHeader "cache-ctrol", "no-cache"
	Response.ContentType = "application/x-javascript" 
	
	Side_Module_Replace '处理系统侧栏模块信息
	outputSideBar side_html_static
%>
<script runat="server" Language="javascript">
	function outputSideBar(html){
		html = html.replace(/[\n\r]/g,"");
		html = html.replace(/\'/g,"\\'");
		Response.Write ("fillSideBar('" + html + "');")
	}
</script>
<%
Session.CodePage = 936
Call CloseDB
%>
