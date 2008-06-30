<!--#include file="BlogCommon.asp" -->
﻿<!--#include file="common/function.asp" -->
<!--#include file="common/library.asp" -->
<!--#include file="common/cache.asp" -->
<!--#include file="common/checkUser.asp" -->
<!--#include file="common/ModSet.asp" -->
<!--#include file="plugins.asp" -->
<!--#include file="class/cls_article.asp" -->
<%
	Dim id, tKey
	id = 0
	
	If CheckStr(Request.QueryString("id"))<>Empty Then
	    id = CheckStr(Request.QueryString("id"))
	End If
	
	'Response.AddHeader "ETag", "test"
	Response.AddHeader "Cache-Control", "no-cache"
	'Response.Expires = 10
	'Response.AddHeader "Last-Modified", getNow()
	Response.ContentType = "application/x-javascript" 

	
	
	Side_Module_Replace '处理系统侧栏模块信息
	outputSideBar side_html_static
	
	'日志计数
	if IsInteger(id) = True then
		if id>0 then
			Dim log_View,log_ViewArr
			SQL = "SELECT top 1 log_ViewNums FROM blog_Content WHERE log_ID="&id&" and log_IsDraft=false"
			Set log_View = Server.CreateObject("ADODB.RecordSet")
			log_View.Open SQL, Conn, 1, 3
			log_View("log_ViewNums") = log_View("log_ViewNums") + 1
			log_View.UPDATE
			log_ViewArr = log_View.GetRows
			log_View.Close
			Set log_View = Nothing
			Call updateViewNums(id, log_ViewArr(0, 0))
		end if
	end if 
%>
<script runat="server" Language="javascript">
	function outputSideBar(html){
		html = html.replace(/[\n\r]/g,"");
		html = html.replace(/\\/g,"\\\\");
		html = html.replace(/\'/g,"\\'");
		Response.Write ("fillSideBar('" + html + "');")
	}
	function getNow(){
		var d = new Date();
		return d.toGMTString();
	}
</script>
<%
Session.CodePage = 936
Call CloseDB
%>
