<!--#include file="BlogCommon.asp" -->
﻿<!--#include file="common/function.asp" -->
<!--#include file="common/library.asp" -->
<!--#include file="common/cache.asp" -->
<!--#include file="common/checkUser.asp" -->
<!--#include file="common/ModSet.asp" -->
<!--#include file="class/cls_article.asp" -->
<%
	Dim id, tKey
	id = 0
	Response.CacheControl = "no-cache"
	Response.ContentType = "application/x-javascript"
	If CheckStr(Request.QueryString("id"))<>Empty Then
	    id = CheckStr(Request.QueryString("id"))
	End If

	'日志计数
	if IsInteger(id) = True then
		if id>0 then
			Dim log_View,log_ViewArr,c
			SQL = "SELECT top 1 log_ViewNums FROM blog_Content WHERE log_ID="&id&" and log_IsDraft=false"
			Set log_View = Server.CreateObject("ADODB.RecordSet")
			log_View.Open SQL, Conn, 1, 3
			c = log_View("log_ViewNums") + 1
			log_View("log_ViewNums") = c
			log_View.UPDATE
			log_ViewArr = log_View.GetRows
			log_View.Close
			Set log_View = Nothing
			Call updateViewNums(id, log_ViewArr(0, 0))
			outputPing c
		end if
	end if 
%>
<script runat="server" Language="javascript">
	function outputPing(num){
		Response.Clear();
		Response.Write ("﻿callPing('" + num + "');")
	}
</script>
<%
Session.CodePage = 936
Call CloseDB
%>
