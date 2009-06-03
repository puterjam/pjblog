<!--#include file="BlogCommon.asp" -->
﻿<!--#include file="common/function.asp" -->
<!--#include file="common/library.asp" -->
<!--#include file="common/cache.asp" -->
<!--#include file="common/checkUser.asp" -->
<!--#include file="common/ubbcode.asp" -->
<!--#include file="class/cls_article.asp" -->

<%
	Dim id, comDesc
	id = 0
	If Request.QueryString("comDesc") =  "Desc" Then comDesc = "Desc" Else comDesc = "Asc" End If	
	
	If CheckStr(Request.QueryString("id"))<>Empty Then
	    id = CheckStr(Request.QueryString("id"))
	End If
	
	'Response.Expires = 100
	Response.AddHeader "pragma", "no-cache"
	Response.AddHeader "cache-ctrol", "no-cache"
	Response.ContentType = "application/x-javascript" 
	
	if IsInteger(id) = True then
		if id>0 then
			outputComment ShowComm(id, comDesc, False, True, True, False, True, True)
		end if
	end if 
%>
<script runat="server" Language="javascript">
	function outputComment(html){
		html = html.replace(/[\n\r]/g,"");
		html = html.replace(/\\/g,"\\\\");
		html = html.replace(/\'/g,"\\'");
		Response.Write ("fillComment('" + html + "');")
	}
</script>
<%
Session.CodePage = 936
Call CloseDB
%>
