<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'*************** PJblog4 基本设置 *******************
' PJblog4 Copyright 2009
' Update: 2009-11-03
' Author: evio
' Owner : PuterJam
' Mode  : Normal
' MoName: Access Connect
'***************************************************

Option Explicit

Response.Buffer = True
Response.Charset = "UTF-8"
Server.ScriptTimeOut = 90
Session.CodePage = 65001
Session.LCID = 2057

%>

<!--#include file = "cls_const.asp" -->

<%
Dim Sys, Conn

Set Sys = New Sys_Connection
	Sys.version = "4.1.2.400"   					' 设置当前版本号
	Sys.UpdateTime = "2009-11-03"					' 设置当前版本更新日期
	Sys.CookieName = "PJBlog4"						' 设置你的Cookie名称
	Sys.AccessFile = "pjblog/PBLog4.asp"			' 设置你的数据库名称
	Set Conn = Sys.open
%>