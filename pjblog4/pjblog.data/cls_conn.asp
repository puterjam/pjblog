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
<script language="jscript" type="text/jscript" runat="server" src="../pjblog.common/language.js"></script>

<%
Dim Sys, Conn, memStatus, memName, RefUrl, RedoUrl, StartTime

Dim SiteName, SiteURL, blogPerPage, blog_LogNums, blog_CommNums, blog_MemNums, blog_VisitNums, blogBookPage, blog_MessageNums, blogcommpage, blogaffiche, blogabout, blogcolsize, blog_colNums, blog_TbCount, blog_showtotal, Register_UserNames, Register_UserName, FilterIPs, FilterIP, blog_commTimerout, blog_commUBB, blog_commIMG, blog_postFile, blog_postCalendar, blog_DefaultSkin, blog_SkinName, blog_SplitType, blog_introChar, blog_introLine, blog_validate, blog_Title, blog_ImgLink, blog_commLength, blog_downLocal, blog_DisMod, blog_Disregister, blog_master, blog_email, blog_CountNum, blog_wapNum, blog_wapImg, blog_wapHTML, blog_wapLogin, blog_wapComment, blog_wap, blog_wapURL, blog_KeyWords, blog_Description, blog_SaveTime, blog_IsPing, blog_html, blog_commAduit

Dim stat_title, stat_AddAll, stat_EditAll, stat_DelAll, stat_Add, stat_Edit, stat_Del, stat_CommentAdd
Dim stat_CommentDel, stat_Admin, stat_code, UP_FileType, UP_FileSize, UP_FileTypes, stat_FileUpLoad
Dim stat_CommentEdit, stat_ShowHiddenCate

RedoUrl = Cstr(Request.ServerVariables("HTTP_REFERER"))
StartTime = Timer()

memStatus = "Guest"

Set Sys = New Sys_Connection
	Sys.version = "4.1.2.400"   					' 设置当前版本号
	Sys.UpdateTime = "2009-11-03"					' 设置当前版本更新日期
	Sys.CookieName = "PJBlog4"						' 设置你的Cookie名称
	Sys.AccessFile = "pjblog/PBLog4.asp"			' 设置你的数据库名称
	Set Conn = Sys.open								' 打开数据库
%>