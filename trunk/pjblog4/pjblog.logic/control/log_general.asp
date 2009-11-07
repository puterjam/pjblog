<!--#include file = "../../include.asp" -->
<!--#include file = "../../pjblog.data/control/cls_general.asp" -->
<%
'*************** PJblog4 基本设置 *******************
' PJblog4 Copyright 2009
' Update: 2009-11-03
' Author: evio
' Owner : PuterJam
' Mode  : Class
' MoName: User Model
'***************************************************
Dim dogeneral
Set dogeneral = New do_general
Set dogeneral = Nothing

Class do_general
	
	Public Property Get Action
		Action = Request.QueryString("action")
	End Property
	
	Private ReConSio

	' ***********************************************
	'	用户操作方法类初始化
	' ***********************************************
	Private Sub Class_Initialize
		Select Case Action
			Case "SaveGeneral" Call SaveGeneral
		End Select
    End Sub 
     
	' ***********************************************
	'	用户操作方法类终结化
	' ***********************************************
    Private Sub Class_Terminate
		Sys.Close
    End Sub
	
	Private Sub SaveGeneral
		general.SiteName = Asp.checkURL(Asp.CheckStr(Request.Form("SiteName")))
		general.blog_Title = Asp.checkURL(Asp.CheckStr(Request.Form("blog_Title")))
		general.blog_master = Asp.checkURL(Asp.CheckStr(Request.Form("blog_master")))
		general.blog_email = Asp.checkURL(Asp.CheckStr(Request.Form("blog_email")))
		general.blog_KeyWords = Asp.checkURL(Asp.CheckStr(Request.Form("blog_KeyWords")))
		general.blog_Description = Asp.checkURL(Asp.CheckStr(Request.Form("blog_Description")))
		If Right(Asp.CheckStr(Request.Form("SiteURL")), 1) <> "/" Then
			general.SiteURL = Asp.checkURL(Asp.CheckStr(Request.Form("SiteURL"))) & "/"
		Else
			general.SiteURL = Asp.checkURL(Asp.CheckStr(Request.Form("SiteURL")))
		End If
		general.blog_about = Asp.checkURL(Asp.CheckStr(Request.Form("blog_about")))
		ReConSio = general.SaveGeneral
		If ReConSio(0) Then Call Cache.GlobalCache(2)
		Session(Sys.CookieName & "_ShowMsg") = True
		Session(Sys.CookieName & "_MsgText") = ReConSio(1)
		RedirectUrl(RedoUrl)
	End Sub
	
End Class
%>