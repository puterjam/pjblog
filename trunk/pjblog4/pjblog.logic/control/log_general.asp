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
		general.SiteName = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("SiteName"))))
		general.blog_Title = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_Title"))))
		general.blog_master = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_master"))))
		general.blog_email = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_email"))))
		general.blog_KeyWords = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_KeyWords"))))
		general.blog_Description = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_Description"))))
		If Right(Asp.CheckStr(Request.Form("SiteURL")), 1) <> "/" Then
			general.SiteURL = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("SiteURL")))) & "/"
		Else
			general.SiteURL = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("SiteURL"))))
		End If
		general.blog_about = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_about"))))
		general.blog_html = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_html"))))
		general.blog_postFile = Int(Asp.checkURL(Asp.CheckStr(Request.Form("blog_postFile"))))
		general.blog_SplitType = CBool(Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_SplitType")))))
		general.blog_introChar = Int(Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_introChar")))))
		general.blog_introLine = Int(Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_introLine")))))
		general.blog_PerPage = Int(Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blogPerPage")))))
		If Asp.IsInteger(Request.Form("blog_ImgLink")) And Int(Request.Form("blog_ImgLink")) = 1 Then
			general.blog_ImgLink = True
		Else
			general.blog_ImgLink = False
		End If
		general.blog_commPage = Int(Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blogcommpage")))))
		general.blog_commTimerout = Int(Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_commTimerout")))))
		general.blog_commLength = Int(Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_commLength")))))
		If Asp.IsInteger(Request.Form("blog_validate")) And Int(Request.Form("blog_validate")) = 1 Then
			general.blog_validate = True
		Else
			general.blog_validate = False
		End If
		If Asp.IsInteger(Request.Form("blog_commUBB")) And Int(Request.Form("blog_commUBB")) = 1 Then
			general.blog_commUBB = True
		Else
			general.blog_commUBB = False
		End If
		If Asp.IsInteger(Request.Form("blog_commIMG")) And Int(Request.Form("blog_commIMG")) = 1 Then
			general.blog_commIMG = True
		Else
			general.blog_commIMG = False
		End If
		general.blog_Disregister = Int(Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_Disregister")))))
		general.blog_CountNum = Int(Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_CountNum")))))
		general.blog_FilterName = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_FilterName"))))
		general.blog_FilterIP = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_FilterIP"))))
		ReConSio = general.SaveGeneral
		If ReConSio(0) Then Call Cache.GlobalCache(2)
		Session(Sys.CookieName & "_ShowMsg") = True
		Session(Sys.CookieName & "_MsgText") = "保存信息成功!"
		RedirectUrl(RedoUrl)
	End Sub
	
End Class
%>