<!--#include file = "../include.asp" -->
<!--#include file = "../pjblog.data/cls_User.asp" -->
<!--#include file = "../pjblog.common/md5.asp" -->
<!--#include file = "../pjblog.common/sha1.asp" -->
<%
'*************** PJblog4 基本设置 *******************
' PJblog4 Copyright 2009
' Update: 2009-11-03
' Author: evio
' Owner : PuterJam
' Mode  : Class
' MoName: User Model
'***************************************************
Dim doUser
Set doUser = New do_User
Set doUser = Nothing

Class do_User
	
	Private MsgInfo, outStr
	
	Public Property Get Action
		Action = Request.QueryString("action")
	End Property
	
	' ***********************************************
	'	用户操作方法类初始化
	' ***********************************************
	Private Sub Class_Initialize
		Select Case Action
			Case "login" Call UserLogin
			Case Else outStr = Asp.MessageInfo(Array(lang.Set.Asp(6), lang.Set.Asp(0), "ErrorIcon")) : Response.Write(outStr)
		End Select
    End Sub 
     
	' ***********************************************
	'	用户操作方法类终结化
	' ***********************************************
    Private Sub Class_Terminate
		Sys.Close
    End Sub
	
	' ***********************************************
	'	用户登入
	' ***********************************************
	Private Sub UserLogin
		Dim UserName, PassWord, validate, KeepLogin
		RefUrl = Session(Sys.CookieName & "_Login_Referer_Url")
		If len(RefUrl) < 8 Then RefUrl = Cstr(Request.ServerVariables("HTTP_REFERER"))
    	If len(RefUrl) < 8 Then RefUrl = "http://" & Request.ServerVariables("HTTP_HOST")
		
		UserName = Request.Form("UserName")
		PassWord = Request.Form("Password")
		validate = Request.Form("validate")
		KeepLogin = Request.Form("KeepLogin")
		
		MsgInfo = User.UserLogin(UserName, PassWord, validate, KeepLogin)
		MsgInfo(1) = Replace(Replace(MsgInfo(1), "default.asp", RefUrl), lang.Set.Asp(15) & "</a>", lang.Set.Asp(16) & "</a>&nbsp;|&nbsp;<a href=""default.asp"">" & lang.Set.Asp(13) & "</a><br/>")
		outStr = Asp.MessageInfo(MsgInfo)
		Response.Write(outStr)
	End Sub
	
End Class
%>