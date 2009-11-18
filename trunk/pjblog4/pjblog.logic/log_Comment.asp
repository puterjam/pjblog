<!--#include file = "../include.asp" -->
<!--#include file = "../pjblog.data/cls_comment.asp" -->
<!--#include file = "../pjblog.model/cls_Stream.asp" -->
<%
Dim Comm
Set Comm = New do_Comment
Set Comm = Nothing

Class do_Comment

	Private comm_ID, blog_ID, comm_Content, comm_Author, comm_PostTime, comm_DisSM, comm_DisUBB, comm_DisIMG, comm_AutoURL, comm_PostIP, comm_AutoKEY, comm_IsAudit, comm_Email, comm_WebSite, reply_author, reply_time, replay_content
	Private OK, doo, Str, local
	Private DisSM, DisUBB, DisIMG, AutoURL, AutoKEY
	
	Public Property Get Action
		Action = Request.QueryString("action")
	End Property
	
	' ***********************************************
	'	类初始化
	' ***********************************************
	Private Sub Class_Initialize
		local = "../pjblog.template/" & blog_DefaultSkin & "/"
		Select Case Action
			Case "add" Call Add
		End Select
    End Sub 
     
	' ***********************************************
	'	类终结化
	' ***********************************************
    Private Sub Class_Terminate
		Sys.Close
    End Sub
	
	Private Sub Add
		blog_ID = Asp.CheckUrl(Asp.CheckStr(Request.Form("blog_ID")))
		comm_Author = Trim(Asp.CheckUrl(Asp.CheckStr(Request.Form("comm_Author"))))
		comm_Content = Trim(Asp.CheckUrl(Asp.CheckStr(Request.Form("comm_Content"))))
		comm_IsAudit = True
		comm_Email = Trim(Asp.CheckUrl(Asp.CheckStr(Request.Form("comm_Email"))))
		comm_WebSite = Trim(Asp.CheckUrl(Asp.CheckStr(Request.Form("comm_WebSite"))))
		comm_PostTime = Now()
		
		If Not Asp.IsInteger(blog_ID) Then
			Response.Write("{Suc : false, Info : '非法参数'}")
			Exit Sub
		End If
		
		If Asp.IsInteger(Request.Form("comm_DisSM")) And Int(Request.Form("comm_DisSM")) = 1 Then
			DisSM = 1
		Else
			DisSM = 2
		End If
		
		If Asp.IsInteger(Request.Form("comm_AutoURL")) And Int(Request.Form("comm_AutoURL")) = 1 Then
			AutoURL = 1
		Else
			AutoURL = 2
		End If
		
		If Asp.IsInteger(Request.Form("comm_AutoKEY")) And Int(Request.Form("comm_AutoKEY")) = 1 Then
			AutoKEY = 1
		Else
			AutoKEY = 2
		End If

		comm_DisSM = DisSM
		comm_AutoURL = AutoURL
		comm_AutoKEY = AutoKEY
		
		If Left(comm_Content, 1)= Chr(32) then
			Response.Write("{Suc : false, Info : '评论内容中不允许首字空格'}")
			Exit Sub
		End If
		
		If Left(comm_Content,1)= Chr(13) then
			Response.Write("{Suc : false, Info : '评论内容中不允许首字换行'}")
			Exit Sub
		End If
		
		'If DateDiff("s", Request.Cookies(Sys.CookieName)("memLastPost"), Asp.DateToStr(now(),"Y-m-d H:I:S")) < blog_commTimerout Then
			'Response.Write("{Suc : false, Info : '发言太快,请 " & blog_commTimerout & " 秒后再发表评论'}")
			'Exit Sub
		'End If
		
		If Len(comm_Author) < 1 Then
			Response.Write("{Suc : false, Info : '请输入你的昵称'}")
			Exit Sub
		End If
		
		If Asp.IsValidUserName(comm_Author) = False Then
			Response.Write("{Suc : false, Info : '<b>非法用户名！</b><br/>请尝试使用其他用户名！'}")
			Exit Sub
		End If
		
		If Int(blog_commUBB) <> 0 Then blog_commUBB = 1
		If Int(blog_commIMG) <> 0 Then blog_commIMG = 1
		
		Comment.blog_ID = blog_ID
		Comment.comm_Author = comm_Author
		Comment.comm_Content = comm_Content
		Comment.comm_DisSM = comm_DisSM
		Comment.comm_DisUBB = blog_commUBB
		Comment.comm_DisIMG = blog_commIMG
		Comment.comm_AutoURL = comm_AutoURL
		Comment.comm_AutoKEY = comm_AutoKEY
		Comment.comm_IsAudit = comm_IsAudit
		Comment.comm_Email = comm_Email
		Comment.comm_WebSite = comm_WebSite
		Comment.comm_PostTime = comm_PostTime

		OK = Comment.Add
		If OK(0) Then
			Response.Cookies(Sys.CookieName)("memLastpost") = Asp.DateToStr(comm_PostTime, "Y-m-d H:I:S")
			Set doo = New Stream
				Str = doo.LoadFile(local & "l_comment.html")
			Set doo = Nothing
			Str = Replace(Str, "<#comm_id#>", Int(OK(1)), 1, -1, 1)
			Str = Replace(Str, "<#comm_author#>", memName, 1, -1, 1)
			Str = Replace(Str, "<#comm_Content#>", UBBCode(comm_Content, DisSM, blog_commUBB, blog_commIMG, AutoURL, AutoKEY, True), 1, -1, 1)
			Str = Replace(Str, "<#comm_mail#>", comm_Email, 1, -1, 1)
			Str = Replace(Str, "<#comm_weburl#>", comm_WebSite, 1, -1, 1)
			Str = Replace(Str, "<#comm_ip#>", Asp.getIP, 1, -1, 1)
			Str = Replace(Str, "<#comm_posttime#>", Asp.DateToStr(comm_PostTime, "Y-m-d H:I:S"), 1, -1, 1)
			Response.Write("{Suc : true, Info : '" & cee.encode(Str) & "', id : " & OK(1) & "}")
		Else
			Response.Write("{Suc : false, Info : '" & OK(1) & "', id : 0}")
		End If
	End Sub
	
End Class
%>