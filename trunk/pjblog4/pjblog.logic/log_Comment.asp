<!--#include file = "../include.asp" -->
<!--#include file = "../pjblog.data/cls_comment.asp" -->
<!--#include file = "../pjblog.model/cls_Stream.asp" -->
<!--#include file = "../pjblog.data/cls_User.asp" -->
<!--#include file = "../pjblog.common/md5.asp" -->
<!--#include file = "../pjblog.common/sha1.asp" -->
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
			Case "del" Call del
			Case "Aduit" Call Aduit
		End Select
    End Sub 
     
	' ***********************************************
	'	类终结化
	' ***********************************************
    Private Sub Class_Terminate
		Sys.Close
    End Sub
	
	Private Sub Add
		Dim LastMSG, FlowControl, validate, password
		blog_ID = Asp.CheckUrl(Asp.CheckStr(Request.Form("blog_ID")))
		comm_Author = Trim(Asp.CheckUrl(Asp.CheckStr(Request.Form("comm_Author"))))
		comm_Content = Trim(Asp.CheckUrl(Asp.CheckStr(Request.Form("comm_Content"))))
		comm_IsAudit = True
		comm_Email = Trim(Asp.CheckUrl(Asp.CheckStr(Request.Form("comm_Email"))))
		comm_WebSite = Trim(Asp.CheckUrl(Asp.CheckStr(Request.Form("comm_WebSite"))))
		validate = Trim(Asp.CheckUrl(Asp.CheckStr(Request.Form("validate"))))
		password = Trim(Asp.CheckUrl(Asp.CheckStr(Request.Form("password"))))
		comm_PostTime = Now()
		FlowControl = False
		
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
		
		If stat_Admin = False Then
			If (memName = empty or blog_validate = true) And CStr(LCase(Session("GetCode"))) <> CStr(LCase(validate)) Then
				Response.Write("{Suc : false, Info : '验证码有误，请返回重新输入！'}")
				Exit Sub
			End If
			
			Set LastMSG = conn.Execute("select top 1 comm_Content from blog_Comment order by comm_ID desc")
				If LastMSG.EOF Then
					FlowControl = False
				Else
					If Trim(LastMSG("comm_Content")) = Trim(comm_Content) Then FlowControl = True
				End If
			Set LastMSG = Nothing
			
			If Left(comm_Content, 1)= Chr(32) then
				Response.Write("{Suc : false, Info : '评论内容中不允许首字空格'}")
				Exit Sub
			End If
			
			If Left(comm_Content,1)= Chr(13) then
				Response.Write("{Suc : false, Info : '评论内容中不允许首字换行'}")
				Exit Sub
			End If
			
			If FlowControl Then
				Response.Write("{Suc : false, Info : '禁止恶意灌水！'}")
				Exit Sub
			End If
			
			If DateDiff("s", Request.Cookies(Sys.CookieName)("memLastPost"), Asp.DateToStr(now(),"Y-m-d H:I:S")) < blog_commTimerout Then
				Response.Write("{Suc : false, Info : '发言太快,请 " & blog_commTimerout & " 秒后再发表评论'}")
				Exit Sub
			End If
		End If
	
		If Not Asp.IsInteger(blog_ID) Then
			Response.Write("{Suc : false, Info : '非法参数'}")
			Exit Sub
		End If
		
		If Len(comm_Author) < 1 Then
			Response.Write("{Suc : false, Info : '请输入你的昵称'}")
			Exit Sub
		End If
		
		If Asp.IsValidUserName(comm_Author) = False Then
			Response.Write("{Suc : false, Info : '<b>非法用户名！</b><br/>请尝试使用其他用户名！'}")
			Exit Sub
		End If
		
		If Not stat_CommentAdd Then
			Response.Write("{Suc : false, Info : '您没有权限发表评论'}")
			Exit Sub
		End If
		
		If Conn.Execute("select log_DisComment from blog_Content where log_ID=" & blog_ID)(0) Then
			Response.Write("{Suc : false, Info : '该日志不允许发表任何评论'}")
			Exit Sub
		End If
		
		If Len(comm_Content) < 1 Then
			Response.Write("{Suc : false, Info : '不允许发表空评论'}")
			Exit Sub
		End If
		
		If Len(comm_Content) > blog_commLength Then
			Response.Write("{Suc : false, Info : '评论超过最大字数限制'}")
			Exit Sub
		End If
		
		If Len(comm_Email) > 0 Then
			If Not Asp.RegMatch(comm_Email, "\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*") Then
				Response.Write("{Suc : false, Info : '验证邮箱格式错误'}")
				Exit Sub
			End If
		End If
		
		If Len(comm_WebSite) > 0 Then
			If Not Asp.RegMatch(comm_WebSite, "(http):\/\/[\w\-_]+(\.[\w\-_]+)+([\w\-\.,@?^=%&:/~\+#]*[\w\-\@?^=%&/~\+#])?") Then
				Response.Write("{Suc : false, Info : '网址格式错误,请用http://开头'}")
				Exit Sub
			End If
		End If
		
		' --------------------------------------------------------------
		' 	验证密码
		' --------------------------------------------------------------
		Dim CheckMem
		If Not stat_Admin Then
			If Len(password) > 0 Then
				CheckMem = User.UserLogin(comm_Author, password, CStr(LCase(validate)), 1)
				If Not CheckMem(3) Then
					Response.Write("{Suc : false, Info : '用户名或密码错误'}")
					Exit Sub
				End If
			End If
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
	
	Private Sub del
		Dim id
		id = Trim(Asp.CheckUrl(Asp.CheckStr(Request.QueryString("id"))))
		Comment.comm_ID = id
		OK = Comment.del
		If OK(0) Then
			Response.Write("{Suc : true, Info : '删除评论成功'}")
		Else
			Response.Write("{Suc : false, Info : '" & OK(1) & "'}")
		End If
	End Sub
	
	Private Sub Aduit
		Dim id, which
		id = Trim(Asp.CheckUrl(Asp.CheckStr(Request.QueryString("id"))))
		which = Trim(Asp.CheckUrl(Asp.CheckStr(Request.QueryString("which"))))
		Comment.comm_ID = id
		Comment.comm_IsAudit = CBool(which)
		OK = Comment.Aduit
		If OK(0) Then
			Response.Write("{Suc : true, Info : '修改审核成功'}")
		Else
			Response.Write("{Suc : false, Info : '" & OK(1) & "'}")
		End If
	End Sub
	
End Class
%>