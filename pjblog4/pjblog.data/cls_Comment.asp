<%
Dim Comment
Set Comment = New Sys_Comment

Class Sys_Comment

	Public comm_ID, blog_ID, comm_Content, comm_Author, comm_PostTime, comm_DisSM, comm_DisUBB, comm_DisIMG, comm_AutoURL, comm_PostIP, comm_AutoKEY, comm_IsAudit, comm_Email, comm_WebSite, reply_author, reply_time, replay_content
	
	Private Arrays
	
	' ***********************************************
	'	类初始化
	' ***********************************************
	Private Sub Class_Initialize
		
    End Sub 
     
	' ***********************************************
	'	类终结化
	' ***********************************************
    Private Sub Class_Terminate
    End Sub
	
	Public Function Add
		If Not Asp.IsInteger(blog_ID) Then
			Add = Array(False, "非法参数")
			Exit Function
		End If
		If Not stat_CommentAdd Then
			Add = Array(False, "您没有权限发表评论")
			Exit Function
		End If
		Arrays = Array(Array("blog_ID", blog_ID), Array("comm_Content", comm_Content), Array("comm_Author", comm_Author), Array("comm_PostTime", comm_PostTime), Array("comm_DisSM", comm_DisSM), Array("comm_DisUBB", comm_DisUBB), Array("comm_DisIMG", comm_DisIMG), Array("comm_AutoURL", comm_AutoURL), Array("comm_PostIP", Asp.getIP), Array("comm_AutoKEY", comm_AutoKEY), Array("comm_IsAudit", comm_IsAudit), Array("comm_Email", comm_Email), Array("comm_WebSite", comm_WebSite))
		Add = Sys.doRecord("blog_Comment", Arrays, "insert", "comm_ID", "")
		Application.Lock()
		Application(Sys.CookieName & "_Comment_Edit") = True
		Application.UnLock()
	End Function
	
	Public Function edit
		If Not Asp.IsInteger(comm_ID) Then
			edit = Array(False, "非法参数")
			Exit Function
		End If
		If stat_CommentEdit And (memName = comm_Author) Then
			Arrays = Array(Array("comm_Content", comm_Content), Array("comm_DisSM", comm_DisSM), Array("comm_DisUBB", comm_DisUBB), Array("comm_DisIMG", comm_DisIMG), Array("comm_AutoURL", comm_AutoURL), Array("comm_AutoKEY", comm_AutoKEY), Array("comm_Email", comm_Email), Array("comm_WebSite", comm_WebSite))
			edit = Sys.doRecord("blog_Comment", Arrays, "update", "comm_ID", comm_ID)
			Application.Lock()
			Application(Sys.CookieName & "_Comment_Edit") = True
			Application.UnLock()
		Else
			edit = Array(False, "您没有权限编辑该评论")
		End If
	End Function
	
	Public Function del
		If Not Asp.IsInteger(comm_ID) Then
			del = Array(False, "非法参数")
			Exit Function
		End If
		If (stat_CommentDel And (memName = comm_Author)) Or stat_Admin Then
			Arrays = Array(Array("blog_Comment", "comm_ID", comm_ID))
			del = Sys.doRecDel(Arrays)
			Application.Lock()
			Application(Sys.CookieName & "_Comment_Edit") = True
			Application.UnLock()
		Else
			del = Array(False, "您没有权限删除该评论")
		End If
	End Function
	
	Public Function PostReply
		If Not stat_Admin Then
			PostReply = Array(False, "您没有权限编辑或发表回复")
			Exit Function
		End If
		Arrays = Array(Array("reply_author", memName), Array("reply_time", Now()), Array("replay_content", replay_content))
		PostReply = Sys.doRecord("blog_Comment", Arrays, "update", "comm_ID", comm_ID)
	End Function
	
	Public Function delReply
		If Not stat_Admin Then
			editReply = Array(False, "您没有权限删除回复")
			Exit Function
		End If
		replay_content = ""
		delReply = PostReply
	End Function
	
	Public Function Aduit
		If Not stat_Admin Then
			Aduit = Array(False, "您没有权限审核评论")
			Exit Function
		End If
		Arrays = Array(Array("comm_IsAudit", comm_IsAudit))
		Aduit = Sys.doRecord("blog_Comment", Arrays, "update", "comm_ID", comm_ID)
		Application.Lock()
		Application(Sys.CookieName & "_Comment_Edit") = True
		Application.UnLock()
	End Function
	
End Class
%>