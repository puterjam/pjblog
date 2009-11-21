<%
If Application(Sys.CookieName & "_Comment_Edit") Or Application(Sys.CookieName & "_plugin_Setting_NewComment") Then
	Application.Lock()
	Application(Sys.CookieName & "_Comment_Edit") = False
	Application(Sys.CookieName & "_plugin_Setting_NewComment") = False
	Application.UnLock()
	NewComment(2)
End If

Function NewComment(i)
If IsEmpty(Application(Sys.CookieName & "_NewComment")) Or Int(i) = 2 Then
	dim TopNum, IsShowHidden, TitleLen, IsShowAuthor, Comment(), TmpNum, Comments, CommentTitle
	
	Model.open("NewComment")
	TopNum = Model.getKeyValue("TopNum")
	IsShowHidden = Model.getKeyValue("IsShowHidden")
	TitleLen = Model.getKeyValue("TitleLen")
	IsShowAuthor = Model.getKeyValue("IsShowAuthor")
	
	TmpNum = 0 : ReDim Comment(-1)   '下标从0开始, 如果没有记录则下标为-1#
	If IsShowHidden="0" Then    '显示未审核评论#
		Set Comments = Conn.Execute("SELECT top "&TopNum&" comm_ID,blog_ID,comm_Author,comm_Content,comm_PostTime,comm_IsAudit FROM blog_Comment as C,blog_Content as T,blog_Category as A where C.blog_ID=T.log_ID and T.log_IsShow=true and T.log_CateID=A.cate_ID and A.cate_Secret=false order by C.comm_PostTime Desc")
	Else     
		Set Comments = Conn.Execute("SELECT top "&TopNum&" comm_ID,blog_ID,comm_Author,comm_Content,comm_PostTime,comm_IsAudit FROM blog_Comment as C,blog_Content as T,blog_Category as A where C.blog_ID=T.log_ID and T.log_IsShow=true and T.log_CateID=A.cate_ID and A.cate_Secret=false and comm_IsAudit order by C.comm_PostTime Desc")
	End If
	
	Do While Not Comments.Eof
		redim preserve Comment(TmpNum)
		If Comments("comm_IsAudit") then
			If IsShowAuthor = "0" Then    '显示作者#
				CommentTitle = Asp.CutStr("<b>"&Comments("comm_Author")&": </b>"&Comments("comm_Content"), TitleLen)
			Else
				CommentTitle = Asp.CutStr(Comments("comm_Content"), TitleLen)
			End If
		Else
			CommentTitle = "[未审核评论]"
		End If
		Comment(TmpNum)="<a href="""&Init.doArticleUrl(Comments("blog_ID"))&""" class=""sideA"" title="" "&Comments("comm_Author")&" 评论于 "&Comments("comm_postTime")&" "">"&CommentTitle&"</a>"
		Comments.MoveNext
		TmpNum = TmpNum+1
	Loop	
	
	dim TmpStr : TmpStr = ""
	For TmpNum = 0 to Ubound(Comment)
		TmpStr = TmpStr & Comment(TmpNum)
	Next
	NewComment = TmpStr
	Application.Lock()
	Application(Sys.CookieName & "_NewComment") = TmpStr
	Application.UnLock()
Else
	NewComment = Application(Sys.CookieName & "_NewComment")
End If
End Function
%>
	PluginTempValue = ['plugin_NewComment', '<%=outputSideBar(NewComment(1))%>'];
	PluginOutPutString.push(PluginTempValue);
