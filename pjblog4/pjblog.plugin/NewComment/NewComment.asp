<%
' **********************************
' 最新评论 Recent Comment For PJBlog4
' Author: lyowbeen 
' WebSite: http://www.yb13.cn
' **********************************

' --------------------------------
' 判断是否需要更新缓存
' --------------------------------
If Application(Sys.CookieName & "_Comment_Edit") Or Application(Sys.CookieName & "_plugin_Setting_NewComment") Then
	Application.Lock()
	Application(Sys.CookieName & "_Comment_Edit") = False
	Application(Sys.CookieName & "_plugin_Setting_NewComment") = False
	Application.UnLock()
	NewComment(2)
End If

Function NewComment(i)
If IsEmpty(Application(Sys.CookieName & "_NewComment")) Or Int(i) = 2 Then
	' ----------------------------
	' 读取插件基本设置信息
	' ----------------------------
	Dim TopNum, IsShowHidden, TitleLen, IsShowAuthor
	Model.open("NewComment")
	TopNum = Model.getKeyValue("TopNum")
	IsShowHidden = Model.getKeyValue("IsShowHidden")
	TitleLen = Model.getKeyValue("TitleLen")
	IsShowAuthor = Model.getKeyValue("IsShowAuthor")

	' ----------------------------
	' 从数据库读取内容
	' ----------------------------
	Dim Comment(), TmpNum, Comments, CommentTitle	
	TmpNum = 0 : ReDim Comment(-1)   '下标从0开始, 如果没有记录则下标为-1#
	If IsShowHidden="0" Then    '显示未审核评论#
		Set Comments = Conn.Execute("SELECT top "&TopNum&" comm_ID,blog_ID,comm_Author,comm_Content,comm_PostTime,comm_IsAudit FROM blog_Comment as C,blog_Content as T,blog_Category as A where C.blog_ID=T.log_ID and T.log_IsShow=true and T.log_CateID=A.cate_ID and A.cate_Secret=false order by C.comm_PostTime Desc")
	Else     
		Set Comments = Conn.Execute("SELECT top "&TopNum&" comm_ID,blog_ID,comm_Author,comm_Content,comm_PostTime,comm_IsAudit FROM blog_Comment as C,blog_Content as T,blog_Category as A where C.blog_ID=T.log_ID and T.log_IsShow=true and T.log_CateID=A.cate_ID and A.cate_Secret=false and comm_IsAudit order by C.comm_PostTime Desc")
	End If

	' ----------------------------
	' 从缓存读取模板内容
	' ----------------------------
	Dim Template, TmpStr
	Plus.open("NewComment")
	Template = Split(Asp.UnCheckStr(Plus.getSingleTemplate("NewComment")), "|$|")
		
	Do While Not Comments.Eof
		ReDim preserve Comment(TmpNum)
		If Comments("comm_IsAudit") then
			If IsShowAuthor = "0" Then    '显示作者#
				CommentTitle = Asp.CutStr("<b>"&Comments("comm_Author")&": </b>"&Comments("comm_Content"), TitleLen)
			Else
				CommentTitle = Asp.CutStr(Comments("comm_Content"), TitleLen)
			End If
		Else
			CommentTitle = "[未审核评论]"
		End If
		' ----------------------------
		' 替换模板内容
		' ----------------------------		
		TmpStr = Template(1)
		TmpStr = Replace(TmpStr, "<#id#>", Init.doArticleUrl(Comments("blog_ID")))
		TmpStr = Replace(TmpStr, "<#Author#>", Comments("comm_Author"))
		TmpStr = Replace(TmpStr, "<#Time#>", Comments("comm_postTime"))
		TmpStr = Replace(TmpStr, "<#Title#>", CommentTitle)
		Comment(TmpNum)=TmpStr
		Comments.MoveNext
		TmpNum = TmpNum+1
	Loop
	
	' ----------------------------
	' 得到最新评论并写入缓存
	' ----------------------------	
	TmpStr = ""
	For TmpNum = 0 to Ubound(Comment)
		TmpStr = TmpStr & Comment(TmpNum)
	Next
	NewComment = Template(0) & TmpStr & Template(2)
	Application.Lock()
	Application(Sys.CookieName & "_NewComment") = NewComment
	Application.UnLock()
Else
	NewComment = Application(Sys.CookieName & "_NewComment")
End If
End Function
%>
	PluginTempValue = ['NewComment', '<%=outputSideBar(NewComment(1))%>'];
	PluginOutPutString.push(PluginTempValue);
