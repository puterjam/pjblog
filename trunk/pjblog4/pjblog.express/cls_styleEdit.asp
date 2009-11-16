<%
Dim Sd
Set Sd = New StyleEdit
Class StyleEdit

	Private Mud, Deep, page
	
	' ***********************************************
	'	类初始化
	' ***********************************************
	Private Sub Class_Initialize
		Set Mud = New template
		Deep = "../"
		Mud.Path = Deep & "pjblog.template/" & Trim(blog_DefaultSkin)
    End Sub 
     
	' ***********************************************
	'	类终结化
	' ***********************************************
    Private Sub Class_Terminate
		Set Mud = Nothing
    End Sub
	
	' ***********************************************
	'	全局变量
	' ***********************************************
	Public Sub General
		Mud.Sets("SiteName") = SiteName
		Mud.Sets("SiteURL") = SiteURL
		Mud.Sets("blogabout") = blogabout
		Mud.Sets("Skin") = blog_DefaultSkin
		Mud.Sets("blogTitle") = blog_Title
		Mud.Sets("master") = blog_master
		Mud.Sets("email") = blog_email
		Mud.Sets("cookie") = Sys.CookieName
		Mud.Sets("blogPerPage") = blogPerPage
	End Sub
	
	' ***********************************************
	'	静态化首页方法
	' ***********************************************
	Public Function default
		'On Error Resume Next
		Dim PageCount, c
		If Int(blog_LogNums) Mod Int(blogPerPage) = 0 Then
			PageCount = Int(blog_LogNums) / Int(blogPerPage)
		Else
			PageCount = Int(Int(blog_LogNums) / Int(blogPerPage)) + 1
		End If
		Mud.FileName = "index.html"
		Mud.TemplateContent = ""
		Mud.open
		Call General
		Mud.Sets("KeyWords") = blog_KeyWords
		Mud.Sets("Description") = blog_Description
		Mud.PageUrl = "index_{$page}.html"
		page = 1
		Mud.CurrentPage = Asp.CheckPage(page)
		Mud.Buffer
		'If Err.Number > 0 Then
			'default = Array(False, Err.Description)
		'Else
			default = Array(True, Mud.TemplateContent)
		'End If
	End Function
	
	' ***********************************************
	'	内容页静态化方法
	' ***********************************************
	Public Function Article(ByVal id)
		'On Error Resume Next
		Dim cPath, Rs, CateFolder, Ts, cCate, Title
		Set Rs = Conn.Execute("Select log_CateID, log_cname, log_ctype, log_Title From blog_Content Where log_ID=" & Int(id) & " And log_IsShow=True And log_IsDraft=False")
		If Rs.Bof Or Rs.Eof Then
			cPath = ""
		Else
			Title = Rs(3).value
			Set Ts = Server.CreateObject("Adodb.RecordSet")
			Ts.open "Select cate_Folder From blog_Category Where cate_ID=" & Int(Rs(0).value), Conn, 1, 1
			If Ts.Bof Or Ts.Eof Then
				cPath = ""
			Else
				cCate = Ts(0).value
				cPath = Init.ArticleUrl("../", cCate, Rs(1).value, Rs(2).value)
			End If
			Ts.Close
			Set Ts = Nothing
		End If
		If Len(cPath) > 0 Then
			Mud.FileName = "Article.html"
			Mud.open
			Call General
			Call ArticleFlied(id, Mud)
			Mud.Buffer
			Mud.Save(cPath)
			Article = Array(True, cPath, Title)
		Else
			Article = Array(False, "路径不正确", Title)
		End If
		'If Err.Number > 0 Then
			'Article = Array(False, Err.Description, Title)
		'End If
	End Function
	
	Private Sub ArticleFlied(ByVal i, ByRef o)
		Dim Rs, j
		Set Rs = Conn.Execute("Select * From blog_Content Where log_ID=" & Int(i))
			If Rs.Bof Or Rs.Eof Then
				Exit Sub
			Else
				o.Sets("KeyWord") = Rs("log_KeyWords").value
				o.Sets("Description") = Rs("log_KeyWords").value
				For j = 0 To Rs.Fields.Count - 1
					o.Sets(Rs(j).Name) = Rs(j).value
					'Response.Write(Rs(j).Name & " : " & Rs(j).value & "<br />")
				Next
			End If
		Set Rs = Nothing
	End Sub
	
	' ***********************************************
	'	静态化分类页方法
	' ***********************************************
	Public Function category(ByVal CateID, ByVal Folder)
		On Error Resume Next
		Dim PageCount, c
		If Int(blog_LogNums) Mod Int(blogPerPage) = 0 Then
			PageCount = Int(blog_LogNums) / Int(blogPerPage)
		Else
			PageCount = Int(Int(blog_LogNums) / Int(blogPerPage)) + 1
		End If
		For c = 1 To PageCount
			Mud.FileName = "category.html"
			Mud.TemplateContent = ""
			Mud.open
			Call General
			Mud.Sets("CateID") = Int(CateID)
			Mud.Sets("KeyWords") = blog_KeyWords
			Mud.Sets("Description") = blog_Description
			If Len(Folder) > 0 Then
				Mud.PageUrl = "cate_" & Folder & "_" & c & ".html"
			Else
				Mud.PageUrl = "cate_" & CateID & "_" & c & ".html"
			End If
			page = Trim(Asp.CheckStr(c))
			Mud.CurrentPage = Asp.CheckPage(page)
			Mud.Buffer
			If Len(Folder) > 0 Then
				Mud.Save(Deep & "html/cate_" & Folder & "_" & c & ".html")
			Else
				Mud.Save(Deep & "html/cate_" & CateID & "_" & c & ".html")
			End If
		Next
		If Err.Number > 0 Then
			category = Array(False, Err.Description, Folder)
		Else
			category = Array(True, "html/cate_" & Folder & ".html", Folder)
		End If
	End Function
	
	
End Class
%>