<%
Dim Article
Set Article = New Sys_Article

Class Sys_Article

	Public log_CateID, log_editType, log_IsDraft, log_Title, log_cname, log_ctype, log_weather, log_Level, log_comorder, log_DisComment, log_IsTop, log_Meta, log_KeyWords, log_Description, log_From, log_FromURL, log_Intro, log_Content, log_PostTime, log_ubbFlags, log_IsShow, log_tag, log_Author, log_ID
	
	Private Str, Arrays, this

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
		If stat_AddAll <> True And stat_Add <> True Then
            Add = Array(False, lang.Set.Asp(25))
            Exit Function
        End If
		Arrays = Array(Array("log_CateID", log_CateID), Array("log_editType", log_editType), Array("log_IsDraft", log_IsDraft), Array("log_Title", log_Title), Array("log_cname", log_cname), Array("log_ctype", log_ctype), Array("log_weather", log_weather), Array("log_Level", log_Level), Array("log_comorder", log_comorder), Array("log_DisComment", log_DisComment), Array("log_IsTop", log_IsTop), Array("log_KeyWords", log_KeyWords), Array("log_Description", log_Description), Array("log_From", log_From), Array("log_FromURL", log_FromURL), Array("log_Intro", log_Intro), Array("log_Content", log_Content), Array("log_PostTime", log_PostTime), Array("log_ubbFlags", log_ubbFlags), Array("log_IsShow", log_IsShow), Array("log_tag", log_tag), Array("log_Meta", log_Meta), Array("log_Author", memName))
		Add = Sys.doRecord("blog_Content", Arrays, "insert", "log_ID", "")
		Application.Lock()
		Application(Sys.CookieName & "_Article_Edit") = True
		Application.UnLock()
	End Function
	
	Public Function edit
		If stat_EditAll <> True And (stat_Edit And log_Author = memName) <> True Then
			edit = Array(False, lang.Set.Asp(106))
			Exit Function
		End If
		Arrays = Array(Array("log_CateID", log_CateID), Array("log_IsDraft", log_IsDraft), Array("log_Title", log_Title), Array("log_cname", log_cname), Array("log_ctype", log_ctype), Array("log_weather", log_weather), Array("log_Level", log_Level), Array("log_comorder", log_comorder), Array("log_DisComment", log_DisComment), Array("log_IsTop", log_IsTop), Array("log_KeyWords", log_KeyWords), Array("log_Description", log_Description), Array("log_From", log_From), Array("log_FromURL", log_FromURL), Array("log_Intro", log_Intro), Array("log_Content", log_Content), Array("log_ubbFlags", log_ubbFlags), Array("log_IsShow", log_IsShow), Array("log_tag", log_tag), Array("log_Meta", log_Meta), Array("log_Author", memName))
		edit = Sys.doRecord("blog_Content", Arrays, "update", "log_ID", log_ID)
		Application.Lock()
		Application(Sys.CookieName & "_Article_Edit") = True
		Application.UnLock()
	End Function
	
	Public Function del
		Application.Lock()
		Application(Sys.CookieName & "_Article_Edit") = True
		Application.UnLock()
	End Function
	
	Public Sub Articleopen(ByVal id)
		Set this = Conn.Execute("Select * From blog_Content Where log_ID=" & id)
			If this.Bof Or this.Eof Then
				Exit Sub
			Else
				log_CateID = this("log_CateID").value
				log_editType = this("log_editType").value
				log_IsDraft = this("log_IsDraft").value
				log_Title = this("log_Title").value
				log_Author = this("log_Author").value
				log_cname = this("log_cname").value
				log_ctype = this("log_ctype").value
				log_weather = this("log_weather").value
				log_Level = this("log_Level").value
				log_comorder = this("log_comorder").value
				log_DisComment = this("log_DisComment").value
				log_IsTop = this("log_IsTop").value
				log_Meta = this("log_Meta").value
				log_KeyWords = this("log_KeyWords").value
				log_Description = this("log_Description").value
				log_From = this("log_From").value
				log_FromURL = this("log_FromURL").value
				log_Intro = this("log_Intro").value
				log_Content = this("log_Content").value
				log_PostTime = this("log_PostTime").value
				log_ubbFlags = this("log_ubbFlags").value
				log_IsShow = this("log_IsShow").value
				log_tag = this("log_tag").value
			End If
		Set this = Nothing
	End Sub

End Class
%>