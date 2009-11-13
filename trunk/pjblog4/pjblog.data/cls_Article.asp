<%
Dim Article
Set Article = New Sys_Article

Class Sys_Article

	Public log_CateID, log_editType, log_IsDraft, log_Title, log_cname, log_ctype, log_weather, log_Level, log_comorder, log_DisComment, log_IsTop, log_Meta, log_KeyWords, log_Description, log_From, log_FromURL, log_Intro, log_Content, log_PostTime, log_ubbFlags, log_IsShow, log_tag
	
	Private Str, Arrays

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
		Arrays = Array(Array("log_CateID", log_CateID), Array("log_editType", log_editType), Array("log_IsDraft", log_IsDraft), Array("log_Title", log_Title), Array("log_cname", log_cname), Array("log_ctype", log_ctype), Array("log_weather", log_weather), Array("log_Level", log_Level), Array("log_comorder", log_comorder), Array("log_DisComment", log_DisComment), Array("log_IsTop", log_IsTop), Array("log_KeyWords", log_KeyWords), Array("log_Description", log_Description), Array("log_From", log_From), Array("log_FromURL", log_FromURL), Array("log_Intro", log_Intro), Array("log_Content", log_Content), Array("log_PostTime", log_PostTime), Array("log_ubbFlags", log_ubbFlags), Array("log_IsShow", log_IsShow), Array("log_tag", log_tag), Array("log_Meta", log_Meta))
		Add = Sys.doRecord("blog_Content", Arrays, "insert", "log_ID", "")
		Sys.doGet("Update blog_Info Set blog_LogNums = blog_LogNums + 1")
	End Function
	
	Public Function edit
	
	End Function
	
	Public Function del
		
	End Function

End Class
%>