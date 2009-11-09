<%
Class Sys_Article

	Public log_Title, log_Cname, log_Ctype, log_weather, log_Level, log_comorder, log_DisComment, log_IsShow, log_Meta, log_KeyWords, log_Description, log_From, log_PostTime, log_Tag, log_Content
	
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
		Arrays = Array(Array("log_Title", log_Title), Array("log_Cname", log_Cname), Array("log_Ctype", log_Ctype), Array("log_weather", log_weather), Array("log_Level", log_Level), Array("log_comorder", log_comorder), Array("log_DisComment", log_DisComment), Array("log_IsShow", log_IsShow), Array("log_Meta", log_Meta), Array("log_KeyWords", log_KeyWords), Array("log_Description", log_Description), Array("log_From", log_From), Array("log_PostTime", log_PostTime), Array("log_Tag", log_Tag), Array("log_Content", log_Content))
		Str = Sys.doRecord()
	End Function
	
	Public Function edit
	
	End Function
	
	Public Function del
		
	End Function

End Class
%>