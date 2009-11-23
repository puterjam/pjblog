<%
Dim Temp
Set Temp = New Sys_Template

Class Sys_Template

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
	
	Public Function doUpdate(ByVal f1, ByVal f2)
		Arrays = Array(Array("blog_DefaultSkin", f1), Array("blog_style", f2))
		doUpdate = Sys.doRecord("blog_Info", Arrays, "update", "blog_ID", 1)
	End Function
	
End Class
%>