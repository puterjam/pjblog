<%
Dim web
Set web = New webConfig

Class webConfig

	Private Mud, Deep
	
	' ***********************************************
	'	类初始化
	' ***********************************************
	Private Sub Class_Initialize
		Set Mud = New template
    End Sub 
     
	' ***********************************************
	'	类终结化
	' ***********************************************
    Private Sub Class_Terminate
		Deep = "../"
		Set Mud =  = Nothing
		Mud.Path = Deep & "pjblog.template/" & Trim(blog_DefaultSkin)
		Sys.Close
    End Sub
	
	' ***********************************************
	'	全局变量
	' ***********************************************
	Public Sub General(ByRef Mud)
		Mud.Sets("SiteName") = SiteName
		Mud.Sets("SiteURL") = SiteURL
		Mud.Sets("blogabout") = blogabout
		Mud.Sets("Skin") = blog_DefaultSkin
		Mud.Sets("blogTitle") = blog_Title
		Mud.Sets("master") = blog_master
		Mud.Sets("email") = blog_email
	End Sub
	
	Public Sub default(ByRef Mud)
		Mud.FileName = "index.html"
		Mud.open
		Mud.Sets("KeyWords") = blog_KeyWords
		Mud.Sets("Description") = blog_Description
		Mud.Buffer
		Mud.Save(Deep & "html/index.html")
	End Sub
	
End Class
%>