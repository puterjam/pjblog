<%
'*************** PJblog4 基本设置 *******************
' PJblog4 Copyright 2009
' Update: 2009-11-03
' Author: evio
' Owner : PuterJam
' Mode  : Class
' MoName: User Model
'***************************************************
Dim general : Set general = New c_general

Class c_general

	Public SiteName, blog_Title, blog_master, blog_email, SiteURL, blog_KeyWords, blog_Description, blog_about
	
	Private ReConSio

	' ***********************************************
	'	用户操作方法类初始化
	' ***********************************************
	Private Sub Class_Initialize
    End Sub 
     
	' ***********************************************
	'	用户操作方法类终结化
	' ***********************************************
    Private Sub Class_Terminate
    End Sub
	
	' ***********************************************
	'	保存全局信息
	' ***********************************************
	Public Function SaveGeneral
		Dim Arrays
		Arrays = Array(Array("blog_Name", SiteName), Array("blog_Title", blog_Title), Array("blog_master", blog_master), Array("blog_email", blog_email), Array("blog_KeyWords", blog_KeyWords), Array("blog_Description", blog_Description), Array("blog_URL", SiteURL), Array("blog_affiche",""), Array("blog_about", blog_about))
		ReConSio = Sys.doRecord("blog_Info", Arrays, "update", "blog_ID", 1)
		SaveGeneral = ReConSio
	End Function
	
End Class
%>