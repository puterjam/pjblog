<%
Dim link
Set link = New Sys_Link
Class Sys_Link

	Public link_ID, link_Name, link_URL, link_Image, link_Order, link_IsShow, link_IsMain, Link_ClassID, Link_Type
	
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
	
	Public Function edit
		Arrays = Array(Array("link_Name", link_Name), Array("link_URL", link_URL), Array("link_Image", link_Image), Array("link_Order", link_Order))
		edit = Sys.doRecord("blog_Links", Arrays, "update", "link_ID", link_ID)
	End Function
	
	Public Function addNewClass
		Arrays = Array(Array("link_Name", link_Name), Array("Link_Type", True))
		addNewClass = Sys.doRecord("blog_Links", Arrays, "insert", "link_ID", "")
	End Function
	
	Public Function addNewLink
		Arrays = Array(Array("link_Name", link_Name), Array("link_URL", link_URL), Array("link_Image", link_Image), Array("link_Order", link_Order), Array("Link_Type", False), Array("Link_ClassID", Link_ClassID))
		addNewLink = Sys.doRecord("blog_Links", Arrays, "insert", "link_ID", "")
	End Function
	
	Public Function del(ByVal id)
		del = Sys.doRecDel(Array(Array("blog_Links", "link_ID", id)))
	End Function
	
End Class
%>