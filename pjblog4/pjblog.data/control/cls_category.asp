<%
Dim Category
Set Category = New Sys_Category

Class Sys_Category

	Public cate_ID, cate_Order, cate_icon, cate_Name, cate_Intro, cate_Folder, cate_URL, cate_local, cate_Secret, cate_count, cate_OutLink, cate_Lock
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
		Arrays = Array(Array("cate_Order", cate_Order), Array("cate_icon", cate_icon), Array("cate_Name", cate_Name), Array("cate_Intro", cate_Intro), Array("cate_Folder", cate_Folder), Array("cate_URL", cate_URL), Array("cate_local", cate_local), Array("cate_Secret", cate_Secret), Array("cate_count", cate_count), Array("cate_OutLink", cate_OutLink), Array("cate_Lock", cate_Lock))
		Add = Sys.doRecord("blog_Category", Arrays, "insert", "cate_ID", "")
	End Function
	
	Public Function edit
		Arrays = Array(Array("cate_Order", cate_Order), Array("cate_icon", cate_icon), Array("cate_Name", cate_Name), Array("cate_Intro", cate_Intro), Array("cate_Folder", cate_Folder), Array("cate_URL", cate_URL), Array("cate_local", cate_local), Array("cate_Secret", cate_Secret), Array("cate_count", cate_count), Array("cate_OutLink", cate_OutLink), Array("cate_Lock", cate_Lock))
		Add = Sys.doRecord("blog_Category", Arrays, "update", "cate_ID", cate_ID)
	End Function
	
End Class
%>