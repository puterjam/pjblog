<%
'*************** PJblog4 基本设置 *******************
' PJblog4 Copyright 2009
' Update: 2009-11-03
' Author: evio
' Owner : PuterJam
' Mode  : Class
' MoName: Cache
'***************************************************
Dim Data : Set Data = New Sys_SoeData

Class Sys_SoeData

	Private Rs, Arrays, SQL
	
	' ***********************************************
	'	整站缓存类初始化
	' ***********************************************
	Private Sub Class_Initialize
    End Sub 
     
	' ***********************************************
	'	整站缓存类终结化
	' ***********************************************
    Private Sub Class_Terminate
    End Sub
	
	' ***********************************************
	'	全局缓存
	' ***********************************************
	Public Function NavList(ByVal i)
		If Not IsArray(Application(Sys.CookieName & "_NavList")) Then
			SQL = "cate_ID, cate_Name, cate_Intro, cate_OutLink, cate_URL, cate_Folder"
			'			0		1			2			3			4			5
			Set Rs = Conn.Execute("Select " & SQL & " From blog_Category Where cate_local in(0, 1) And cate_Secret=False Order By cate_Order ASC")
			If Rs.Bof or Rs.Eof Then
				ReDim Arrays(0, 0)
			Else
				Arrays = Rs.GetRows
			End If
			Set Rs = Nothing
			Application.Lock()
			Application(Sys.CookieName & "_NavList") = Arrays
			Application.UnLock()
		Else
			Arrays = Application(Sys.CookieName & "_NavList")
		End If
		NavList = Arrays
	End Function

End Class
%>