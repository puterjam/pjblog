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
	'	导航数据
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
	
	' ***********************************************
	'	首页日志数据
	' ***********************************************
	Public Function ArticleList(ByVal i)
		If Not IsArray(Application(Sys.CookieName & "_ArticleList")) Then
			SQL = "log_ID, log_Title, log_Author, log_PostTime, log_Intro, log_Content, log_CateID, log_CommNums, log_ViewNums, log_QuoteNums"
			'			0		1			2			3			4			5			6			7				8	
'		9			
			Set Rs = Conn.Execute("Select " & SQL & " From blog_Content Where log_IsShow=True And log_IsDraft=False Order By log_PostTime Desc")
			If Rs.Bof or Rs.Eof Then
				ReDim Arrays(0, 0)
			Else
				Arrays = Rs.GetRows
			End If
			Set Rs = Nothing
			Application.Lock()
			Application(Sys.CookieName & "_ArticleList") = Arrays
			Application.UnLock()
		Else
			Arrays = Application(Sys.CookieName & "_ArticleList")
		End If
		ArticleList = Arrays
	End Function
	
	' ***********************************************
	'	获取分类信息
	' ***********************************************
	Public Function GetCateInfo(Byval ID, ByVal i)
			Dim Rs
			Set Rs = Conn.Execute("Select * From blog_Category Where cate_ID=" & ID)
			If Rs.Bof Or Rs.Eof Then
			Response.Write(i & ",")
				GetCateInfo = ""
			Else
			
				GetCateInfo = Rs(i).value
			End If
			Set Rs = Nothing
	End Function

End Class
%>