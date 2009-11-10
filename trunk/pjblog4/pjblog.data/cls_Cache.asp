<%
'*************** PJblog4 基本设置 *******************
' PJblog4 Copyright 2009
' Update: 2009-11-03
' Author: evio
' Owner : PuterJam
' Mode  : Class
' MoName: Cache
'***************************************************
Dim Cache : Set Cache = New Sys_Cache

Class Sys_Cache
	
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
	Public Function GlobalCache(ByVal i)
		Dim Sql, log_Infos, blog_Infos
		If Not IsArray(Application(Sys.CookieName & "_blog_Infos")) Or Int(i) = 2 Then
				Sql = "select top 1 blog_Name,blog_URL,blog_PerPage,blog_LogNums,blog_CommNums,blog_MemNums," & _
              		"blog_VisitNums,blog_BookPage,blog_MessageNums,blog_commPage,blog_affiche," & _
              		"blog_about,blog_colPage,blog_colNums,blog_tbNums,blog_showtotal," & _
              		"blog_FilterName,blog_FilterIP,blog_commTimerout,blog_commUBB,blog_commImg," & _
              		"blog_postFile,blog_postCalendar,blog_DefaultSkin,blog_SkinName,blog_SplitType," & _
              		"blog_introChar,blog_introLine,blog_validate,blog_Title,blog_ImgLink," & _
              		"blog_commLength,blog_downLocal,blog_DisMod,blog_Disregister,blog_master,blog_email,blog_CountNum," & _
			  		"blog_wapNum,blog_wapImg,blog_wapHTML,blog_wapLogin,blog_wapComment,blog_wap,blog_wapURL,blog_KeyWords," & _
			  		"blog_Description, blog_SaveTime, blog_IsPing, blog_html" & _
              		" from blog_Info"
			  	Set log_Infos = Conn.Execute(Sql)
			  		blog_Infos = log_Infos.GetRows()
        		Set log_Infos = Nothing
				Application.Lock
        		Application(Sys.CookieName & "_blog_Infos") = blog_Infos
       			Application.UnLock
		Else
			blog_Infos = Application(Sys.CookieName & "_blog_Infos")
		End If
		GlobalCache = blog_Infos
	End Function
	
	' ***********************************************
	'	用户权限缓存
	' ***********************************************
	Public Function UserRight(ByVal i)
		Dim blog_Status
		If Not IsArray(Application(Sys.CookieName & "_blog_rights")) Or Int(i) = 2 Then
			Dim log_Status, SQL
				SQL = "select stat_name, stat_title, stat_Code, stat_attSize, stat_attType from blog_status"
				Set log_Status = Conn.Execute(SQL)
				blog_Status = log_Status.GetRows()
				Set log_Status = Nothing
				Application.Lock
				Application(Sys.CookieName & "_blog_rights") = blog_Status
				Application.UnLock
		Else
			blog_Status = Application(Sys.CookieName & "_blog_rights")
		End If
		UserRight = blog_Status
	End Function
	
	' ***********************************************
	'	分类缓存
	' ***********************************************
	Public Function CategoryCache(ByVal i)
		Dim blog_Category
		If Not IsArray(Application(Sys.CookieName & "_blog_Category")) Or Int(i) = 2 Then
			Dim log_Category, SQL
				SQL = "SELECT cate_ID,cate_Name,cate_Order,cate_Intro,cate_OutLink,cate_URL,cate_icon,cate_count,cate_Lock,cate_local,cate_Secret,cate_Folder FROM blog_Category ORDER BY cate_Order ASC"
				Set log_Category = Conn.Execute(SQL)
				blog_Category = log_Category.GetRows()
				Set log_Category = Nothing
				Application.Lock
				Application(Sys.CookieName & "_blog_Category") = blog_Category
				Application.UnLock
		Else
			blog_Category = Application(Sys.CookieName & "_blog_Category")
		End If
		CategoryCache = blog_Category
	End Function
	
	' ***********************************************
	'	分类缓存
	' ***********************************************
	Public Function SmiliesCache(ByVal i)
    If Not IsArray(Application(Sys.CookieName & "_blog_Smilies")) Or Int(i) = 2 Then
        Dim log_Smilies, log_SmiliesList, TempVar, Arr_Smilies
        Set log_Smilies = Conn.Execute("SELECT sm_ID,sm_Image,sm_Text FROM blog_Smilies")
        TempVar = ""
        Do While Not log_Smilies.EOF
            log_SmiliesList = log_SmiliesList&TempVar&log_Smilies("sm_ID")&"|"&log_Smilies("sm_Image")&"|"&log_Smilies("sm_Text")
            TempVar = ","
            log_Smilies.MoveNext
        Loop
        Set log_Smilies = Nothing
        Arr_Smilies = Split(log_SmiliesList, ",")
        Application.Lock
        Application(Sys.CookieName & "_blog_Smilies") = Arr_Smilies
        Application.UnLock
    Else
        Arr_Smilies = Application(Sys.CookieName & "_blog_Smilies")
    End If
	SmiliesCache = Arr_Smilies
End Function
	
End Class
%>