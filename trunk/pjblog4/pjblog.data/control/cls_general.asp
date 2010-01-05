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

	Public SiteName, blog_Title, blog_master, blog_email, SiteURL, blog_KeyWords, blog_Description, blog_about, blog_html, blog_postFile, blog_SplitType, blog_introChar, blog_introLine, blog_PerPage, blog_ImgLink, blog_commPage, blog_commTimerout, blog_commLength, blog_validate, blog_commUBB, blog_commIMG, blog_Disregister, blog_CountNum, blog_FilterName, blog_FilterIP, blog_commAduit, blog_IsPing
	
	Private ReConSio, Arrays

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
		Arrays = Array(Array("blog_Name", SiteName), Array("blog_Title", blog_Title), Array("blog_master", blog_master), Array("blog_email", blog_email), Array("blog_KeyWords", blog_KeyWords), Array("blog_Description", blog_Description), Array("blog_URL", SiteURL), Array("blog_affiche",""), Array("blog_about", blog_about), Array("blog_html", blog_html), Array("blog_postFile", blog_postFile), Array("blog_SplitType", blog_SplitType), Array("blog_introChar", blog_introChar), Array("blog_introLine", blog_introLine), Array("blog_PerPage", blog_PerPage), Array("blog_ImgLink", blog_ImgLink), Array("blog_commPage", blog_commPage), Array("blog_commTimerout", blog_commTimerout), Array("blog_commLength", blog_commLength), Array("blog_validate", blog_validate), Array("blog_commUBB", blog_commUBB), Array("blog_commIMG", blog_commIMG), Array("blog_Disregister", blog_Disregister), Array("blog_CountNum", blog_CountNum), Array("blog_FilterName", blog_FilterName), Array("blog_FilterIP", blog_FilterIP), Array("blog_commAduit", blog_commAduit))
		ReConSio = Sys.doRecord("blog_Info", Arrays, "update", "blog_ID", 1)
		Call Data.ArticleList(2)
		SaveGeneral = ReConSio
	End Function
	
	Public Function BaseSetting
		Arrays = Array(Array("blog_Name", SiteName), Array("blog_Title", blog_Title), Array("blog_master", blog_master), Array("blog_URL", SiteURL), Array("blog_about", blog_about))
		BaseSetting = Sys.doRecord("blog_Info", Arrays, "update", "blog_ID", 1)
		Call Cache.GlobalCache(2)
		Call Data.ArticleList(2)
	End Function
	
	Public Function ArtComm
		Arrays = Array(Array("blog_html", blog_html), Array("blog_SplitType", blog_SplitType), Array("blog_introChar", blog_introChar), Array("blog_introLine", blog_introLine), Array("blog_PerPage", blog_PerPage), Array("blog_commPage", blog_commPage), Array("blog_commTimerout", blog_commTimerout), Array("blog_commLength", blog_commLength), Array("blog_validate", blog_validate), Array("blog_commUBB", blog_commUBB), Array("blog_commIMG", blog_commIMG), Array("blog_commAduit", blog_commAduit))
		ArtComm = Sys.doRecord("blog_Info", Arrays, "update", "blog_ID", 1)
		Call Cache.GlobalCache(2)
		Call Data.ArticleList(2)
	End Function
	
	Public Function SEO
		Arrays = Array(Array("blog_KeyWords", blog_KeyWords), Array("blog_Description", blog_Description), Array("blog_IsPing", blog_IsPing))
		SEO = Sys.doRecord("blog_Info", Arrays, "update", "blog_ID", 1)
		Call Cache.GlobalCache(2)
		Call Data.ArticleList(2)
	End Function
	
	Public Function WebMode
		Arrays = Array(Array("blog_postFile", blog_postFile))
		WebMode = Sys.doRecord("blog_Info", Arrays, "update", "blog_ID", 1)
		Call Cache.GlobalCache(2)
		Call Data.ArticleList(2)
	End Function
	
End Class
%>