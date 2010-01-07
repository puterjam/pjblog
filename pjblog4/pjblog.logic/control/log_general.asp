<!--#include file = "../../include.asp" -->
<!--#include file = "../../pjblog.data/control/cls_general.asp" -->
<%
'*************** PJblog4 基本设置 *******************
' PJblog4 Copyright 2009
' Update: 2009-11-03
' Author: evio
' Owner : PuterJam
' Mode  : Class
' MoName: User Model
'***************************************************
Dim dogeneral
Set dogeneral = New do_general
Set dogeneral = Nothing

Class do_general
	
	Public Property Get Action
		Action = Request.QueryString("action")
	End Property
	
	Private ReConSio

	' ***********************************************
	'	用户操作方法类初始化
	' ***********************************************
	Private Sub Class_Initialize
		Select Case Action
			Case "SaveGeneral" Call SaveGeneral
			Case "SaveTipWord" Call SaveTipWord
			Case "BaseSetting" Call BaseSetting
			Case "ArtComm" Call ArtComm
			Case "SEO" Call SEO
			Case "WebMode" Call WebMode
			Case "WapMail" Call WapMail
			Case "RegFilter" Call RegFilter
			Case "ReBuidWebData" Call ReBuidWebData
			Case "ClearVistor" Call ClearVistor
			Case "AppCount" Call AppCount
			Case "AppRemove" Call AppRemove
			Case "savePing" Call savePing
		End Select
    End Sub 
     
	' ***********************************************
	'	用户操作方法类终结化
	' ***********************************************
    Private Sub Class_Terminate
		Set general = Nothing
		Sys.Close
    End Sub
	
	Private Sub SaveGeneral
		general.SiteName = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("SiteName"))))
		general.blog_Title = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_Title"))))
		general.blog_master = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_master"))))
		general.blog_email = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_email"))))
		general.blog_KeyWords = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_KeyWords"))))
		general.blog_Description = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_Description"))))
		If Right(Asp.CheckStr(Request.Form("SiteURL")), 1) <> "/" Then
			general.SiteURL = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("SiteURL")))) & "/"
		Else
			general.SiteURL = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("SiteURL"))))
		End If
		general.blog_about = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_about"))))
		general.blog_html = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_html"))))
		general.blog_postFile = Int(Asp.checkURL(Asp.CheckStr(Request.Form("blog_postFile"))))
		general.blog_SplitType = CBool(Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_SplitType")))))
		general.blog_introChar = Int(Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_introChar")))))
		general.blog_introLine = Int(Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_introLine")))))
		general.blog_PerPage = Int(Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blogPerPage")))))
		If Asp.IsInteger(Request.Form("blog_ImgLink")) And Int(Request.Form("blog_ImgLink")) = 1 Then
			general.blog_ImgLink = True
		Else
			general.blog_ImgLink = False
		End If
		general.blog_commPage = Int(Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blogcommpage")))))
		general.blog_commTimerout = Int(Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_commTimerout")))))
		general.blog_commLength = Int(Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_commLength")))))
		If Asp.IsInteger(Request.Form("blog_validate")) And Int(Request.Form("blog_validate")) = 1 Then
			general.blog_validate = True
		Else
			general.blog_validate = False
		End If
		If Asp.IsInteger(Request.Form("blog_commUBB")) And Int(Request.Form("blog_commUBB")) = 1 Then
			general.blog_commUBB = True
		Else
			general.blog_commUBB = False
		End If
		If Asp.IsInteger(Request.Form("blog_commIMG")) And Int(Request.Form("blog_commIMG")) = 1 Then
			general.blog_commIMG = True
		Else
			general.blog_commIMG = False
		End If
		If Asp.IsInteger(Request.Form("blog_commAduit")) And Int(Request.Form("blog_commAduit")) = 1 Then
			general.blog_commAduit = True
		Else
			general.blog_commAduit = False
		End If
		general.blog_Disregister = Int(Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_Disregister")))))
		general.blog_CountNum = Int(Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_CountNum")))))
		general.blog_FilterName = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_FilterName"))))
		general.blog_FilterIP = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_FilterIP"))))
		ReConSio = general.SaveGeneral
		If ReConSio(0) Then Call Cache.GlobalCache(2)
		Session(Sys.CookieName & "_ShowMsg") = True
		Session(Sys.CookieName & "_MsgText") = "保存信息成功!"
		RedirectUrl(RedoUrl)
	End Sub
	
	Private Sub SaveTipWord
		Response.Write("{Suc : true, Info : 'OK'}")
	End Sub
	
	Private Sub BaseSetting
		general.SiteName = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("SiteName"))))
		general.blog_Title = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_Title"))))
		general.blog_master = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_master"))))
		If Right(Asp.CheckStr(Request.Form("SiteURL")), 1) <> "/" Then
			general.SiteURL = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("SiteURL")))) & "/"
		Else
			general.SiteURL = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("SiteURL"))))
		End If
		general.blog_about = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_about"))))
		ReConSio = general.BaseSetting
		If ReConSio(0) Then
			Response.Write("{Suc : true}")
		Else
			Response.Write("{Suc : false}")
		End If
	End Sub
	
	Private Sub ArtComm
		general.blog_html = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_html"))))
		general.blog_SplitType = CBool(Int(Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_SplitType"))))))
		general.blog_introChar = Int(Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_introChar")))))
		general.blog_introLine = Int(Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_introLine")))))
		general.blog_PerPage = Int(Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blogPerPage")))))
		general.blog_commPage = Int(Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blogcommpage")))))
		general.blog_commTimerout = Int(Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_commTimerout")))))
		general.blog_commLength = Int(Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_commLength")))))
		If Asp.IsInteger(Request.Form("blog_validate")) And Int(Request.Form("blog_validate")) = 1 Then
			general.blog_validate = True
		Else
			general.blog_validate = False
		End If
		If Asp.IsInteger(Request.Form("blog_commUBB")) And Int(Request.Form("blog_commUBB")) = 1 Then
			general.blog_commUBB = True
		Else
			general.blog_commUBB = False
		End If
		If Asp.IsInteger(Request.Form("blog_commIMG")) And Int(Request.Form("blog_commIMG")) = 1 Then
			general.blog_commIMG = True
		Else
			general.blog_commIMG = False
		End If
		If Asp.IsInteger(Request.Form("blog_commAduit")) And Int(Request.Form("blog_commAduit")) = 1 Then
			general.blog_commAduit = True
		Else
			general.blog_commAduit = False
		End If
		ReConSio = general.ArtComm
		If ReConSio(0) Then
			Response.Write("{Suc : true}")
		Else
			Response.Write("{Suc : false}")
		End If
	End Sub
	
	Private Sub SEO
		general.blog_KeyWords = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_KeyWords"))))
		general.blog_Description = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_Description"))))
		If Asp.IsInteger(Request.Form("blog_IsPing")) And Int(Request.Form("blog_IsPing")) = 1 Then
			general.blog_IsPing = True
		Else
			general.blog_IsPing = False
		End If
		ReConSio = general.SEO
		If ReConSio(0) Then
			Response.Write("{Suc : true}")
		Else
			Response.Write("{Suc : false}")
		End If
	End Sub
	
	Private Sub WebMode
		general.blog_postFile = Int(Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_postFile")))))
		ReConSio = general.WebMode
		If ReConSio(0) Then
			Response.Write("{Suc : true}")
		Else
			Response.Write("{Suc : false}")
		End If
	End Sub
	
	Private Sub WapMail
		general.blog_smtp = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_smtp"))))
		general.blog_email = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_email"))))
		general.blog_emailpass = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_emailpass"))))
		general.blog_sendmailName = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_sendmailName"))))
		
		general.blog_wapNum = Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_wapNum"))))
		If Asp.CheckStr(Request.Form("blog_wapImg")) = "1" Then general.blog_wapImg = True Else general.blog_wapImg = False
		If Asp.CheckStr(Request.Form("blog_wapHTML")) = "1" Then general.blog_wapHTML = True Else general.blog_wapHTML = False
		If Asp.CheckStr(Request.Form("blog_wapLogin")) = "1" Then general.blog_wapLogin = True Else general.blog_wapLogin = False
		If Asp.CheckStr(Request.Form("blog_wapComment")) = "1" Then general.blog_wapComment = True Else general.blog_wapComment = False
		If Asp.CheckStr(Request.Form("blog_wap")) = "1" Then general.blog_wap = True Else general.blog_wap = False
		If Asp.CheckStr(Request.Form("blog_wapURL")) = "1" Then general.blog_wapURL = True Else general.blog_wapURL = False
		ReConSio = general.WapMail
		If ReConSio(0) Then
			Response.Write("{Suc : true}")
		Else
			Response.Write("{Suc : false}")
		End If
	End Sub
	
	Private Sub RegFilter
		general.blog_Disregister = Int(Trim(Asp.checkURL(Asp.CheckStr(Request.Form("blog_Disregister")))))
		ReConSio = general.RegFilter
		If ReConSio(0) Then
			Response.Write("{Suc : true}")
		Else
			Response.Write("{Suc : false}")
		End If
	End Sub
	
	Private Sub ReBuidWebData
		Dim blog_Content_count, blog_Comment_count, ContentCount, TBCount, Count_Member
		ContentCount = 0
		TBCount = conn.Execute("select count(*) from blog_Trackback")(0)
		Count_Member = conn.Execute("select count(*) from blog_Member")(0)
		conn.Execute("update blog_Info set blog_tbNums="&TBCount)
		conn.Execute("update blog_Info set blog_MemNums="&Count_Member)
		Set blog_Content_count = conn.Execute("SELECT log_CateID, Count(log_CateID) FROM blog_Content where log_IsDraft=false GROUP BY log_CateID")
		Set blog_Comment_count = conn.Execute("SELECT Count(*) FROM blog_Comment")
		Do While Not blog_Content_count.EOF
			ContentCount = ContentCount + blog_Content_count(1)
			conn.Execute("update blog_Category set cate_count="&blog_Content_count(1)&" where cate_ID="&blog_Content_count(0))
			blog_Content_count.movenext
		Loop
		conn.Execute("update blog_Info set blog_LogNums="&ContentCount)
		conn.Execute("update blog_Info set blog_CommNums="&blog_Comment_count(0))
		Response.Write("{Suc : true}")
	End Sub
	
	Private Sub ClearVistor
		conn.Execute("delete * from blog_Counter")
		Response.Write("{Suc : true}")
	End Sub
	
	Private Sub AppCount
		Response.Write(Application.Contents.Count)
	End Sub
	
	Private Sub AppRemove
		Application.Contents.Remove(Request.QueryString("index"))
		Response.Write("{Suc : true}")
	End Sub
	
	Private Sub savePing
		Dim id, Ping_url, Ping_Name, Arrays
		id = Request.QueryString("id")
		If Not Asp.IsInteger(id) Then
			Response.Write("{Suc : false, Info : '非法参数'}")
		Else
			Ping_Name = Trim(Asp.CheckStr(Request.QueryString("Ping_Name")))
			Ping_url = Trim(Asp.CheckStr(Request.QueryString("Ping_url")))
			Arrays = Array(Array("Ping_Name", Ping_Name), Array("Ping_url", Ping_url))
			ReConSio = Sys.doRecord("blog_Ping", Arrays, "update", "Ping_ID", id)
			If ReConSio(0) Then
				Response.Write("{Suc : true, Info : '更新成功'}")
			Else
				Response.Write("{Suc : false, Info : '" & ReConSio(1) & "'}")
			End If
		End If
	End Sub
	
End Class
%>