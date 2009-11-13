<!--#include file = "../include.asp" -->
<!--#include file = "../pjblog.data/cls_article.asp" -->
<!--#include file = "../pjblog.model/cls_tag.asp" -->
<!--#include file = "../pjblog.data/cls_webconfig.asp" -->
<!--#include file = "../pjblog.model/cls_template.asp" -->
<!--#include file = "../pjblog.model/cls_fso.asp" -->
<!--#include file = "../pjblog.model/cls_Stream.asp" -->
<%
Dim doArticle
Set doArticle = New do_Article
Set doArticle = Nothing

Class do_Article
'22
	Private log_CateID, log_editType, log_IsDraft, log_Title, log_cname, log_ctype, log_weather, log_Level, log_comorder, log_DisComment, log_IsTop, log_Meta, log_KeyWords, log_Description, log_From, log_FromURL, log_Intro, log_Content, log_PostTime, log_ubbFlags, log_IsShow, log_tag
		
	Private log_disImg, log_DisSM, log_DisURL, log_DisKey, logIntroCustom
	Private PubTimeType, PubTime, Message, log_IntroC
	
	Private Str, Arrays
	
	Public Property Get Action
		Action = Trim(Request.QueryString("action"))
	End Property

	' ***********************************************
	'	类初始化
	' ***********************************************
	Private Sub Class_Initialize
		Select Case Action
			Case "add" Call Add
		End Select
    End Sub 
     
	' ***********************************************
	'	类终结化
	' ***********************************************
    Private Sub Class_Terminate
		Sys.Close
    End Sub
	
	Private Sub Add
		log_CateID = Int(Trim(Asp.CheckUrl(Asp.CheckStr(Request.Form("log_CateID")))))
		log_editType = Int(Trim(Asp.CheckUrl(Asp.CheckStr(Request.Form("log_editType")))))
		log_IsDraft = Lcase(Trim(Asp.CheckUrl(Asp.CheckStr(Request.Form("log_IsDraft")))))
		log_Title = Trim(Asp.CheckUrl(Asp.CheckStr(Request.Form("log_Title"))))
		log_cname = Trim(Asp.CheckUrl(Asp.CheckStr(Request.Form("log_cname"))))
		log_ctype = Trim(Asp.CheckUrl(Asp.CheckStr(Request.Form("log_ctype"))))
		log_weather = Trim(Asp.CheckUrl(Asp.CheckStr(Request.Form("log_weather"))))
		log_Level = Trim(Asp.CheckUrl(Asp.CheckStr(Request.Form("log_Level"))))
		log_comorder = Asp.CheckUrl(Asp.CheckStr(Request.Form("log_comorder")))
		log_DisComment = Asp.CheckUrl(Asp.CheckStr(Request.Form("log_DisComment")))
		log_IsTop = Asp.CheckUrl(Asp.CheckStr(Request.Form("log_IsTop")))
		log_Meta = Asp.CheckUrl(Asp.CheckStr(Request.Form("log_Meta")))
		log_KeyWords = Trim(Asp.CheckUrl(Asp.CheckStr(Request.Form("log_KeyWords"))))
		log_Description = Trim(Asp.CheckUrl(Asp.CheckStr(Request.Form("log_Description"))))
		log_From = Trim(Asp.CheckUrl(Asp.CheckStr(Request.Form("log_From"))))
		log_FromURL = Trim(Asp.CheckUrl(Asp.CheckStr(Request.Form("log_FromURL"))))
		PubTimeType = Trim(Asp.CheckUrl(Asp.CheckStr(Request.Form("PubTimeType"))))
		PubTime = Trim(Asp.CheckUrl(Asp.CheckStr(Request.Form("PubTime"))))
		log_tag = Trim(Asp.CheckUrl(Asp.CheckStr(Request.Form("log_tag"))))
		Message = Trim(Asp.CheckUrl(Asp.CheckStr(Request.Form("Message"))))
		log_IsShow = CBool(Int(Trim(Asp.CheckUrl(Asp.CheckStr(Request.Form("log_IsShow"))))))
		
		If log_editType <> 0 Then
			log_disImg = Trim(Asp.CheckUrl(Asp.CheckStr(Request.Form("log_disImg"))))
			log_DisSM = Trim(Asp.CheckUrl(Asp.CheckStr(Request.Form("log_DisSM"))))
			log_DisURL = Trim(Asp.CheckUrl(Asp.CheckStr(Request.Form("log_DisURL"))))
			log_DisKey = Trim(Asp.CheckUrl(Asp.CheckStr(Request.Form("log_DisKey"))))
		End If
		
		log_IntroC = Asp.CheckUrl(Asp.CheckStr(Request.Form("log_IntroC")))
		log_Intro = Trim(Asp.CheckUrl(Asp.CheckStr(Request.Form("log_Intro"))))
		
		'--------------------------------判断-------------------------------
		If log_IsDraft = "true" Then log_IsDraft = True Else log_IsDraft = False
		If PubTimeType = "com" Then log_PostTime = PubTime Else log_PostTime = Now()
		
		If log_comorder = "1" Then log_comorder = True Else log_comorder = False
		If log_DisComment = "1" Then log_DisComment = True Else log_DisComment = False
		If log_IsTop = "1" Then log_IsTop = True Else log_IsTop = False
		If log_Meta = "1" Then log_Meta = True Else log_Meta = False
		If log_IntroC = "1" Then log_IntroC = True Else log_IntroC = False
		
		If log_editType <> 0 Then
			If log_disImg <> "1" Then log_disImg = "0"
			If log_DisSM <> "1" Then log_DisSM = "0"
			If log_DisURL <> "1" Then log_DisURL = "0"
			If log_DisKey <> "1" Then log_DisKey = "0"
			If log_IntroC Then logIntroCustom = "1" Else logIntroCustom = "0"
			log_ubbFlags = log_DisSM & "0" & log_disImg & log_DisURL & log_DisKey & logIntroCustom
		Else
			log_ubbFlags = "00000" & logIntroCustom
		End If
		'-------------------处理Tags--------------------
		Dim tempTags, getTags, post_taglist, post_tag, tempTags2, post_tag2, logTags
        tempTags = Split(log_tag, ",")
        Set getTags = New Tag
			post_taglist = ""
			'添加新的Tag
			For Each post_tag in tempTags
				tempTags2 = Split(post_tag, " ")
				If UBound(tempTags2)>0 Then
					For Each post_tag2 in tempTags2
						If Len(Trim(post_tag2))>0 Then
							post_taglist = post_taglist & "{" & getTags.insert(Asp.CheckStr(Trim(post_tag2))) & "}"
						End If
					Next
				Else
					If Len(Trim(post_tag))>0 Then
						post_taglist = post_taglist & "{" & getTags.insert(Asp.CheckStr(Trim(post_tag))) & "}"
					End If
				End If
			Next
			logTags = post_taglist
        	Call Cache.TagsCache(2)
        Set getTags = Nothing
		'---------------分割日志--------------------
		If log_IntroC Then ' 如果自定义介绍
			If Int(log_editType) = 1 Then ' 如果是UBB编辑器
				If Asp.IsBlank(log_Intro) Then ' 如果内容为空
					If blog_SplitType Then ' 分割类型
                    	log_Intro = Asp.closeUBB(Asp.SplitLines(Message, blog_introLine))
                	Else
                    	log_Intro = Asp.closeUBB(Asp.CutStr(Message, blog_introChar))
                	End If
				Else
					log_Intro = Asp.closeUBB(log_Intro)
				End If
			Else
				If Asp.IsBlank(log_Intro) Then ' 如果内容为空
					If blog_SplitType Then ' 分割类型
                    	log_Intro = Asp.closeHTML(Asp.SplitLines(Message, blog_introLine))
                	Else
                    	log_Intro = Asp.closeHTML(Asp.CutStr(Message, blog_introChar))
                	End If
				Else
					log_Intro = Asp.closeHTML(log_Intro)
				End If
			End If
		Else
			If Int(log_editType) = 1 Then
                If blog_SplitType Then
                    log_Intro = Asp.closeUBB(Asp.SplitLines(Message, blog_introLine))
                Else
                    log_Intro = Asp.closeUBB(Asp.CutStr(Message, blog_introChar))
                End If
            Else
                log_Intro = Asp.closeHTML(Asp.SplitLines(Message, blog_introLine))
            End If
		End If
		'-----------------关键字描述------------------
		If Not log_Meta Then
			log_Description = Left(Asp.FilterUBBTags(Asp.FilterHtmlTags(Message)),254)
		Else
			If Asp.IsBlank(log_Description) Then
				log_Description = Left(Asp.FilterUBBTags(Asp.FilterHtmlTags(Message)),254)
			Else
				log_Description = Asp.FilterUBBTags(Asp.FilterHtmlTags(log_Description))
			End If
		End If
		
		If Not log_Meta Then
			If Asp.IsBlank(log_KeyWords) Then
				log_KeyWords = Asp.CheckStr(log_Title)
			Else
				log_KeyWords = Replace(Replace(Replace(log_KeyWords, ",", "|"), " ", "|"), "|", ",")
			End If
		End If
		
		'------------------------------------------------------------------
		Article.log_CateID = log_CateID
		Article.log_editType = log_editType
		Article.log_IsDraft = log_IsDraft
		Article.log_Title = log_Title
		Article.log_cname = log_cname
		Article.log_ctype = log_ctype
		Article.log_weather = log_weather
		Article.log_Level = log_Level
		Article.log_comorder = log_comorder
		Article.log_DisComment = log_DisComment
		Article.log_IsTop = log_IsTop
		Article.log_Meta = log_Meta
		Article.log_KeyWords = log_KeyWords
		Article.log_Description = log_Description
		Article.log_From = log_From
		Article.log_FromURL = log_FromURL
		Article.log_PostTime = log_PostTime
		Article.log_tag = logTags
		Article.log_Content = Message
		Article.log_IsShow = log_IsShow
		Article.log_ubbFlags = log_ubbFlags
		Article.log_Intro = log_Intro
		
		Str = Article.Add
		Call Data.ArticleList(2)
		If Str(0) Then
			Call web.default
			RedirectUrl("../pjblog.express/blogpost.asp?action=complete&suc=true&id=" & Str(1))
		Else
			RedirectUrl("../pjblog.express/blogpost.asp?action=complete&suc=false&info=" & Str(1) & "&id=0")
		End If
	End Sub
	
	Public Function edit
	
	End Function
	
	Public Function del
		
	End Function

End Class
%>