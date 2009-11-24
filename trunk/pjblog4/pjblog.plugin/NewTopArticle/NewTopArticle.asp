<%
If Application(Sys.CookieName & "_Article_Edit") Or Application(Sys.CookieName & "_plugin_Setting_NewTopArticle") Then
	Application.Lock()
	Application(Sys.CookieName & "_Article_Edit") = False
	Application(Sys.CookieName & "_plugin_Setting_NewTopArticle") = False
	Application.UnLock()
	NewArticle(2)
End If

Function NewArticle(i)
If Not IsArray(Application(Sys.CookieName & "_NewArticle")) Or Int(i) = 2 Then
	dim TopNum, IsShowHidden, TitleLen, IsShowAuthor, Article(), TmpNum, Articles, ArticleTitle
	
	Dim temps, tempStr, temp1, temp2, temp3
	
	Model.open("NewTopArticle")
	Plus.open("NewTopArticle")
	
	TopNum = Model.getKeyvalue("TopNum")
	IsShowHidden = Model.getKeyValue("IsShowHidden")
	TitleLen = Model.getKeyValue("TitleLen")
	IsShowAuthor = Model.getKeyValue("IsShowAuthor")
	temps = Split(Asp.UnCheckStr(Plus.getSingleTemplate("NewTopArticle")), "|$|")
	temp1 = temps(0)
	temp2 = temps(1)
	temp3 = temps(2)
	
	TmpNum = 0 : ReDim Article(-1)   '下标从0开始, 如果没有记录则下标为-1#
	If IsShowHidden="0" Then    '显示隐藏分类和隐藏日志#
		Set Articles = Conn.Execute("SELECT top "&TopNum&" C.log_ID,C.log_Author,C.log_IsShow,C.log_PostTime,C.log_title,L.cate_ID,L.cate_Secret FROM blog_Content AS C,blog_Category AS L where L.cate_ID=C.log_CateID and L.cate_Secret=false and C.log_IsDraft=false order by log_PostTime Desc")
	Else     '不显示隐藏分类和隐藏日志#
		Set Articles = Conn.Execute("SELECT top "&TopNum&" C.log_ID,C.log_Author,C.log_IsShow,C.log_PostTime,C.log_title,L.cate_ID,L.cate_Secret FROM blog_Content AS C,blog_Category AS L where L.cate_ID=C.log_CateID and L.cate_Secret=false and C.log_IsDraft=false and L.cate_Secret=false and C.log_IsShow=false order by log_PostTime Desc")
	End If
	
	Do While Not Articles.Eof
		redim preserve Article(TmpNum)
		If Articles("cate_Secret") then
			ArticleTitle = "[隐藏分类日志]"
		ElseIf Articles("log_IsShow")=false then
			ArticleTitle = "[隐藏日志]"
		Else
			If IsShowAuthor = "0" Then    '显示作者#
				ArticleTitle = Asp.CutStr("<b>"&Articles("log_Author")&": </b>"&Articles("log_title"), TitleLen)
			Else
				ArticleTitle = Asp.CutStr(Articles("log_title"), TitleLen)
			End If
		End If
		tempStr = temp2
		tempStr = Replace(tempStr, "<#id#>", Init.doArticleUrl(Articles("log_id")))
		tempStr = Replace(tempStr, "<#Author#>", Articles("log_Author"))
		tempStr = Replace(tempStr, "<#Time#>", Articles("log_postTime"))
		tempStr = Replace(tempStr, "<#Title#>", ArticleTitle)
		tempStr = temp1 & tempStr & temp3
		Article(TmpNum)=tempStr
		Articles.MoveNext
		TmpNum = TmpNum+1
	Loop	
	
	dim TmpStr : TmpStr = ""
	For TmpNum = 0 to Ubound(Article)
		TmpStr = TmpStr & Article(TmpNum)
	Next
	NewArticle = TmpStr
	Application.Lock()
	Application(Sys.CookieName & "_NewArticle") = NewArticle
	Application.UnLock()
Else
	NewArticle = Application(Sys.CookieName & "_NewArticle")
End If
End Function
%>
	PluginTempValue = ['plugin_NewTopArticle', '<%=outputSideBar(NewArticle(1))%>'];
	PluginOutPutString.push(PluginTempValue);
