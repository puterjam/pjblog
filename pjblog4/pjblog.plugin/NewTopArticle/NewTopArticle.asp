<%
' **********************************
' 最新日志 Recent Article For PJBlog4
' Author: lyowbeen 
' WebSite: http://www.yb13.cn
' **********************************

' --------------------------------
' 判断是否需要更新缓存
' --------------------------------
If Application(Sys.CookieName & "_Article_Edit") Or Application(Sys.CookieName & "_plugin_Setting_NewTopArticle") Then
	Application.Lock()
	Application(Sys.CookieName & "_Article_Edit") = False
	Application(Sys.CookieName & "_plugin_Setting_NewTopArticle") = False
	Application.UnLock()
	NewArticle(2)
End If

Function NewArticle(i)
If IsEmpty(Application(Sys.CookieName & "_NewArticle")) Or Int(i) = 2 Then    '更新缓存#
	
	' ----------------------------
	' 读取插件基本设置信息
	' ----------------------------
	Dim TopNum, IsShowHidden, TitleLen, IsShowAuthor
	Model.open("NewTopArticle")
	TopNum = Model.getKeyvalue("TopNum")
	IsShowHidden = Model.getKeyValue("IsShowHidden")
	TitleLen = Model.getKeyValue("TitleLen")
	IsShowAuthor = Model.getKeyValue("IsShowAuthor")

	' ----------------------------
	' 从数据库读取内容
	' ----------------------------	
	Dim Article(), TmpNum, Articles, ArticleTitle
	TmpNum = 0 : ReDim Article(-1)   '下标从0开始, 如果没有记录则下标为-1#
	If IsShowHidden="0" Then    '显示隐藏日志#
		Set Articles = Conn.Execute("SELECT top "&TopNum&" C.log_ID,C.log_Author,C.log_IsShow,C.log_PostTime,C.log_title,L.cate_ID,L.cate_Secret FROM blog_Content AS C,blog_Category AS L where L.cate_ID=C.log_CateID and L.cate_Secret=false and C.log_IsDraft=false order by log_PostTime Desc")
	Else     '不显示隐藏日志#
		Set Articles = Conn.Execute("SELECT top "&TopNum&" C.log_ID,C.log_Author,C.log_IsShow,C.log_PostTime,C.log_title,L.cate_ID,L.cate_Secret FROM blog_Content AS C,blog_Category AS L where L.cate_ID=C.log_CateID and L.cate_Secret=false and C.log_IsDraft=false and C.log_IsShow=false order by log_PostTime Desc")
	End If

	' ----------------------------
	' 从缓存读取模板内容
	' ----------------------------	
	Dim Template, TmpStr			
	Plus.open("NewTopArticle")
	Template = Split(Asp.UnCheckStr(Plus.getSingleTemplate("NewTopArticle")), "|$|")
			
	Do While Not Articles.Eof
		ReDim preserve Article(TmpNum)
		If Articles("log_IsShow")=false then
			ArticleTitle = "[隐藏日志]"
		Else
			If IsShowAuthor = "0" Then    '显示作者#
				ArticleTitle = Asp.CutStr("<b>"&Articles("log_Author")&": </b>"&Articles("log_title"), TitleLen)
			Else
				ArticleTitle = Asp.CutStr(Articles("log_title"), TitleLen)
			End If
		End If
		
		' ----------------------------
		' 替换模板内容
		' ----------------------------		
		TmpStr = Template(1)
		TmpStr = Replace(TmpStr, "<#id#>", Init.doArticleUrl(Articles("log_id")))
		TmpStr = Replace(TmpStr, "<#Author#>", Articles("log_Author"))
		TmpStr = Replace(TmpStr, "<#Time#>", Articles("log_postTime"))
		TmpStr = Replace(TmpStr, "<#Title#>", ArticleTitle)
		Article(TmpNum)=TmpStr
		Articles.MoveNext
		TmpNum = TmpNum+1
	Loop
	
	' ----------------------------
	' 得到最新文章并写入缓存
	' ----------------------------		
	TmpStr = ""
	For TmpNum = 0 to Ubound(Article)
		TmpStr = TmpStr & Article(TmpNum)
	Next
	NewArticle = Template(0) & TmpStr & Template(2)
	Application.Lock()
	Application(Sys.CookieName & "_NewArticle") = NewArticle
	Application.UnLock()
Else      '从缓存中读取#
	NewArticle = Application(Sys.CookieName & "_NewArticle")
End If
End Function
%>
<!--以下代码用于输出插件内容到页面<plugin:NewTopArticle/>-->
	PluginTempValue = ['NewTopArticle', '<%=outputSideBar(NewArticle(1))%>'];
	PluginOutPutString.push(PluginTempValue);