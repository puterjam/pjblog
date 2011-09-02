<%
'***************PJblog2 缓存处理*******************
' PJblog2 Copyright 2006
' Update:2008-8-26
'**************************************************

'-------------------------Blog基本参数--------------------------
Dim blog_Infos, SiteName, SiteUrl, blogPerPage, blog_LogNums, blog_CommNums, blog_MemNums
Dim blog_VisitNums, blogBookPage, blog_MessageNums, blogcommpage, blogaffiche
Dim blogabout, blogcolsize, blog_colNums, blog_TbCount, blog_showtotal, blog_commTimerout
Dim blog_commUBB, blog_commImg, blog_version, blog_UpdateDate, blog_DefaultSkin, blog_SkinName, blog_SplitType
Dim blog_ImgLink, blog_postFile, blog_postCalendar, log_SplitType, blog_introChar, blog_introLine
Dim blog_validate, Register_UserNames, Register_UserName, FilterIPs, FilterIP, blog_Title, blog_KeyWords, blog_Description, blog_SaveTime, blog_PasswordProtection, blog_AuditOpen, blog_GravatarOpen, blog_html
Dim blog_commLength, blog_downLocal, blog_DisMod, blog_Disregister, blog_master, blog_email, blog_CountNum
Dim blog_wapNum, blog_wapImg, blog_wapHTML, blog_wapLogin, blog_wapComment, blog_wap, blog_wapURL, blog_currentCategoryID
Dim memoryCache, blog_UpLoadSet
Dim blog_smtp, blog_smtpuser, blog_smtppassword, blog_jmail, IsObj, msg, objMail, VerObj, TestObj, blog_smtpmail, blog_Isjmail, blog_reply_Isjmail, blog_IsPing

'一些初始化的值
blog_version = "3.2.9.506" '当前PJBlog版本号
blog_UpdateDate = "2011-09-01" 'PJBlog最新更新时间
memoryCache = false '全内存cache

'=========================日志基本信息缓存=======================

Sub getInfo(ByVal action)
    Dim blog_Infos
    '--------------写入基本信息缓存------------------
    If Not IsArray(Application(CookieName&"_blog_Infos")) Or action = 2 Then
        Dim log_Infos
        SQL = "select top 1 blog_Name,blog_URL,blog_PerPage,blog_LogNums,blog_CommNums,blog_MemNums," & _
              "blog_VisitNums,blog_BookPage,blog_MessageNums,blog_commPage,blog_affiche," & _
              "blog_about,blog_colPage,blog_colNums,blog_tbNums,blog_showtotal," & _
              "blog_FilterName,blog_FilterIP,blog_commTimerout,blog_commUBB,blog_commImg," & _
              "blog_postFile,blog_postCalendar,blog_DefaultSkin,blog_SkinName,blog_SplitType," & _
              "blog_introChar,blog_introLine,blog_validate,blog_Title,blog_ImgLink," & _
              "blog_commLength,blog_downLocal,blog_DisMod,blog_Disregister,blog_master,blog_email,blog_CountNum," & _
			  "blog_wapNum,blog_wapImg,blog_wapHTML,blog_wapLogin,blog_wapComment,blog_wap,blog_wapURL,blog_KeyWords,blog_Description," & _
			  "blog_SaveTime,blog_UpLoadSet,blog_PasswordProtection,blog_AuditOpen,blog_GravatarOpen," & _
			  "blog_smtp,blog_smtpuser,blog_smtppassword,blog_jmail,blog_smtpmail,blog_Isjmail,blog_reply_Isjmail,blog_html,blog_IsPing" & _
              " from blog_Info"
        Set log_Infos = Conn.Execute(SQL)
        SQLQueryNums = SQLQueryNums + 1
        blog_Infos = log_Infos.GetRows()
        Set log_Infos = Nothing
        Application.Lock
        Application(CookieName&"_blog_Infos") = blog_Infos
        Application.UnLock
    Else
        blog_Infos = Application(CookieName&"_blog_Infos")
    End If

    '--------------读取基本信息缓存------------------
    If action<>2 Then
        SiteName = blog_Infos(0, 0)'站点名字
        SiteURL = blog_Infos(1, 0)'站点地址
        blogPerPage = Int(blog_Infos(2, 0))'每页日志数
        blog_LogNums = Int(blog_Infos(3, 0))'日志总数
        blog_CommNums = Int(blog_Infos(4, 0))'评论总数
        blog_MemNums = Int(blog_Infos(5, 0))'会员总数
        blog_VisitNums = Int(blog_Infos(6, 0))'访问量
        blogBookPage = Int(blog_Infos(7, 0))'每页留言数(备用)
        blog_MessageNums = Int(blog_Infos(8, 0))'留言总数(备用)
        blogcommpage = Int(blog_Infos(9, 0))'每页评论数
        blogaffiche = blog_Infos(10, 0)'公告
        blogabout = blog_Infos(11, 0)'备案信息
        blogcolsize = Int(blog_Infos(12, 0))'每页书签数(备用)
        blog_colNums = Int(blog_Infos(13, 0))'书签总数(备用)
        blog_TbCount = Int(blog_Infos(14, 0))'引用通告总数
        blog_showtotal = CBool(blog_Infos(15, 0))'是否显示统计(备用)
        Register_UserNames = blog_Infos(16, 0)'注册名字过滤
        Register_UserName = Split(Register_UserNames, "|")
        FilterIPs = blog_Infos(17, 0)'IP地址过滤
        FilterIP = Split(FilterIPs, "|")
        blog_commTimerout = Int(blog_Infos(18, 0))'发表评论时间间隔
        blog_commUBB = Int(blog_Infos(19, 0))'是否禁用评论UBB代码
        blog_commIMG = Int(blog_Infos(20, 0))'是否禁用评论贴图
        blog_postFile = blog_Infos(21, 0) '动态输出日志文件
        blog_postCalendar = CBool(blog_Infos(22, 0)) '动态输出日志日历文件
        blog_DefaultSkin = blog_Infos(23, 0)'默认界面
        blog_SkinName = blog_Infos(24, 0)'界面名称
        blog_SplitType = CBool(blog_Infos(25, 0))'日志分割类型
        blog_introChar = blog_Infos(26, 0)'日志预览最大字符数
        blog_introLine = blog_Infos(27, 0)'日志预览切割行数
        blog_validate = CBool(blog_Infos(28, 0))'发表评论是否都需要验证
        blog_Title = blog_Infos(29, 0)'Blog副标题
        blog_ImgLink = CBool(blog_Infos(30, 0))'是否在首页显示图片友情链接
        blog_commLength = Int(blog_Infos(31, 0))'评论长度
        blog_downLocal = CBool(blog_Infos(32, 0))'是否使用防盗链下载
        blog_DisMod = CBool(blog_Infos(33, 0))'默认显示内容
        blog_Disregister = CBool(blog_Infos(34, 0))'是否允许注册
        blog_master = blog_Infos(35, 0)'博主姓名
        blog_email = blog_Infos(36, 0)'博主邮件地址
        blog_CountNum = blog_Infos(37, 0)'访客统计最大次数
        blog_wapNum = Int(blog_Infos(38, 0))'Wap 文章列表数量
        blog_wapImg = CBool(blog_Infos(39, 0))'Wap 文章显示图片
        blog_wapHTML = CBool(blog_Infos(40, 0))'Wap 文章使用简单HTML
        blog_wapLogin = CBool(blog_Infos(41, 0))'Wap 允许登录
        blog_wapComment = CBool(blog_Infos(42, 0))'Wap 允许评论
        blog_wap = CBool(blog_Infos(43, 0))'使用 wap
        blog_wapURL = CBool(blog_Infos(44, 0))'使用 wap 转换文章超链接
        blog_KeyWords = blog_Infos(45, 0)'站点首页KeyWords
        blog_Description = blog_Infos(46, 0)'站点首页Description
        blog_SaveTime = blog_Infos(47, 0)'Ajax草稿自动保存时间间隔
        blog_UpLoadSet = blog_Infos(48, 0)'附件管理
		blog_PasswordProtection = CBool(blog_Infos(49, 0))'找回密码功能
		blog_AuditOpen = CBool(blog_Infos(50, 0))'是否开启评论审核功能
		blog_GravatarOpen = CBool(blog_Infos(51, 0)) ' 是否开启Gravatar头像功能
        blog_smtp=trim(blog_Infos(52, 0))'SMTP地址
        blog_smtpuser=trim(blog_Infos(53, 0))'SMTP用户名
        blog_smtppassword=trim(blog_Infos(54, 0))'SMTP密码
        blog_jmail=trim(blog_Infos(55, 0))'发送邮件组件
        blog_smtpmail=trim(blog_Infos(56, 0))'发送邮箱
        blog_Isjmail=blog_Infos(57, 0)'是否邮件通知
		blog_reply_Isjmail=blog_Infos(58, 0)'回复是否邮件通知
		blog_html = blog_Infos(59, 0)'文章后缀列表
		blog_IsPing = CBool(blog_Infos(60, 0)) ' 是否开启Ping功能
    End If
End Sub

'======================End Sub=======================


'-------------------------Blog权限变量---------------
Dim stat_title, stat_AddAll, stat_EditAll, stat_DelAll, stat_Add, stat_Edit, stat_Del, stat_CommentAdd
Dim stat_CommentDel, stat_Admin, stat_code, UP_FileType, UP_FileSize, UP_FileTypes, stat_FileUpLoad
Dim stat_CommentEdit, stat_ShowHiddenCate

'=====================日志权限缓存===================

Sub UserRight(ByVal action) '读取日志权限
    Dim blog_Status
    '--------------写入日志权限缓存------------------
    If Not IsArray(Application(CookieName&"_blog_rights")) Or action = 2 Then
        Dim log_Status, log_StatusList
        SQL = "select stat_name,stat_title,stat_Code,stat_attSize,stat_attType from blog_status"
        Set log_Status = Conn.Execute(SQL)
        SQLQueryNums = SQLQueryNums + 1
        blog_Status = log_Status.GetRows()
        Set log_Status = Nothing

        Application.Lock
        Application(CookieName&"_blog_rights") = blog_Status
        Application.UnLock
    Else
        blog_Status = Application(CookieName&"_blog_rights")
    End If

    '--------------写入日志权限缓存------------------
    If action<>2 Then
        Dim blog_Status_Len, i
        blog_Status_Len = UBound(blog_Status, 2)
        For i = 0 To blog_Status_Len
            If blog_Status(0, i) = memStatus Then
                stat_title = blog_Status(1, i)
                FillRight blog_Status(2, i)
                UP_FileSize = blog_Status(3, i)
                UP_FileTypes = blog_Status(4, i)
                UP_FileType = Split(UP_FileTypes, "|")
                'exit Sub
            End If
        Next
    End If
End Sub

Sub FillRight(StatusCode) '写入权限变量
    stat_AddAll = CBool(Mid(StatusCode, 1, 1))
    stat_Add = CBool(Mid(StatusCode, 2, 1))
    stat_EditAll = CBool(Mid(StatusCode, 3, 1))
    stat_Edit = CBool(Mid(StatusCode, 4, 1))
    stat_DelAll = CBool(Mid(StatusCode, 5, 1))
    stat_Del = CBool(Mid(StatusCode, 6, 1))
    stat_CommentAdd = CBool(Mid(StatusCode, 7, 1))
    stat_CommentEdit = CBool(Mid(StatusCode, 8, 1))
    stat_CommentDel = CBool(Mid(StatusCode, 9, 1))
    stat_FileUpLoad = CBool(Mid(StatusCode, 10, 1))
    stat_Admin = CBool(Mid(StatusCode, 11, 1))
    stat_ShowHiddenCate = CBool(Mid(StatusCode, 12, 1))
    
    Response.Cookies(CookieName)("memRight") = StatusCode
	If DateDiff("d",Date(),Request.Cookies(CookieName)("exp"))>0 Then
        Response.Cookies(CookieName).Expires = Date + DateDiff("d",Date(),Request.Cookies(CookieName)("exp"))
    End If
End Sub

'=========================End Sub========================



'========================日志分类缓存=========================
Function CategoryList(ByVal action) '日志分类
    '写入日志分类
    'action=0 横向菜单 action=1 树状菜单 action=2重建分类, 默认尝试返回Arr_Category

    '--------------写入日志分类缓存------------------
    Dim Arr_Category, i
    If Not IsArray(Application(CookieName&"_blog_Category")) Or action = 2 Then
        Dim log_Category
        TempVar = ""
        SQL = "SELECT cate_ID,cate_Name,cate_Order,cate_Intro,cate_OutLink,cate_URL,cate_icon,cate_count,cate_Lock,cate_local,cate_Secret,cate_Part FROM blog_Category ORDER BY cate_Order ASC"
        Set log_Category = Conn.Execute(SQL)
        SQLQueryNums = SQLQueryNums + 1
        If log_Category.EOF Or log_Category.bof Then
            ReDim Arr_Category(0, 0)
        Else
            Arr_Category = log_Category.GetRows()
        End If
        Set log_Category = Nothing
        Application.Lock
        Application(CookieName&"_blog_Category") = Arr_Category
        Application.UnLock
    Else
        Arr_Category = Application(CookieName&"_blog_Category")
    End If

	CategoryList = "" '初始化
	
    Dim Category_Len, Menu_Diver
    '--------------输出日志横向菜单------------------
    If action = 0 Then
       Menu_Diver = ""
       CategoryList = "<div id=""menu""><div id=""Left""></div><div id=""Right""></div><ul><li class=""menuL""></li>"

        If UBound(Arr_Category, 1) = 0 Then
            CategoryList = CategoryList&"<li class=""menuR""></li></ul></div>"
            Exit Function
        End If

        Category_Len = UBound(Arr_Category, 2)
		
		Dim cClass,URL,inMod
		URL = request.ServerVariables("script_name")
		inMod = inStrRev(URL,"LoadMod.asp",-1,1)
		
        For i = 0 To Category_Len
            If Int(Arr_Category(9, i)) = 0 Or Int(Arr_Category(9, i)) = 1 Then
                CategoryList = CategoryList&Menu_Diver
                
                If Arr_Category(4, i) Then '自定义链接
                    If CBool(Arr_Category(10, i)) Then
                        If stat_ShowHiddenCate Or stat_Admin Then CategoryList = CategoryList&"<li><a class=""menuA"" href="""&Arr_Category(5, i)&""" title="""&Arr_Category(3, i)&""">"&Arr_Category(1, i)&"</a></li>"
                    Else
                    	cClass = "menuA"
                    	if isEmpty(blog_currentCategoryID) and InStrRev(URL,Arr_Category(5,i),-1,1) then 
                    		cClass = "menuA menuB"
                    	end if
                    	
                    	if inMod then
                    			if inStrRev(Arr_Category(5,i),Request.QueryString("plugins"),-1,1) then
                    				cClass = "menuA menuB"
                    			end if
                    	end if
                        CategoryList = CategoryList&"<li><a class="""&cClass&""" href="""&Arr_Category(5, i)&""" title="""&Arr_Category(3, i)&""">"&Arr_Category(1, i)&"</a></li>"
                    End If
                Else'日志分类
                    If CBool(Arr_Category(10, i)) Then
                        If stat_ShowHiddenCate Or stat_Admin Then CategoryList = CategoryList&"<li><a class=""menuA"" href=""default.asp?cateID="&Arr_Category(0, i)&""" title="""&Arr_Category(3, i)&""">"&Arr_Category(1, i)&"</a></li>"
                    Else
                    	cClass = "menuA"
                    	if Int(blog_currentCategoryID) = Arr_Category(0,i) then 
                    		cClass = "menuA menuB"
                    	end if
                        CategoryList = CategoryList&"<li><a class="""&cClass&""" href=""default.asp?cateID="&Arr_Category(0, i)&""" title="""&Arr_Category(3, i)&""">"&Arr_Category(1, i)&"</a></li>"
                    End If
                End If
                Menu_Diver = "<li class=""menuDiv""></li>"
            End If
        Next

        CategoryList = CategoryList&"<li class=""menuR""></li></ul></div>"
    End If

    If action = 1 Then
        CategoryList = ""
        If UBound(Arr_Category, 1) = 0 Then Exit Function

        Category_Len = UBound(Arr_Category, 2)

        For i = 0 To Category_Len
            If Int(Arr_Category(9, i)) = 0 Or Int(Arr_Category(9, i)) = 2 Then
                If Arr_Category(4, i) Then
                    If CBool(Arr_Category(10, i)) Then
                        If stat_ShowHiddenCate Or stat_Admin Then CategoryList = CategoryList&("<img src="""&Arr_Category(6, i)&""" border=""0"" style=""margin:3px 4px -4px 0px;"" alt="""&Arr_Category(3, i)&"""/><a class=""CategoryA"" href="""&Arr_Category(5, i)&""" title="""&Arr_Category(3, i)&""">"&Arr_Category(1, i)&"</a><br/>")
                    Else
                        CategoryList = CategoryList&("<img src="""&Arr_Category(6, i)&""" border=""0"" style=""margin:3px 4px -4px 0px;"" alt="""&Arr_Category(3, i)&"""/><a class=""CategoryA"" href="""&Arr_Category(5, i)&""" title="""&Arr_Category(3, i)&""">"&Arr_Category(1, i)&"</a><br/>")
                    End If
                Else
                    If CBool(Arr_Category(10, i)) Then
                        If stat_ShowHiddenCate Or stat_Admin Then CategoryList = CategoryList&("<img src="""&Arr_Category(6, i)&""" border=""0"" style=""margin:3px 4px -4px 0px;"" alt="""&Arr_Category(3, i)&"""/><a class=""CategoryA"" href=""default.asp?cateID="&Arr_Category(0, i)&""" title="""&Arr_Category(3, i)&""">"&Arr_Category(1, i)&" ["&Arr_Category(7, i)&"]</a> <a href=""feed.asp?cateID="&Arr_Category(0, i)&""" title=""订阅该分类内容""><img src=""images/rss.png"" border=""0"" style=""margin:3px 4px -1px 0px;"" alt=""""/></a><br/>")
                    Else
                        CategoryList = CategoryList&("<img src="""&Arr_Category(6, i)&""" border=""0"" style=""margin:3px 4px -4px 0px;"" alt="""&Arr_Category(3, i)&"""/><a class=""CategoryA"" href=""default.asp?cateID="&Arr_Category(0, i)&""" title="""&Arr_Category(3, i)&""">"&Arr_Category(1, i)&" ["&Arr_Category(7, i)&"]</a> <a href=""feed.asp?cateID="&Arr_Category(0, i)&""" title=""订阅该分类内容""><img src=""images/rss.png"" border=""0"" style=""margin:3px 4px -1px 0px;"" alt=""""/></a><br/>")
                    End If
                End If
            End If
        Next
    End If
End Function

'========================End Sub===============================


'========================日志归档缓存============================

Function Archive(ByVal action)'日志归档
    Dim blog_archive, i
    '-----------------写入日志归档缓存--------------------
    If Not IsArray(Application(CookieName&"_blog_archive")) Or action = 2 Then
        Dim log_archives
        SQL = "SELECT Count(log_ID) AS [count], Year([log_PostTime]) AS PostYear, Month([log_PostTime]) AS PostMonth " &_
        "FROM blog_Content where blog_Content.log_IsDraft=false "&_
        "GROUP BY Year([log_PostTime]), Month([log_PostTime]) "&_
        "ORDER BY Year([log_PostTime]) Desc, Month([log_PostTime]) ASC"
        Set log_archives = Conn.Execute(SQL)
        SQLQueryNums = SQLQueryNums + 1
        If log_archives.EOF Or log_archives.bof Then
            ReDim blog_archive(0, 0)
        Else
            blog_archive = log_archives.GetRows()
        End If
        Set log_archives = Nothing

        Application.Lock
        Application(CookieName&"_blog_archive") = blog_archive
        Application.UnLock
    Else
        blog_archive = Application(CookieName&"_blog_archive")
    End If

    '-----------------读取日志归档缓存--------------------
    If action<>2 Then
        Dim archive_item_Len, Month_array, TempYear, MonthCounter
        If UBound(blog_archive, 1) = 0 Then
            Archive = ""
            Exit Function
        End If
        Month_array = Array("01月", "02月", "03月", "04月", "05月", "06月", "07月", "08月", "09月", "10月", "11月", "12月")
        archive_item_Len = UBound(blog_archive, 2)
        TempYear = blog_archive(1, 0)
        MonthCounter = 0
        For i = 0 To archive_item_Len
            If i = 0 Then Archive = "<a class=""sideA"" style=""margin:0px 0px 0px -2px;"" href=""default.asp?log_Year="&blog_archive(1, i)&""" title=""查看"&blog_archive(1, i)&"年的日志"">"&blog_archive(1, i)&"</a>"
            If blog_archive(1, i) = TempYear Then
                Archive = Archive&"<a style=""margin-right:5px;"" href=""default.asp?log_Year="&blog_archive(1, i)&"&log_Month="&blog_archive(2, i)&""" title="""&blog_archive(1, i)&"年"&blog_archive(2, i)&"月有"&blog_archive(0, i)&"篇日志"">"&Month_array(blog_archive(2, i) -1)&"</a>"
                MonthCounter = MonthCounter + 1
                If MonthCounter = 5 Then
                    MonthCounter = 0
                    Archive = Archive&"<br/>"
                End If
            Else
                MonthCounter = 1
                Archive = Archive&"<a class=""sideA"" style=""margin:6px 0px 0px -2px;"" href=""default.asp?log_Year="&blog_archive(1, i)&""" title=""查看"&blog_archive(1, i)&"年的日志"">"&blog_archive(1, i)&"</a>"
                Archive = Archive&"<a style=""margin-right:5px;"" href=""default.asp?log_Year="&blog_archive(1, i)&"&log_Month="&blog_archive(2, i)&""" title="""&blog_archive(1, i)&"年"&blog_archive(2, i)&"月有"&blog_archive(0, i)&"篇日志"">"&Month_array(blog_archive(2, i) -1)&"</a>"
                TempYear = blog_archive(1, i)
            End If
        Next
    End If
End Function

'=====================End Function========================



'=====================最新评论缓存=====================

Function NewComment(ByVal action)
    Dim blog_Comment, ShowLen, i
    ShowLen = 10 '显示最新评论预览数量
    '-----------------写入最新评论缓存--------------------
    If Not IsArray(Application(CookieName&"_blog_Comment")) Or action = 2 Then
        Dim log_Comments
        SQL = "SELECT top "&ShowLen&" comm_ID,blog_ID,comm_Author,comm_Content,comm_PostTime,comm_IsAudit" &_
        " FROM blog_Comment as C,blog_Content as T,blog_Category as A where C.blog_ID=T.log_ID and T.log_IsShow=true and T.log_CateID=A.cate_ID and A.cate_Secret=false order by C.comm_PostTime Desc"
        Set log_Comments = Conn.Execute(SQL)
        SQLQueryNums = SQLQueryNums + 1
        If log_Comments.EOF Or log_Comments.bof Then
            ReDim blog_Comment(0, 0)
        Else
            blog_Comment = log_Comments.GetRows(ShowLen)
        End If
        Set log_Comments = Nothing
        Application.Lock
        Application(CookieName&"_blog_Comment") = blog_Comment
        Application.UnLock
    Else
        blog_Comment = Application(CookieName&"_blog_Comment")
    End If

    '-----------------读取最新评论缓存--------------------
    If action<>2 Then
        Dim Comment_Item_Len, NewTitleInfo
        If UBound(blog_Comment, 1) = 0 Then
            NewComment = ""
            Exit Function
        End If
        Comment_Item_Len = UBound(blog_Comment, 2)
        dim url
        For i = 0 To Comment_Item_Len
		    If blog_postFile = 2 Then
		   		url = caload(blog_Comment(1, i))&"#comm_"&blog_Comment(0, i)
		      else
		   		url = "article.asp?id="&blog_Comment(1, i)&"#comm_"&blog_Comment(0, i)
		    end if
			If blog_AuditOpen Then
				If blog_Comment(5, i) Then
					NewTitleInfo = blog_Comment(2, i)&" 于 "&DateToStr(blog_Comment(4, i),"Y-m-d H:I A")&" 发表评论"&Chr(10)&CCEncode(CutStr(DelQuote(blog_Comment(3, i)), 100))
				Else
					NewTitleInfo = "[此评论正在审核中...]"
				End If
			Else
				NewTitleInfo = blog_Comment(2, i)&" 于 "&DateToStr(blog_Comment(4, i),"Y-m-d H:I A")&" 发表评论"&Chr(10)&CCEncode(CutStr(DelQuote(blog_Comment(3, i)), 100))
			End If
            NewComment = NewComment&"<a class=""sideA"" href="""&url&""" title="""&NewTitleInfo&""">"
			If blog_AuditOpen Then
				If blog_Comment(5, i) Then
					NewComment = NewComment&CCEncode(CutStr(DelQuote(blog_Comment(3, i)), 25))
				Else
					NewComment = NewComment&"[未审核评论]"
				End If
			Else
				NewComment = NewComment&CCEncode(CutStr(DelQuote(blog_Comment(3, i)), 25))
			End If
			NewComment = NewComment&"</a>"
      
        Next
    End If
End Function

'=====================End Function========================

'====================写入标签Tag缓存=====================
Dim Arr_Tags

Function Tags(ByVal action)
    If Not IsArray(Application(CookieName&"_blog_Tags")) Or action = 2 Then
        Dim log_Tags, log_TagsList
        Set log_Tags = Conn.Execute("SELECT tag_id,tag_name,tag_count FROM blog_tag")
        SQLQueryNums = SQLQueryNums + 1
        TempVar = ""
        Do While Not log_Tags.EOF
            log_TagsList = log_TagsList&TempVar&log_Tags("tag_id")&"||"&log_Tags("tag_name")&"||"&log_Tags("tag_count")
            TempVar = ","
            log_Tags.MoveNext
        Loop
        Set log_Tags = Nothing
        Arr_Tags = Split(log_TagsList, ",")
        Application.Lock
        Application(CookieName&"_blog_Tags") = Arr_Tags
        Application.UnLock
    Else
        Arr_Tags = Application(CookieName&"_blog_Tags")
    End If
End Function

'======================End Function========================

'====================写入表情符号缓存=====================
Dim Arr_Smilies

Function Smilies(ByVal action)
    If Not IsArray(Application(CookieName&"_blog_Smilies")) Or action = 2 Then
        Dim log_Smilies, log_SmiliesList
        Set log_Smilies = Conn.Execute("SELECT sm_ID,sm_Image,sm_Text FROM blog_Smilies")
        SQLQueryNums = SQLQueryNums + 1
        TempVar = ""
        Do While Not log_Smilies.EOF
            log_SmiliesList = log_SmiliesList&TempVar&log_Smilies("sm_ID")&"|"&log_Smilies("sm_Image")&"|"&log_Smilies("sm_Text")
            TempVar = ","
            log_Smilies.MoveNext
        Loop
        Set log_Smilies = Nothing
        Arr_Smilies = Split(log_SmiliesList, ",")
        Application.Lock
        Application(CookieName&"_blog_Smilies") = Arr_Smilies
        Application.UnLock
    Else
        Arr_Smilies = Application(CookieName&"_blog_Smilies")
    End If
End Function

'======================End Function========================

'======================写入关键字列表======================
Dim Arr_Keywords

Function Keywords(ByVal action)
    If Not IsArray(Application(CookieName&"_blog_Keywords")) Or action = 2 Then
        Dim log_Keywords, log_KeywordsList
        Set log_Keywords = Conn.Execute("SELECT key_ID,key_Text,key_URL,key_Image FROM blog_Keywords")
        SQLQueryNums = SQLQueryNums + 1
        TempVar = ""
        Do While Not log_Keywords.EOF
            If log_Keywords("key_Image")<>Empty Then
                log_KeywordsList = log_KeywordsList&TempVar&log_Keywords("key_ID")&"$|$"&log_Keywords("key_Text")&"$|$"&log_Keywords("key_URL")&"$|$"&log_Keywords("key_Image")
            Else
                log_KeywordsList = log_KeywordsList&TempVar&log_Keywords("key_ID")&"$|$"&log_Keywords("key_Text")&"$|$"&log_Keywords("key_URL")&"$|$None"
            End If
            TempVar = "|$|"
            log_Keywords.MoveNext
        Loop
        Set log_Keywords = Nothing
        Arr_Keywords = Split(log_KeywordsList, "|$|")
        Application.Lock
        Application(CookieName&"_blog_Keywords") = Arr_Keywords
        Application.UnLock
    Else
        Arr_Keywords = Application(CookieName&"_blog_Keywords")
    End If
End Function

'======================End Function=========================

'=======================写入首页链接列表====================
Dim Arr_Bloglinks

Function Bloglinks(ByVal action)
    If Not IsArray(Application(CookieName&"_blog_Bloglinks")) Or action = 2 Then
        Dim log_Bloglinks, log_BloglinksList
        Set log_BlogLinks = Conn.Execute("SELECT link_Name,link_URL,link_Image FROM blog_Links WHERE link_IsMain=True ORDER BY link_Order ASC")
        SQLQueryNums = SQLQueryNums + 1
        TempVar = ""
        Do While Not log_BlogLinks.EOF
            If log_BlogLinks("link_Image")<>Empty Then
                log_BloglinksList = log_BloglinksList&TempVar&log_BlogLinks("link_Name")&"$|$"&log_BlogLinks("link_URL")&"$|$"&log_BlogLinks("link_Image")
            Else
                log_BloglinksList = log_BloglinksList&TempVar&log_BlogLinks("link_Name")&"$|$"&log_BlogLinks("link_URL")&"$|$None"
            End If
            TempVar = "|$|"
            log_BlogLinks.MoveNext
        Loop
        Set log_BlogLinks = Nothing
        Arr_Bloglinks = Split(log_BloglinksList, "|$|")
        Application.Lock
        Application(CookieName&"_blog_Bloglinks") = Arr_Bloglinks
        Application.UnLock
    Else
        Arr_Bloglinks = Application(CookieName&"_blog_Bloglinks")
    End If

    If action = 1 Then
        Dim Arr_Bloglink, Arr_BloglinkItem, ImgLink, TextLink
        Bloglinks = ""
        For Each Arr_Bloglink in Arr_Bloglinks
            Arr_BloglinkItem = Split(Arr_Bloglink, "$|$")
            If blog_ImgLink Then
                If Arr_BloglinkItem(2) = "None" Then
                    TextLink = TextLink&"<a class=""sideA"" href="""&Arr_BloglinkItem(1)&""" target=""_blank"" title="""&Arr_BloglinkItem(0)&""">"&Arr_BloglinkItem(0)&"</a>"
                Else
                    ImgLink = ImgLink&"<a href="""&Arr_BloglinkItem(1)&""" target=""_blank"" title="""&Arr_BloglinkItem(0)&"""><img src="""&Arr_BloglinkItem(2)&""" border=""0"" alt=""""/></a> "
                End If
            Else
                Bloglinks = Bloglinks&"<a class=""sideA"" href="""&Arr_BloglinkItem(1)&""" target=""_blank"" title="""&Arr_BloglinkItem(0)&""">"&Arr_BloglinkItem(0)&"</a>"
            End If
        Next
        If blog_ImgLink Then Bloglinks = ImgLink&TextLink
    End If
End Function

'=====================End Function=======================

Function BlogArticleID(ByVal action)
	If IsEmpty(Application(CookieName & "_blog_ReBuildID")) Or IsNull(Application(CookieName & "_blog_ReBuildID")) Or action = 2 Then
		Dim BlogArticleDB, BlogArticleList
		Set BlogArticleDB = Conn.Execute("SELECT log_ID FROM blog_Content Where log_IsDraft=False ORDER BY log_ID ASC")
		SQLQueryNums = SQLQueryNums + 1
		TempVar = ""
		If BlogArticleDB.Bof Or BlogArticleDB.Eof Then
			TempVar = "[]"
		Else
			TempVar = TempVar & "["
			Do While Not BlogArticleDB.Eof
				TempVar = TempVar & BlogArticleDB(0) & ","
			BlogArticleDB.MoveNext
			Loop
			TempVar = Mid(TempVar, 1, (Len(TempVar) - 1))
			TempVar = TempVar & "]"
		End If
		Application.Lock()
		Application(CookieName & "_blog_ReBuildID") = TempVar
		Application.UnLock()
		BlogArticleID = TempVar
	Else
		BlogArticleID = Application(CookieName & "_blog_ReBuildID")
	End If
End Function

'======================自定义模块缓存=====================
Dim side_html_default, side_html, side_html_static, content_html_Top_default, content_html_Top, content_html_Top_static, content_html_Bottom_default, content_html_Bottom, content_html_Bottom_static, function_Plugin

Function log_module(ByVal action)
    Dim blog_modules
    side_html_default = "" '首页侧栏代码
    side_html = "" '普通页面侧栏代码
    side_html_static = "" '静态页面侧栏代码
    
    content_html_Top_default = "" '首页内容代码顶部
    content_html_Top = "" '普通页面内容代码顶部
    content_html_Top_static = "" '静态页面内容代码顶部
    content_html_Bottom_default = "" '首页内容代码底部
    content_html_Bottom = "" '普通页面内容代码底部
    content_html_Bottom_static = "" '静态页面内容代码底部
    function_Plugin = "" 'Blog功能插件

    If Not IsArray(Application(CookieName&"_blog_module")) Or action = 2 Then
        Dim blog_module, blog_module_array, TempDiv
        TempDiv = ""
        SQL = "SELECT type,title,name,HtmlCode,IndexOnly,SortID,PluginPath,InstallFolder,IsSystem FROM blog_module where IsHidden=false order by SortID"
        Set blog_module = Conn.Execute(SQL)
        SQLQueryNums = SQLQueryNums + 1
        Do Until blog_module.EOF
            If blog_module("type") = "sidebar" Then
                side_html_default = side_html_default&"<div id=""Side_"&blog_module("name")&""" class=""sidepanel"">"
                If Len(blog_module("title"))>0 Then side_html_default = side_html_default&"<h4 class=""Ptitle"">"&blog_module("title")&"</h4>"
                side_html_default = side_html_default&"<div class=""Pcontent"">"&blog_module("HtmlCode")&"</div><div class=""Pfoot""></div></div>"
                
                If blog_module("IndexOnly") = False Then
                    side_html = side_html&"<div id=""Side_"&blog_module("name")&""" class=""sidepanel"">"
                    If Len(blog_module("title"))>0 Then side_html = side_html&"<h4 class=""Ptitle"">"&blog_module("title")&"</h4>"
                    side_html = side_html&"<div class=""Pcontent"">"&blog_module("HtmlCode")&"</div><div class=""Pfoot""></div></div>"
                End If
                
                If blog_module("IndexOnly") = False Then
                    side_html_static = side_html_static&"<div id=""Side_"&blog_module("name")&""" class=""sidepanel"">"
                    If Len(blog_module("title"))>0 Then side_html_static = side_html_static&"<h4 class=""Ptitle"">"&blog_module("title")&"</h4>"
                    side_html_static = side_html_static&"<div class=""Pcontent"">"&blog_module("HtmlCode")&"</div><div class=""Pfoot""></div></div>"
                End If
            End If
            
            If blog_module("type") = "content" And blog_module("name")<>"ContentList" Then
                If blog_module("SortID")<0 Then
                    content_html_Top_default = content_html_Top_default&"<div id=""Content_"&blog_module("name")&""" class=""content-width"">"&blog_module("HtmlCode")&"</div>"
                    If blog_module("IndexOnly") = False Then
                        content_html_Top = content_html_Top&"<div id=""Content_"&blog_module("name")&""" class=""content-width"">"&blog_module("HtmlCode")&"</div>"
                        content_html_Top_static = content_html_Top_static&"<div id=""Content_"&blog_module("name")&""" class=""content-width"">"&blog_module("HtmlCode")&"</div>"
                    End If
                Else
                    content_html_Bottom_default = content_html_Bottom_default&"<div id=""Content_"&blog_module("name")&""" class=""content-width"">"&blog_module("HtmlCode")&"</div>"
                    If blog_module("IndexOnly") = False Then
                        content_html_Bottom = content_html_Bottom&"<div id=""Content_"&blog_module("name")&""" class=""content-width"">"&blog_module("HtmlCode")&"</div>"
                        content_html_Bottom_static = content_html_Bottom_static&"<div id=""Content_"&blog_module("name")&""" class=""content-width"">"&blog_module("HtmlCode")&"</div>"
                    End If
                End If
            End If
            
            If blog_module("type") = "function" Then
                function_Plugin = function_Plugin&TempDiv&blog_module("name")&"%|%"&blog_module("PluginPath")&"%|%"&blog_module("InstallFolder")
                TempDiv = "$*$"
            End If
            
            blog_module.movenext
        Loop
        Set blog_module = Nothing
        blog_modules = Array(side_html_default, side_html, content_html_Top_default, content_html_Top, content_html_Bottom_default, content_html_Bottom, function_Plugin,side_html_static,content_html_Top_static,content_html_Bottom_static)
        Application.Lock
        Application(CookieName&"_blog_module") = blog_modules
        Application.UnLock
    Else
        blog_modules = Application(CookieName&"_blog_module")
    End If

    If action<>2 Then

        side_html_default = UnCheckStr(blog_modules(0)) '首页侧栏代码
        side_html = UnCheckStr(blog_modules(1)) '普通页面侧栏代码
        side_html_static = UnCheckStr(blog_modules(7)) '静态页面侧栏代码
        
        content_html_Top_default = UnCheckStr(blog_modules(2)) '首页内容代码顶部
        content_html_Top = UnCheckStr(blog_modules(3)) '普通页面内容代码顶部
        content_html_Top_static = UnCheckStr(blog_modules(8)) '静态页面内容代码顶部
        content_html_Bottom_default = UnCheckStr(blog_modules(4)) '首页内容代码底部
        content_html_Bottom = UnCheckStr(blog_modules(5)) '普通页面内容代码底部
        content_html_Bottom_static = UnCheckStr(blog_modules(9)) '静态页面内容代码底部
        function_Plugin = blog_modules(6) 'Blog功能插件
    End If

End Function

'========================End function=========================


'======================重新加载Blog缓存=====================

Sub reloadcache
    getInfo(2)
    UserRight(2)
    CategoryList(2)
    Archive(2)
    NewComment(2)
    Tags(2)
    Smilies(2)
    Keywords(2)
	BlogArticleID(2)
    Bloglinks(2)
    log_module(2)
    Calendar "", "", "", 2
End Sub

'=====================304 支持==========================
'更新Etag
Sub newEtag
        Application.Lock
        Application(CookieName&"_Etag") = randomStr(10)
        Application.UnLock
End Sub

'返回服务器的etag，由随机数和登录用户名组成
Function getEtag
 	if IsEmpty(Application(CookieName&"_Etag")) then 
 		call newEtag
 	end if
 	
 	getEtag = Application(CookieName&"_Etag") & "-" & CheckStr(Request.Cookies(CookieName)("memName"))
End Function

'更新Etag,使其支持静态侧边插件
Sub EmptyEtag
	Application(CookieName&"_Etag") = empty
	reloadcache
End Sub
%>
