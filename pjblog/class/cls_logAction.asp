<%
'==================================
'  日志编辑类
'    更新时间: 2006-1-22
'==================================
Class logArticle
    Private weblog
    Public categoryID, logTitle, logAuthor, logEditType
    Public logIsShow, logIsDraft, logWeather, logLevel, logCommentOrder, logReadpw, logPwtips, TitleHidden
    Public logDisableComment, logIsTop, logFrom, logFromURL
    Public logDisableImage, logDisableSmile, logDisableURL, logDisableKeyWord, logKeyWords, logDescription, logmeta
    Public logQuote, logMessage, logIntro, logIntroCustom, logTags, logPublishTimeType, logPubTime, logTrackback, logCommentCount, logQuoteCount, logViewCount, logcname, logctype
    Private logUbbFlags, PubTime, sqlString

    Private Sub Class_Initialize()
        Set weblog = Server.CreateObject("ADODB.RecordSet")
        categoryID = 0
        logTitle = ""
        logEditType = 1
        logIntroCustom = 0
        logIntro = ""
        logAuthor = "null"
        logWeather = "sunny"
        logLevel = "level3"
        logCommentOrder = 1
        logDisableComment = 0
        logIsShow = True
        logIsTop = False
        logIsDraft = False
        logFrom = "本站原创"
        logFromURL = siteURL
        logDisableImage = 0
        logDisableSmile = 0
        logDisableURL = 0
        logDisableKeyWord = 0
        logCommentCount = 0
        logQuoteCount = 0
        logViewCount = 0
        logMessage = ""
        logTrackback = ""
        logTags = ""
        logPubTime = "2006-1-1 00:00:00"
        logPublishTimeType = "now"
        logReadpw = ""
        logPwtips = ""
		TitleHidden = false
		If blog_postFile = 2 Then
		logcname=""
		logctype= "0"
		end if
		logKeyWords = "http://www.evio.name"
		logDescription = "http://www.evio.name"
		logmeta = 0
    End Sub

    Private Sub Class_Terminate()
        Set weblog = Nothing
    End Sub

    '*********************************************
    '发表新日志
    '*********************************************

    Public Function postLog()
        postLog = Array( -4, "准备发表日志", -1)
        weblog.Open "blog_Content", Conn, 1, 2
        SQLQueryNums = SQLQueryNums + 1

        If stat_AddAll<>True And stat_Add<>True Then
            postLog = Array( -3, "没有权限发表日志", -1)
            Exit Function
        End If

        '-------------------处理Tags--------------------
        Dim tempTags,tempTags2, loadTagString, loadTags, loadTag, getTags
        tempTags = Split(CheckStr(logTags), ",")

        Set getTags = New Tag

        Dim post_tag,post_tag2, post_taglist
        post_taglist = ""

        '添加新的Tag
        For Each post_tag in tempTags
      	  	tempTags2 = Split(post_tag," ")
      	  	If UBound(tempTags2)>0 Then
	      	  	For Each post_tag2 in tempTags2
		            If Len(Trim(post_tag2))>0 Then
		                post_taglist = post_taglist & "{" & getTags.insert(CheckStr(Trim(post_tag2))) & "}"
		            End If
	      	  	Next
      	  	Else
            If Len(Trim(post_tag))>0 Then
                post_taglist = post_taglist & "{" & getTags.insert(CheckStr(Trim(post_tag))) & "}"
            End If
            End if
        Next
        logTags = post_taglist
        Call Tags(2)
        Set getTags = Nothing
        '--------------处理日期---------------------
        If CheckStr(logPublishTimeType) = "now" Then
            PubTime = DateToStr(Now(), "Y-m-d H:I:S")
        Else
            PubTime = DateToStr(CheckStr(logPubTime), "Y-m-d H:I:S")
        End If

        '---------------分割日志--------------------
        If logIntroCustom = 1 Then
            If Int(logEditType) = 1 Then
                logIntro = closeUBB(CheckStr(HTMLEncode(logIntro)))
            Else
                logIntro = closeHTML(CheckStr(logIntro))
            End If
        Else
            If Int(logEditType) = 1 Then
                If blog_SplitType Then
                    logIntro = closeUBB(SplitLines(CheckStr(HTMLEncode(logMessage)), blog_introLine))
                Else
                    logIntro = closeUBB(CutStr(CheckStr(HTMLEncode(logMessage)), blog_introChar))
                End If
            Else
                logIntro = closeHTML(SplitLines(CheckStr(logMessage), blog_introLine))
            End If
        End If

        '日志基本状态
       ' If logIsShow = 1 Then logIsShow = 0 Else logIsShow = 1 '私密日志特别属性
        logIsShow = CBool(logIsShow)
        logCommentOrder = CBool(logCommentOrder)
        logDisableComment = CBool(logDisableComment)
        logIsTop = CBool(logIsTop)
        logIsDraft = CBool(logIsDraft)

        'UBB 特别属性
        If logDisableSmile = 1 Then logDisableSmile = 1 Else logDisableSmile = 0
        If logDisableImage = 1 Then logDisableImage = 1 Else logDisableImage = 0
        If logDisableURL = 1 Then logDisableURL = 0 Else logDisableURL = 1
        If logDisableKeyWord = 1 Then logDisableKeyWord = 0 Else logDisableKeyWord = 1
        If logIntroCustom = 1 Then logIntroCustom = 0 Else logIntroCustom = 1
        logUbbFlags = logDisableSmile & "0" & logDisableImage & logDisableURL & logDisableKeyWord & logIntroCustom
		
		'Meta特别属性
		if logmeta = 0 then
			logDescription = logIntro
			logKeyWords = CheckStr(logTags)
			if len(logKeyWords) = 0 then
				logKeyWords = CheckStr(logTitle)
			else
				logKeyWords = Replace(Replace(Replace(logKeyWords, ",", "|"), " ", "|"), "|", ",")
			end if
		end if

        weblog.addNew
		Dim CnameID:CnameID = weblog("log_ID")
        weblog("log_CateID") = CheckStr(categoryID)
        weblog("log_Author") = CheckStr(logAuthor)
        weblog("log_Title") = CheckStr(logTitle)
        weblog("log_weather") = CheckStr(logWeather)
        weblog("log_Level") = CheckStr(logLevel)
        weblog("log_From") = CheckStr(logFrom)
        weblog("log_FromURL") = CheckStr(logFromURL)
        weblog("log_Content") = CheckStr(logMessage)
        weblog("log_Intro") = logIntro
        weblog("log_tag") = logTags
        weblog("log_ubbFlags") = logUbbFlags
        weblog("log_IsShow") = logIsShow
        weblog("log_IsTop") = logIsTop
        weblog("log_PostTime") = PubTime
        weblog("log_IsDraft") = logIsDraft
        weblog("log_DisComment") = logDisableComment
        weblog("log_edittype") = logEditType
        weblog("log_comorder") = logCommentOrder
		
		if len(logcname) < 1 or logcname = "" or logcname = empty or logcname = null then
		logcname = CnameID
		end if
		weblog("log_cname") = logcname
		weblog("log_ctype") = logctype

		if int(TitleHidden) = 1 then
		weblog("Title_Hidden") = true
		else
		weblog("Title_Hidden") = false
		end if
		weblog("log_Readpw") = logReadpw
		weblog("log_Pwtips") = logPwtips
		
		weblog("log_KeyWords") = logKeyWords
		weblog("log_Description") = logDescription
		weblog("log_Meta") = logmeta
		
        SQLQueryNums = SQLQueryNums + 2
        weblog.update
        weblog.Close


        '------------------统计日志-----------------------------
        Dim PostLogID
        PostLogID = Conn.Execute("SELECT TOP 1 log_ID FROM blog_Content ORDER BY log_ID DESC")(0)
        Conn.Execute("UPDATE blog_Member SET mem_PostLogs=mem_PostLogs+1 WHERE mem_Name='"&logAuthor&"'")
        If Not logIsDraft Then
            Conn.Execute("UPDATE blog_Info SET blog_LogNums=blog_LogNums+1")
            Conn.Execute("UPDATE blog_Category SET cate_count=cate_count+1 where cate_ID="&categoryID)
            SQLQueryNums = SQLQueryNums + 2
        End If


        '-------------------输出静态日志档案--------------------
        Dim preLog, nextLog
        '输出日志到文件
        PostArticle PostLogID, False

        '输出附近的日志到文件
        Set preLog = Conn.Execute("SELECT TOP 1 T.log_Title,T.log_ID FROM blog_Content As T,blog_Category As C WHERE T.log_PostTime<#"&PubTime&"# and T.log_CateID=C.cate_ID and (T.log_IsShow=true or T.log_Readpw<>'') and C.cate_Secret=False and T.log_IsDraft=false ORDER BY T.log_PostTime DESC")
        Set nextLog = Conn.Execute("SELECT TOP 1 T.log_Title,T.log_ID FROM blog_Content As T,blog_Category As C WHERE T.log_PostTime>#"&PubTime&"# and T.log_CateID=C.cate_ID and (T.log_IsShow=true or T.log_Readpw<>'') and C.cate_Secret=False and T.log_IsDraft=false ORDER BY T.log_PostTime ASC")
        If Not preLog.EOF Then PostArticle preLog("log_ID"), False
        If Not nextLog.EOF Then PostArticle nextLog("log_ID"), False

        Call updateCache

        Session(CookieName&"_LastDo") = "AddArticle"
        session(CookieName&"_draft_"&logAuthor) = conn.Execute("select count(log_ID) from blog_Content where log_Author='"&logAuthor&"' and log_IsDraft=true")(0)
        SQLQueryNums = SQLQueryNums + 1



        If logIsDraft Then
            postLog = Array(1, "日志成功保存为草稿", PostLogID)
        Else
            postLog = Array(0, "恭喜!日志发表成功", PostLogID)
        End If

        '-------------------引用通告-------------------
        If logTrackback<>Empty And logIsShow = True And logIsDraft = False Then
            Dim log_QuoteEvery, log_QuoteArr, logid, LastID
            Set LastID = Conn.Execute("SELECT TOP 1 log_ID FROM blog_Content ORDER BY log_ID DESC")
            logid = LastID("log_ID")
            log_QuoteArr = Split(logTrackback, ",")
            For Each log_QuoteEvery In log_QuoteArr
                Trackback Trim(log_QuoteEvery), siteURL&"default.asp?id="&logid, logTitle, CutStr(CheckStr(logIntro), 252), siteName
                Set LastID = Nothing
            Next
        End If
    End Function

    '*********************************************
    '编辑日志
    '*********************************************

    Public Function editLog(id)
        editLog = Array( -4, "准备编辑日志", -1)
        If IsEmpty(id) Then
            getLog = Array( -5, "ID号不能为空")
            Exit Function
        End If
        If Not IsInteger(id) Then
            editLog = Array( -1, "非法ID号", -1)
            Exit Function
        End If

        sqlString = "SELECT top 1 * FROM blog_Content WHERE log_ID="&id&""
        weblog.Open sqlString, Conn, 1, 3
        SQLQueryNums = SQLQueryNums + 1

        If weblog.EOF Or weblog.bof Then
            editLog = Array( -2, "无法找到相应文章", -1)
            Exit Function
        End If

        If stat_EditAll<>True And (stat_Edit And weblog("log_Author") = logAuthor)<>True Then
            editLog = Array( -3, "您没有权限编辑日志", -1)
            Exit Function
        End If

        logAuthor = weblog("log_Author")
        Conn.Execute("UPDATE blog_Category SET cate_count=cate_count-1 where cate_ID="&weblog("log_CateID"))
        Conn.Execute("UPDATE blog_Category SET cate_count=cate_count+1 where cate_ID="&CheckStr(categoryID))


        '-------------------处理Tags--------------------
        Dim tempTags,tempTags2, loadTagString, loadTags, loadTag, getTags
        tempTags = Split(CheckStr(logTags), ",")
        loadTagString = weblog("log_tag")

        Set getTags = New Tag

        '清除旧的Tag
        If Len(loadTagString)>0 Then
            loadTagString = Replace(loadTagString, "}{", ",")
            loadTagString = Replace(loadTagString, "}", "")
            loadTagString = Replace(loadTagString, "{", "")
            loadTags = Split(loadTagString, ",")

            For Each loadTag in loadTags
                getTags.Remove loadTag
            Next
        End If

        Dim post_tag,post_tag2, post_taglist
        post_taglist = ""

        '添加新的Tag
        For Each post_tag in tempTags
      	  	tempTags2 = Split(post_tag," ")
      	  	If UBound(tempTags2)>0 Then
	      	  	For Each post_tag2 in tempTags2
		            If Len(Trim(post_tag2))>0 Then
		                post_taglist = post_taglist & "{" & getTags.insert(CheckStr(Trim(post_tag2))) & "}"
		            End If
	      	  	Next
      	  	Else
            If Len(Trim(post_tag))>0 Then
                post_taglist = post_taglist & "{" & getTags.insert(CheckStr(Trim(post_tag))) & "}"
            End If
            End if
        Next
        logTags = post_taglist
        Call Tags(2)
        Set getTags = Nothing
        '--------------处理日期---------------------
        If CheckStr(logPublishTimeType) = "now" Then
            PubTime = DateToStr(Now(), "Y-m-d H:I:S")
        Else
            PubTime = DateToStr(CheckStr(logPubTime), "Y-m-d H:I:S")
        End If

        '---------------分割日志--------------------
        If logIntroCustom = 1 Then
            If Int(logEditType) = 1 Then
                logIntro = closeUBB(CheckStr(HTMLEncode(logIntro)))
            Else
                logIntro = closeHTML(CheckStr(logIntro))
            End If
        Else
            If Int(logEditType) = 1 Then
                If blog_SplitType Then
                    logIntro = closeUBB(SplitLines(CheckStr(HTMLEncode(logMessage)), blog_introLine))
                Else
                    logIntro = closeUBB(CutStr(CheckStr(HTMLEncode(logMessage)), blog_introChar))
                End If
            Else
                logIntro = closeHTML(SplitLines(CheckStr(logMessage), blog_introLine))
            End If
        End If

        '日志基本状态
       ' If logIsShow = 1 Then logIsShow = 0 Else logIsShow = 1 '私密日志特别属性
        logIsShow = CBool(logIsShow)
        logCommentOrder = CBool(logCommentOrder)
        logDisableComment = CBool(logDisableComment)
        logIsTop = CBool(logIsTop)
        logIsDraft = CBool(logIsDraft)

        'UBB 特别属性
        If logDisableSmile = 1 Then logDisableSmile = 1 Else logDisableSmile = 0
        If logDisableImage = 1 Then logDisableImage = 1 Else logDisableImage = 0
        If logDisableURL = 1 Then logDisableURL = 0 Else logDisableURL = 1
        If logDisableKeyWord = 1 Then logDisableKeyWord = 0 Else logDisableKeyWord = 1
        If logIntroCustom = 1 Then logIntroCustom = 0 Else logIntroCustom = 1
        logUbbFlags = logDisableSmile & "0" & logDisableImage & logDisableURL & logDisableKeyWord & logIntroCustom

        If logIsDraft = False Then weblog("log_Modify") = "[本日志由 "&memName&" 于 "&DateToStr(Now(), "Y-m-d H:I A")&" 编辑]"
        If logIsDraft = False And weblog("log_IsDraft")<>logIsDraft Then
            Conn.Execute("UPDATE blog_Info SET blog_LogNums=blog_LogNums+1")
            Conn.Execute("UPDATE blog_Category SET cate_count=cate_count+1 where cate_ID=" & CheckStr(categoryID))
            SQLQueryNums = SQLQueryNums + 2
        End If
		'Meta特别属性
		if logmeta = 0 then
			logDescription = logIntro
			logKeyWords = CheckStr(logTags)
			if len(logKeyWords) = 0 then
				logKeyWords = CheckStr(logTitle)
			else
				logKeyWords = Replace(Replace(Replace(logKeyWords, ",", "|"), " ", "|"), "|", ",")
			end if
		end if

        weblog("log_Title") = CheckStr(logTitle)
        weblog("log_weather") = CheckStr(logWeather)
        weblog("log_Level") = CheckStr(logLevel)
        weblog("log_From") = CheckStr(logFrom)
        weblog("log_FromURL") = CheckStr(logFromURL)
        weblog("log_Content") = CheckStr(logMessage)
        weblog("log_Intro") = logIntro
        weblog("log_CateID") = CheckStr(categoryID)
        weblog("log_tag") = logTags
        weblog("log_ubbFlags") = logUbbFlags
        weblog("log_IsShow") = logIsShow
        weblog("log_IsTop") = logIsTop
        weblog("log_PostTime") = PubTime
        weblog("log_IsDraft") = logIsDraft
        weblog("log_DisComment") = logDisableComment
        weblog("log_edittype") = logEditType
        weblog("log_comorder") = logCommentOrder
		
		weblog("log_cname")=logcname
		weblog("log_ctype")=logctype

		if int(TitleHidden) = 1 then
		weblog("Title_Hidden")=true
		else
		weblog("Title_Hidden")=false
		end if
		weblog("log_Readpw") = logReadpw
		weblog("log_Pwtips") = logPwtips
		
		weblog("log_KeyWords") = logKeyWords
		weblog("log_Description") = logDescription
		weblog("log_Meta") = logmeta

        SQLQueryNums = SQLQueryNums + 2
        weblog.update
        weblog.Close

        Dim preLog, nextLog

        '-------------------输出静态日志档案--------------------
        '输出日志到文件
		If blog_postFile = 2 Then
		dim oldcate,oldctype,oldcname,A,B,C,D
		
		On Error Resume Next
		
		'之前如果调用过request.BinaryRead后，不能直接调用request.form了
		'live write 就挂在这里
		oldcname=request.form("oldcname")
		oldcate=request.form("oldcate")
		oldctype=request.form("oldtype")

		D = conn.execute("select cate_Part from blog_Category where cate_ID="&oldcate)(0)
		A="article/"&D
		if D = "" or len(D) = 0 then
			A = "article"
		end if
			B=oldcname
		if oldctype="0" then
			C="htm"
		else
			C="html"
		end if
		if oldcname<>request.Form("cname") or oldcate<>request.Form("log_CateID") or oldctype<>request.Form("ctype") then
			DeleteFiles Server.MapPath(A&"/"&B&"."&C)
		end if
		
		'用来检查是否有
        If Err Then
            Err.Clear
        End If
        
        PostArticle logid, False
        end if
        '输出附近的日志到文件
        Set preLog = Conn.Execute("SELECT TOP 1 T.log_Title,T.log_ID FROM blog_Content As T,blog_Category As C WHERE T.log_PostTime<#"&PubTime&"# and T.log_CateID=C.cate_ID and (T.log_IsShow=true or T.log_Readpw<>'') and C.cate_Secret=False and T.log_IsDraft=false ORDER BY T.log_PostTime DESC")
        Set nextLog = Conn.Execute("SELECT TOP 1 T.log_Title,T.log_ID FROM blog_Content As T,blog_Category As C WHERE log_PostTime>#"&PubTime&"# and T.log_CateID=C.cate_ID and (T.log_IsShow=true or T.log_Readpw<>'') and C.cate_Secret=False and T.log_IsDraft=false ORDER BY T.log_PostTime ASC")
        If Not preLog.EOF Then PostArticle preLog("log_ID"), False
        If Not nextLog.EOF Then PostArticle nextLog("log_ID"), False

        Call updateCache

        Session(CookieName&"_LastDo") = "EditArticle"
        Session(CookieName&"_draft_"&logAuthor) = conn.Execute("select count(log_ID) from blog_Content where log_Author='"&logAuthor&"' and log_IsDraft=true")(0)
        SQLQueryNums = SQLQueryNums + 1

        If logIsDraft Then
            editLog = Array(1, "日志成功保存为草稿", id)
        Else
            editLog = Array(0, "恭喜!日志编辑成功", id)
        End If

        '-------------------引用通告-------------------
        If logTrackback<>Empty And logIsShow = True And logIsDraft = False Then
            Dim log_QuoteEvery, log_QuoteArr
            log_QuoteArr = Split(logTrackback, ",")
            For Each log_QuoteEvery In log_QuoteArr
                Trackback Trim(log_QuoteEvery), siteURL&"default.asp?id="&logid, logTitle, CutStr(CheckStr(logIntro), 252), siteName
            Next
        End If
		If blog_postFile = 1 Then
		PostHalfStatic id,false
		end if
    End Function

    '*********************************************
    '删除日志
    '*********************************************

    Public Function deleteLog(id)
	    Dim pcmpad
	    pcmpad=Alias(id)
        deleteLog = Array( -4, "准备删除")
        If IsEmpty(id) Then
            getLog = Array( -5, "ID号不能为空")
            Exit Function
        End If

        If Not IsInteger(id) Then
            deleteLog = Array( -1, "非法ID号")
            Exit Function
        End If

        sqlString = "SELECT top 1 * FROM blog_Content WHERE log_ID="&id&""
        weblog.Open sqlString, Conn, 1, 3
        SQLQueryNums = SQLQueryNums + 1

        If weblog.EOF Or weblog.bof Then
            deleteLog = Array( -2, "找不到相应文章")
            Exit Function
        End If

        If stat_DelAll<>True And (stat_Del And weblog("log_Author") = logAuthor)<>True Then
            deleteLog = Array( -3, "没有权限删除")
            Exit Function
        End If

        Dim Pdate, getTag
        Dim tempTags, loadTagString, loadTags, loadTag, getTags, post_tag
        Pdate = weblog("log_PostTime")
        Conn.Execute("UPDATE blog_Member SET mem_PostLogs=mem_PostLogs-1 WHERE mem_Name='"&weblog("log_Author")&"'")
        If Not weblog("log_IsDraft") Then
            Conn.Execute("UPDATE blog_Category SET cate_count=cate_count-1 where cate_ID="&weblog("log_CateID"))
            Conn.Execute("UPDATE blog_Info SET blog_LogNums=blog_LogNums-1")
            Conn.Execute("update blog_Info set blog_CommNums=blog_CommNums-"&weblog("log_CommNums"))
        End If

        loadTag = weblog("log_tag")
        Set getTag = New Tag

        '清除旧的Tag
        If Len(loadTag)>0 Then
            loadTag = Replace(loadTag, "}{", ",")
            loadTag = Replace(loadTag, "}", "")
            loadTag = Replace(loadTag, "{", "")
            loadTags = Split(loadTag, ",")

            For Each post_tag in loadTags
                getTag.Remove post_tag
            Next
        End If
        Call Tags(2)
        Set getTag = Nothing
        Dim preLog, nextLog
        Conn.Execute("DELETE * FROM blog_Content WHERE log_ID="&id)
        Conn.Execute("DELETE * FROM blog_Comment WHERE blog_ID="&id)
        
        DeleteFiles Server.MapPath("post/"&logid&".asp")
        DeleteFiles Server.MapPath("cache/"&logid&".asp")
        DeleteFiles Server.MapPath("cache/c_"&logid&".js")
        DeleteFiles Server.MapPath(pcmpad)
		
        Set preLog = Conn.Execute("SELECT TOP 1 T.log_Title,T.log_ID FROM blog_Content As T,blog_Category As C WHERE T.log_PostTime<#"&Pdate&"# and T.log_CateID=C.cate_ID and (T.log_IsShow=true or T.log_Readpw<>'') and C.cate_Secret=False and T.log_IsDraft=false ORDER BY T.log_PostTime DESC")
        Set nextLog = Conn.Execute("SELECT TOP 1 T.log_Title,T.log_ID FROM blog_Content As T,blog_Category As C WHERE T.log_PostTime>#"&Pdate&"# and T.log_CateID=C.cate_ID and (T.log_IsShow=true or T.log_Readpw<>'') and C.cate_Secret=False and T.log_IsDraft=false ORDER BY T.log_PostTime ASC")
        '输出附近的日志到文件
        If Not preLog.EOF Then PostArticle preLog("log_ID"), False
        If Not nextLog.EOF Then PostArticle nextLog("log_ID"), False
        SQLQueryNums = SQLQueryNums + 5
        weblog.Close

        Call updateCache

        Session(CookieName&"_LastDo") = "DelArticle"
        session(CookieName&"_draft_"&logAuthor) = conn.Execute("select count(log_ID) from blog_Content where log_Author='"&logAuthor&"' and log_IsDraft=true")(0)
        SQLQueryNums = SQLQueryNums + 1
        deleteLog = Array(0, "删除成功")

    End Function

    '*********************************************
    '获得日志
    '*********************************************

    Public Function getLog(id)
        Dim getTag
        getLog = Array( -3, "准备提取日志")
        If IsEmpty(id) Then
            getLog = Array( -4, "ID号不能为空")
            Exit Function
        End If

        If Not IsInteger(id) Then
            getLog = Array( -1, "非法ID号")
            Exit Function
        End If

        sqlString = "SELECT top 1 log_CateID,log_Author,log_Title,log_edittype,log_ubbFlags,log_Intro,log_weather,log_Level,log_comorder,log_DisComment,log_IsShow,log_IsTop,log_IsDraft,log_From,log_FromURL,log_Content,log_tag,log_PostTime,log_CommNums,log_QuoteNums,log_ViewNums,log_Readpw,log_Pwtips,log_cname,log_ctype,Title_Hidden,log_KeyWords,log_Description,log_Meta FROM blog_Content WHERE log_ID="&id&""
        weblog.Open sqlString, Conn, 1, 1
        SQLQueryNums = SQLQueryNums + 1

        If weblog.EOF Or weblog.bof Then
            getLog = Array( -2, "找不到相应文章")
            Exit Function
        End If

        categoryID = weblog("log_CateID")
        logAuthor = weblog("log_Author")
        logTitle = weblog("log_Title")
        logEditType = weblog("log_edittype")
        logIntroCustom = Mid(weblog("log_ubbFlags"), 6, 1)
        logIntro = weblog("log_Intro")
        logWeather = weblog("log_weather")
        logLevel = weblog("log_Level")
        logCommentOrder = weblog("log_comorder")
        logDisableComment = weblog("log_DisComment")
        logIsShow = weblog("log_IsShow")
        logIsTop = weblog("log_IsTop")
        logIsDraft = weblog("log_IsDraft")
        logFrom = weblog("log_From")
        logFromURL = weblog("log_FromURL")
        logDisableImage = Mid(weblog("log_ubbFlags"), 3, 1)
        logDisableSmile = Mid(weblog("log_ubbFlags"), 1, 1)
        logDisableURL = Mid(weblog("log_ubbFlags"), 4, 1)
        logDisableKeyWord = Mid(weblog("log_ubbFlags"), 5, 1)
        logMessage = weblog("log_Content")
        logCommentCount = weblog("log_CommNums")
        logQuoteCount = weblog("log_QuoteNums")
        logViewCount = weblog("log_ViewNums")
		if weblog("Title_Hidden") = true then
		TitleHidden = 1
		else
		TitleHidden = 0
		end if
		
		logKeyWords = weblog("log_KeyWords")
		logDescription = weblog("log_Description")
		logmeta = weblog("log_Meta")
		
        logReadpw = Trim(weblog("log_Readpw"))
        logPwtips = weblog("log_Pwtips")
		logcname = weblog("log_cname")
		logctype = weblog("log_ctype")
        logTrackback = ""
        Set getTag = New tag
        logTags = getTag.filterEdit(weblog("log_tag"))
        Set getTag = Nothing
        logPubTime = weblog("log_PostTime")
        logPublishTimeType = "now"

        weblog.Close
        getLog = Array(0, "成功获取日志")

    End Function

    '*********************************************
    '删除文件
    '*********************************************

    Private Function DeleteFiles(FilePath)
        Dim FSO
        Set FSO = Server.CreateObject("Scripting.FileSystemObject")
        If FSO.FileExists(FilePath) Then
            FSO.DeleteFile FilePath, True
            DeleteFiles = True
        Else
            DeleteFiles = False
        End If
        Set FSO = Nothing
    End Function

    '*********************************************
    '更新缓存
    '*********************************************

    Private Sub updateCache
        Call Archive(2)
        Call CategoryList(2)
        Call getInfo(2)
        Call NewComment(2)
        Call Calendar("", "", "", 2)

        If blog_postFile>0 Then
            Dim lArticle
            Set lArticle = New ArticleCache
            lArticle.SaveCache
            Set lArticle = Nothing
        End If
    End Sub

End Class
%>
<%
'======================================================
'  PJblog2 静态缓存类
'======================================================

Class ArticleCache
    Private cacheList
    Private cacheStream
    Private errorCode
    
    Private Sub Class_Initialize()
        cacheList = ""
    End Sub

    Private Sub Class_Terminate()

    End Sub

    Private Function clearT(Str)
        Dim tempLen
        tempLen = Len(Str)
        If tempLen>0 Then
            Str = Left(Str, tempLen -1)
        End If
        clearT = Str
    End Function
    
    Private Function LoadIntro(id, cpart, cname, ctype, aisshow, aRight, outType)
        Dim getIntro, tempI, TempStr, getC, author
        if not IsEmpty(Application(CookieName&"_introCache_"&id)) then
        	getIntro = Application(CookieName&"_introCache_"&id)
        else 
		    if cpart = "" or cpart = empty or cpart = null or len(cpart) = 0 then
			   getIntro = LoadFile("cache/" & id & ".asp")
			else
       	       getIntro = LoadFile("cache/" & id & ".asp")
			end if
       	end if
        If getIntro = "error" or getIntro="" Then
            If stat_Admin Then
                response.Write "<div style=""color:#f00"">编号为[" + id + "]的日志读取失败！建议您重新 <a href=""blogedit.asp?id="&id&""" title=""编辑该日志"" accesskey=""E"">编辑</a> 此文章获得新的缓存</div>"
            End If
            Exit Function
        End If
        getIntro = Split(getIntro, "<"&"%ST(A)%"&">")
        author = Trim(getIntro(1))
        If outType = "list" Then
            If CBool(Int(aRight)) Or stat_Admin Or (Not CBool(Int(aRight)) And memName = author) Then
                tempI = getIntro(4)
            Else
                tempI = getIntro(6)
            End If
			'evio
			dim ceeurl2,chtml2
			if ctype = "0" then
		        chtml2 = "htm"
	        else
		        chtml2 = "html"
	        end if
			chtml2 ="."&chtml2
			
			ceeurl2=""

			If blog_postFile = 2 and aisshow = "True" then
			   if cpart = "" or cpart = empty or cpart = null or len(cpart) = 0 then
			      ceeurl2 = ceeurl2&"article/"&cname&chtml2
			   else
			      ceeurl2 = ceeurl2&"article/"&cpart&"/"&cname&chtml2
			   end if
			else
			   ceeurl2 = ceeurl2&"article.asp?id="&id
			end if
			tempI = Replace(tempI, "<$log_ceeurl$>", ceeurl2)
			'evio
            tempI = Replace(tempI, "<$log_viewC$>", getIntro(2))
            response.Write tempI
        Else
            TempStr = ""
            If stat_EditAll Or (stat_Edit And memName = author) Then
                TempStr = TempStr&" | <a href=""blogedit.asp?id="&id&""" title=""编辑该日志"" accesskey=""E""><img src=""images/icon_edit.gif"" alt="""" border=""0"" style=""margin-bottom:-2px""/></a> "
            End If

            If stat_DelAll Or (stat_Del And memName = author) Then
                TempStr = TempStr&" | <a href=""blogedit.asp?action=del&amp;id="&id&""" onclick=""if (!window.confirm('是否要删除该日志')) return false"" title=""删除该日志"" accesskey=""K""><img src=""images/icon_del.gif"" alt="""" border=""0"" style=""margin-bottom:-2px""/></a>"
            End If
            If CBool(Int(aRight)) Or stat_Admin Or (Not CBool(Int(aRight)) And memName = author) Then
                tempI = getIntro(3)
            Else
                tempI = getIntro(5)
            End If
			'evio
			dim ceeurl,chtml
			if ctype = "0" then
		        chtml = "htm"
	        else
		        chtml = "html"
	        end if
			chtml="."&chtml
			
			ceeurl=""

			If blog_postFile = 2 and aisshow = "True" then
			   if cpart = "" or cpart = empty or cpart = null or len(cpart) = 0 then
			      ceeurl = ceeurl&"article/"&cname&chtml
			   else
			      ceeurl = ceeurl&"article/"&cpart&"/"&cname&chtml
			   end if
			else
			   ceeurl = ceeurl&"article.asp?id="&id
			end if
			tempI = Replace(tempI, "<$log_ceeurl$>", ceeurl)
			'evio
            tempI = Replace(tempI, "<"&"%Article In PJblog2%"&">", "")
            tempI = Replace(tempI, "<$editRight$>", TempStr)
            tempI = Replace(tempI, "<$log_viewC$>", getIntro(2))
            response.Write tempI
        End If
    End Function

    Private Function LoadFile(ByVal File)
        On Error Resume Next
        LoadFile = "error"
        With cacheStream
            .Type = 2
            .Mode = 3
            .Open
            .Charset = "utf-8"
            .Position = cacheStream.Size
            .LoadFromFile Server.MapPath(File)
            If Err Then
                .Close
                Err.Clear
                Exit Function
            End If
            LoadFile = .ReadText
            .Close
        End With
    End Function

    Public Function outHTML(loadType, outType, title)
        Dim re, strMatchs, strMatch, i, j, id, aRight, hiddenC ,acpart, acname, actype, aisshow
        Set cacheStream = Server.CreateObject("ADODB.Stream")
        Set re = New RegExp
        re.IgnoreCase = True
        re.Global = True
        re.Pattern = "\[""([^\r]*?)"";([^\r]*?);\(([^\r]*?)\)\]"
        Set strMatchs = re.Execute(cacheList)
        For Each strMatch in strMatchs
            If loadType = strMatch.SubMatches(0) Then
                Dim aList, pageSize
                pageSize = blogPerPage
                If outType = "list" Then pageSize = pageSize * 4
                aList = Split(strMatch.SubMatches(2), ",")
                hiddenC = strMatch.SubMatches(1)
                If stat_Admin Or stat_ShowHiddenCate Then hiddenC = 0
                If (UBound(aList) + 1 - hiddenC)>0 Then

%>

<div class="pageContent" style="text-align:Right;overflow:hidden;height:18px;line-height:140%"><span style="float:left"><%=title%></span>预览模式: <a href="<%=Url_Add%>distype=normal" accesskey="1">普通</a> | <a href="<%=Url_Add%>distype=list" accesskey="2">列表</a></div>
<%
If outType = "list" Then response.Write "<div class=""Content-body"" style=""text-align:Left""><table cellpadding=""2"" cellspacing=""2"" width=""100%"">"
i = 0
Do Until i >= pageSize
    j = i + (CurPage -1) * pageSize
    If j<= UBound(aList) Then
        id = Split(aList(j), "|")(1)
        aRight = Split(aList(j), "|")(0)
		acpart = Split(aList(j), "|")(2)
		acname = Split(aList(j), "|")(3)
		actype = Split(aList(j), "|")(4)
		aisshow = Split(aList(j), "|")(5)
        LoadIntro id, acpart, acname, actype, aisshow, aRight, outType
        i = i + 1
    Else
        If outType = "list" Then response.Write "</table></div>"

%>
<div class="pageContent"><%=MultiPage(ubound(aList)+1-hiddenC,pageSize,CurPage,Url_Add,"","float:Left","")%></div>
<%
Exit For
End If
Loop
If outType = "list" Then response.Write "</table></div>"

%>
<div class="pageContent"><%=MultiPage(ubound(aList)+1-hiddenC,pageSize,CurPage,Url_Add,"","float:Left","")%></div>
<%
Else
    response.Write "<b>抱歉，没有找到任何日志！</b>"
End If
Set re = Nothing
Exit Function
End If
Next
Set re = Nothing
Set cacheStream = Nothing
End Function

Public Function loadCache
    Dim LoadList
    If blog_postFile<1 Then
        loadCache = False
        Exit Function
    End If
    

    if not isEmpty(Application(CookieName&"_listCache")) then
   		LoadList = Array(0,Application(CookieName&"_listCache"))
    else
   		LoadList = LoadFromFile("cache/listCache.asp")
    end if
    
    If LoadList(0) = 0 Then
        cacheList = LoadList(1)
        loadCache = True
    Else
        loadCache = False
    End If
End Function

Public Function SaveCache
    If blog_postFile<1 Then Exit Function
    Dim LogList, LogListArray, SaveList, CateDic, CateHDic, TagsDic
    Set CateDic = Server.CreateObject("Scripting.Dictionary")
    Set CateHDic = Server.CreateObject("Scripting.Dictionary")
    Set TagsDic = Server.CreateObject("Scripting.Dictionary")

    SQL = "select T.log_ID,T.log_CateID,T.log_IsShow,C.cate_Secret,C.cate_part,T.log_cname,T.log_ctype FROM blog_Content As T,blog_Category As C where T.log_CateID=C.cate_ID and log_IsDraft=false ORDER BY log_IsTop ASC,log_PostTime DESC"
    Set LogList = conn.Execute(SQL)
    If LogList.EOF Or LogList.BOF Then
    	dim temp1 
    	temp1 = "[""A"";0;()]" & Chr(13) & "[""G"";0;()]"
        SaveList = SaveToFile(temp1, "cache/listCache.asp")
'       if memoryCache = true then
			Application.Lock
			Application(CookieName&"_listCache") = temp1
			Application.UnLock
 '       end if
        Set LogList = Nothing
        Exit Function
    End If
    LogListArray = LogList.GetRows()
    Set LogList = Nothing
    Dim i, AList, AListC, GList, GListC, outIndex, tempS, tempCS, hiddenC
    AList = ""
    AListC = 0
    GList = ""
    GListC = 0
    outIndex = ""
    For i = 0 To UBound(LogListArray, 2)
        tempS = 1
        hiddenC = 1
        'response.write LogListArray(0,i) & " "
        If Not LogListArray(2, i) Then tempS = 0
        If Not LogListArray(3, i) Then
            tempCS = 0
            hiddenC = 0
            GList = GList & tempS & "|" & LogListArray(0, i) & "|" & LogListArray(4, i) & "|" & LogListArray(5, i) & "|" & LogListArray(6, i) & "|" & LogListArray(2, i) & ","
            GListC = GListC + hiddenC
        End If

        AList = AList & tempS & "|" & LogListArray(0, i) & "|" & LogListArray(4, i) & "|" & LogListArray(5, i) & "|" & LogListArray(6, i) & "|" & LogListArray(2, i) & ","
        AListC = AListC + hiddenC
        If Not CateDic.Exists("C"&LogListArray(1, i)) Then
            CateDic.Add "C"&LogListArray(1, i), tempS & "|" & LogListArray(0, i) & "|" & LogListArray(4, i) & "|" & LogListArray(5, i) & "|" & LogListArray(6, i) & "|" & LogListArray(2, i) &","
        Else
            CateDic.Item("C"&LogListArray(1, i)) = CateDic.Item("C"&LogListArray(1, i)) & tempS & "|" & LogListArray(0, i) & "|" & LogListArray(4, i) & "|" & LogListArray(5, i) & "|" & LogListArray(6, i) & "|" & LogListArray(2, i) & ","
        End If

        If Not CateHDic.Exists("CH"&LogListArray(1, i)) Then
            CateHDic.Add "CH"&LogListArray(1, i), hiddenC
        Else
            CateHDic.Item("CH"&LogListArray(1, i)) = CateHDic.Item("CH"&LogListArray(1, i)) + hiddenC
        End If

    Next
    outIndex = outIndex & "[""A"";"&AListC&";("&clearT(AList)&")] " & Chr(13)
    outIndex = outIndex & "[""G"";"&GListC&";("&clearT(GList)&")] " & Chr(13)
    Dim CateKeys, CateItems, CateHKeys, CateHItems
    CateKeys = CateDic.Keys
    CateItems = CateDic.Items
    CateHKeys = CateHDic.Keys
    CateHItems = CateHDic.Items
    For i = 0 To CateDic.Count -1
        outIndex = outIndex & "["""&CateKeys(i)&""";"&CateHItems(i)&";("&clearT(CateItems(i))&")] " & Chr(13)
    Next

    SaveList = SaveToFile(outIndex, "cache/listCache.asp")
    
   ' if memoryCache = true then
		Application.Lock
		Application(CookieName&"_listCache") = outIndex
		Application.UnLock
'	end if 
	
    Set CateDic = Nothing
    Set CateHDic = Nothing
    Set TagsDic = Nothing
    
    call newEtag
End Function

End Class
%>
<%
'======================================================
'  PJblog2 动态文章保存
'======================================================

Sub PostArticle(ByVal LogID, ByVal UpdateListOnly)
    If blog_postFile = 1 Then
        PostHalfStatic LogID, UpdateListOnly
    ElseIf blog_postFile = 2 Then
        PostFullStatic LogID, UpdateListOnly
    End If
    
    call newEtag
End Sub

'======================================================
'半静态化
'======================================================

Sub PostHalfStatic(ByVal LogID, ByVal UpdateListOnly)
    Dim SaveArticle, LoadTemplate1, Temp1, TempStr, log_View, preLogC, nextLogC

    '读取日志模块
    LoadTemplate1 = LoadFromFile("Template/Article.asp")

    If LoadTemplate1(0) <> 0 Then Exit Sub'读取成功后写入信息

    '读取分类信息
    Temp1 = LoadTemplate1(1)

    '读取日志内容
    SQL = "SELECT TOP 1 * FROM blog_Content WHERE log_ID=" & LogID
    SQLQueryNums = SQLQueryNums + 1

    Set log_View = conn.Execute(SQL)
    Dim blog_Cate, blog_CateArray, comDesc
    Dim getCate, getTags

    Set getCate = New Category
    Set getTags = New tag
    getCate.load(Int(log_View("log_CateID"))) '获取分类信息
    	
	if UpdateListOnly then '只更新列表缓存
	    PostArticleListCache LogID, log_View, getCate, getTags
	
	    Set log_View = Nothing
	    Set getCate = Nothing
	    Set getTags = Nothing
	    exit Sub
	end if 
	


    Temp1 = Replace(Temp1, "<$Cate_icon$>", getCate.cate_icon)
    Temp1 = Replace(Temp1, "<$Cate_Title$>", getCate.cate_Name)
    Temp1 = Replace(Temp1, "<$log_CateID$>", log_View("log_CateID"))
    Temp1 = Replace(Temp1, "<$LogID$>", LogID)
    Temp1 = Replace(Temp1, "<$log_Title$>", HtmlEncode(log_View("log_Title")))
    Temp1 = Replace(Temp1, "<$log_Author$>", log_View("log_Author"))
    Temp1 = Replace(Temp1, "<$log_PostTime$>", DateToStr(log_View("log_PostTime"), "Y-m-d"))

    Temp1 = Replace(Temp1, "<$log_weather$>", log_View("log_weather"))
    Temp1 = Replace(Temp1, "<$log_level$>", log_View("log_level"))
    Temp1 = Replace(Temp1, "<$log_Author$>", log_View("log_Author"))
    Temp1 = Replace(Temp1, "<$log_IsShow$>", log_View("log_IsShow"))

    If log_View("log_IsShow") and not getCate.cate_Secret Then
        Temp1 = Replace(Temp1, "<$log_hiddenIcon$>", "")
    Else
     	if log_View("log_Readpw") <> "" then
        	Temp1 = Replace(Temp1, "<$log_hiddenIcon$>", "<img src=""images/icon_lock2.gif"" style=""margin:0px 0px -3px 2px;"" alt=""加密日志"" />")
     	else
        	Temp1 = Replace(Temp1, "<$log_hiddenIcon$>", "<img src=""images/icon_lock1.gif"" style=""margin:0px 0px -3px 2px;"" alt=""私密日志"" />")
        end if
    End If

    If Len(log_View("log_tag"))>0 Then
        Temp1 = Replace(Temp1, "<$log_tag$>", getTags.filterHTML(log_View("log_tag")))
    Else
        Temp1 = Replace(Temp1, "<$log_tag$>", "")
    End If

    If log_View("log_comorder") Then comDesc = "Desc" Else comDesc = "Asc" End If

    Temp1 = Replace(Temp1, "<$comDesc$>", comDesc)
    Temp1 = Replace(Temp1, "<$log_DisComment$>", log_View("log_DisComment"))

    If log_View("log_edittype") = 1 Then
        Temp1 = Replace(Temp1, "<$ArticleContent$>", UnCheckStr(UBBCode(HtmlEncode(log_View("log_Content")), Mid(log_View("log_ubbFlags"), 1, 1), Mid(log_View("log_ubbFlags"), 2, 1), Mid(log_View("log_ubbFlags"), 3, 1), Mid(log_View("log_ubbFlags"), 4, 1), Mid(log_View("log_ubbFlags"), 5, 1))))
    Else
        Temp1 = Replace(Temp1, "<$ArticleContent$>", UnCheckStr(log_View("log_Content")))
    End If

    If Len(log_View("log_Modify"))>0 Then
        Temp1 = Replace(Temp1, "<$log_Modify$>", log_View("log_Modify")&"<br/>")
    Else
        Temp1 = Replace(Temp1, "<$log_Modify$>", "")
    End If

    Temp1 = Replace(Temp1, "<$log_FromUrl$>", log_View("log_FromUrl"))
    Temp1 = Replace(Temp1, "<$log_From$>", log_View("log_From"))
    Temp1 = Replace(Temp1, "<$trackback$>", SiteURL&"trackback.asp?tbID="&LogID&"&amp;action=view")

    Temp1 = Replace(Temp1, "<$log_CommNums$>", log_View("log_CommNums"))
    Temp1 = Replace(Temp1, "<$log_QuoteNums$>", log_View("log_QuoteNums"))

    Temp1 = Replace(Temp1, "<$log_IsDraft$>", log_View("log_IsDraft"))

    Set preLogC = Conn.Execute("SELECT TOP 1 T.log_Title,T.log_ID FROM blog_Content As T,blog_Category As C WHERE T.log_PostTime<#"&DateToStr(log_View("log_PostTime"), "Y-m-d H:I:S")&"# and T.log_CateID=C.cate_ID and (T.log_IsShow=true or T.log_Readpw<>'') and C.cate_Secret=False and T.log_IsDraft=false ORDER BY T.log_PostTime DESC")
    Set nextLogC = Conn.Execute("SELECT TOP 1 T.log_Title,T.log_ID FROM blog_Content As T,blog_Category As C WHERE T.log_PostTime>#"&DateToStr(log_View("log_PostTime"), "Y-m-d H:I:S")&"# and T.log_CateID=C.cate_ID and (T.log_IsShow=true or T.log_Readpw<>'') and C.cate_Secret=False and T.log_IsDraft=false ORDER BY T.log_PostTime ASC")

    Dim BTemp,urlLink
    BTemp = ""
    
    If Not preLogC.EOF Then
    	if blog_postFile = 2 then
    		urlLink = Alias(preLogC("log_ID"))
    	else 
    		urlLink = "?id="&preLogC("log_ID")
    	end if
        BTemp = BTemp & "<a href="""&urlLink&""" title=""上一篇日志: "&preLogC("log_Title")&""" accesskey="",""><img border=""0"" src=""images/Cprevious.gif"" alt=""""/>上一篇</a>"
    Else
        BTemp = BTemp & "<img border=""0"" src=""images/Cprevious1.gif"" alt=""这是最新一篇日志""/>上一篇"
    End If

    If Not nextLogC.EOF Then
    	if blog_postFile = 2 then
    		urlLink = Alias(nextLogC("log_ID"))
    	else 
    		urlLink = "?id="&nextLogC("log_ID")
    	end if
        BTemp = BTemp & " | <a href="""&urlLink&""" title=""下一篇日志: "&nextLogC("log_Title")&""" accesskey="".""><img border=""0"" src=""images/Cnext.gif"" alt=""""/>下一篇</a>"
    Else
        BTemp = BTemp & " | <img border=""0"" src=""images/Cnext1.gif"" alt=""这是最后一篇日志""/>下一篇"
    End If

    Temp1 = Replace(Temp1, "<$log_Navigation$>", BTemp)

    SaveArticle = SaveToFile(Temp1, "post/" & LogID & ".asp")

    PostArticleListCache LogID, log_View, getCate, getTags

    Set log_View = Nothing
    Set getCate = Nothing
    Set getTags = Nothing
    'getCate.cate_Secret or (not log_View("Log_IsShow"))
End Sub

'======================================================
'全静态化
'======================================================

Sub PostFullStatic(ByVal LogID, ByVal UpdateListOnly)
    Dim SaveArticle, LoadTemplate1, Temp1, TempStr, log_View, preLogC, nextLogC, Category,baseUrl
    
    '读取日志模块
    LoadTemplate1 = LoadFromFile("Template/static.htm")

    If LoadTemplate1(0) <> 0 Then Exit Sub'读取成功后写入信息
    
    '静态页面特有的属性
	baseUrl = "http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("URL")
	baseUrl = Left(baseUrl, InStrRev(baseUrl,"/"))
	
    '读取分类信息
    Temp1 = LoadTemplate1(1)

    '读取日志内容
    SQL = "SELECT TOP 1 * FROM blog_Content WHERE log_ID=" & LogID
    SQLQueryNums = SQLQueryNums + 1
	
    Set log_View = conn.Execute(SQL)
    
    blog_currentCategoryID = log_View("log_CateID")
    Dim blog_Cate, blog_CateArray, comDesc
    Dim getCate, getTags
	
    Set getCate = New Category
    Set getTags = New tag

    getCate.load(Int(log_View("log_CateID"))) '获取分类信息
    
	if UpdateListOnly then '只更新列表缓存
		PostArticleListCache LogID, log_View, getCate, getTags
		
	    Set log_View = Nothing
	    Set getCate = Nothing
	    Set getTags = Nothing
	    exit Sub
	end if 

	if log_View("log_IsShow") = false or getCate.cate_Secret then '如果是私密日志
    	SaveArticle = SaveToFile("<head><title>显示私密日志</title><meta http-equiv=""Refresh"" content=""0;url="&baseUrl&"article.asp?id="&LogID&"""></head><body><h1>正在跳转到私密日志</h1>您可以在<a HREF="""&baseUrl&"article.asp?id="&LogID&""">此处</a>进行访问.</body>", Alias(LogID))
    	PostHalfStatic LogID, UpdateListOnly
	    
	    Set log_View = Nothing
	    exit Sub
	end if
	
    If log_View("log_comorder") Then comDesc = "Desc" Else comDesc = "Asc" End If	
    
    Temp1 = Replace(Temp1, "<$CategoryList$>", CategoryList(0))
    Temp1 = Replace(Temp1, "<$base$>", baseUrl)
    Temp1 = Replace(Temp1, "<$siteName$>", siteName)    
    Temp1 = Replace(Temp1, "<$blog_Title$>", blog_Title)
    Temp1 = Replace(Temp1, "<$skin$>", blog_DefaultSkin)   
    Temp1 = Replace(Temp1, "<$blogabout$>", blogabout)   
    Temp1 = Replace(Temp1, "<$comDesc$>", comDesc)   
    Temp1 = Replace(Temp1, "<$CookieName$>", CookieName)
    Temp1 = Replace(Temp1, "<$blog_version$>", blog_version)
    
   '输出第一页评论

    Temp1 = Replace(Temp1, "<$comment$>", ShowComm(LogID, comDesc, log_View("log_DisComment"), True, log_View("log_IsShow")))   
    
    
    Temp1 = Replace(Temp1, "<$Cate_icon$>", getCate.cate_icon)
    Temp1 = Replace(Temp1, "<$Cate_Title$>", getCate.cate_Name)
    Temp1 = Replace(Temp1, "<$log_CateID$>", log_View("log_CateID"))
    Temp1 = Replace(Temp1, "<$LogID$>", LogID)
    Temp1 = Replace(Temp1, "<$log_Title$>", HtmlEncode(log_View("log_Title")))
    Temp1 = Replace(Temp1, "<$log_Author$>", log_View("log_Author"))
    Temp1 = Replace(Temp1, "<$log_PostTime$>", DateToStr(log_View("log_PostTime"), "Y-m-d"))

    Temp1 = Replace(Temp1, "<$log_weather$>", log_View("log_weather"))
    Temp1 = Replace(Temp1, "<$log_level$>", log_View("log_level"))
    Temp1 = Replace(Temp1, "<$log_Author$>", log_View("log_Author"))
	
	'evio
	if len(log_View("log_KeyWords")) > 0 then
		Temp1 = Replace(Temp1, "<$keywords$>", log_View("log_KeyWords"))
	end if
	if len(log_View("log_Description")) > 0 then
		Temp1 = Replace(Temp1, "<$description$>", log_View("log_Description"))
	end if
	'evio

    If Len(log_View("log_tag"))>0 Then
        Temp1 = Replace(Temp1, "<$log_tag$>", getTags.filterHTML(log_View("log_tag")))
    Else
        Temp1 = Replace(Temp1, "<$log_tag$>", "")
    End If

    Temp1 = Replace(Temp1, "<$comDesc$>", comDesc)
    Temp1 = Replace(Temp1, "<$log_DisComment$>", log_View("log_DisComment"))

    If log_View("log_edittype") = 1 Then
        Temp1 = Replace(Temp1, "<$ArticleContent$>", UnCheckStr(UBBCode(HtmlEncode(log_View("log_Content")), Mid(log_View("log_ubbFlags"), 1, 1), Mid(log_View("log_ubbFlags"), 2, 1), Mid(log_View("log_ubbFlags"), 3, 1), Mid(log_View("log_ubbFlags"), 4, 1), Mid(log_View("log_ubbFlags"), 5, 1))))
    Else
        Temp1 = Replace(Temp1, "<$ArticleContent$>", UnCheckStr(log_View("log_Content")))
    End If

    If Len(log_View("log_Modify"))>0 Then
        Temp1 = Replace(Temp1, "<$log_Modify$>", log_View("log_Modify")&"<br/>")
    Else
        Temp1 = Replace(Temp1, "<$log_Modify$>", "")
    End If

    Temp1 = Replace(Temp1, "<$log_FromUrl$>", log_View("log_FromUrl"))
    Temp1 = Replace(Temp1, "<$log_From$>", log_View("log_From"))
    Temp1 = Replace(Temp1, "<$trackback$>", SiteURL&"trackback.asp?tbID="&LogID&"&amp;action=view")

    Temp1 = Replace(Temp1, "<$log_CommNums$>", log_View("log_CommNums"))
    Temp1 = Replace(Temp1, "<$log_QuoteNums$>", log_View("log_QuoteNums"))

    Temp1 = Replace(Temp1, "<$log_IsDraft$>", log_View("log_IsDraft"))

    Set preLogC = Conn.Execute("SELECT TOP 1 T.log_Title,T.log_ID FROM blog_Content As T,blog_Category As C WHERE T.log_PostTime<#"&DateToStr(log_View("log_PostTime"), "Y-m-d H:I:S")&"# and T.log_CateID=C.cate_ID and (T.log_IsShow=true or T.log_Readpw<>'') and C.cate_Secret=False and T.log_IsDraft=false ORDER BY T.log_PostTime DESC")
    Set nextLogC = Conn.Execute("SELECT TOP 1 T.log_Title,T.log_ID FROM blog_Content As T,blog_Category As C WHERE T.log_PostTime>#"&DateToStr(log_View("log_PostTime"), "Y-m-d H:I:S")&"# and T.log_CateID=C.cate_ID and (T.log_IsShow=true or T.log_Readpw<>'') and C.cate_Secret=False and T.log_IsDraft=false ORDER BY T.log_PostTime ASC")

    Dim BTemp
    BTemp = ""
    If Not preLogC.EOF Then
        BTemp = BTemp & "<a href="""&Alias(preLogC("log_ID"))&""" title=""上一篇日志: "&preLogC("log_Title")&""" accesskey="",""><img border=""0"" src=""images/Cprevious.gif"" alt=""""/>上一篇</a>"
    Else
        BTemp = BTemp & "<img border=""0"" src=""images/Cprevious1.gif"" alt=""这是最新一篇日志""/>上一篇"
    End If

    If Not nextLogC.EOF Then
        BTemp = BTemp & " | <a href="""&Alias(nextLogC("log_ID"))&""" title=""下一篇日志: "&nextLogC("log_Title")&""" accesskey="".""><img border=""0"" src=""images/Cnext.gif"" alt=""""/>下一篇</a>"
    Else
        BTemp = BTemp & " | <img border=""0"" src=""images/Cnext1.gif"" alt=""这是最后一篇日志""/>下一篇"
    End If

    Temp1 = Replace(Temp1, "<$log_Navigation$>", BTemp)
	
	createfolder "article/"&getCate.cate_Part

    SaveArticle = SaveToFile(Temp1, Alias(LogID))

    PostArticleListCache LogID, log_View, getCate , getTags

    Set log_View = Nothing
    Set getCate = Nothing
    Set getTags = Nothing
    'getCate.cate_Secret or (not log_View("Log_IsShow"))
End Sub

'======================================================
'缓存静态化列表
'======================================================

Sub PostArticleListCache(ByVal LogID,ByVal log_View,ByVal getCate,ByVal getTags)
    Dim LoadTemplate2, Temp2, comDesc, SaveArticle
    LoadTemplate2 = LoadFromFile("Template/ArticleList.asp")
    If LoadTemplate2(0) <> 0 Then Exit Sub

    Temp2 = LoadTemplate2(1)
    Temp2 = Replace(Temp2, "<$Cate_icon$>", getCate.cate_icon)
    Temp2 = Replace(Temp2, "<$Cate_Title$>", getCate.cate_Name)
    Temp2 = Replace(Temp2, "<$log_CateID$>", log_View("log_CateID"))
    Temp2 = Replace(Temp2, "<$LogID$>", LogID)
    Temp2 = Replace(Temp2, "<$log_Title$>", HtmlEncode(log_View("log_Title")))
    Temp2 = Replace(Temp2, "<$log_Author$>", log_View("log_Author"))
    Temp2 = Replace(Temp2, "<$log_PostTime$>", DateToStr(log_View("log_PostTime"), "Y-m-d"))
    Temp2 = Replace(Temp2, "<$log_viewCount$>", log_View("log_ViewNums"))
    
    'article.asp?id=<$LogID$>
    If blog_postFile = 2  and log_View("log_IsShow") and not getCate.cate_Secret Then
        Temp2 = Replace(Temp2, "<$pLink$>", Alias(LogID))
    else
	 	Temp2 = Replace(Temp2, "<$pLink$>", "article.asp?id=" & LogID)
    end if 

    
    
    If log_View("log_IsTop") Then
        Temp2 = Replace(Temp2, "<$ShowButton$>", "<div class=""BttnE"" onclick=""TopicShow(this,'log_"&LogID&"')""></div>")
        Temp2 = Replace(Temp2, "<$ShowStyle$>", " style=""display:none""")
    Else
        Temp2 = Replace(Temp2, "<$ShowButton$>", "")
        Temp2 = Replace(Temp2, "<$ShowStyle$>", "")
    End If

    If log_View("log_IsShow") and not getCate.cate_Secret Then
        Temp2 = Replace(Temp2, "<$log_hiddenIcon$>", "")
    Else
	    If log_View("log_Readpw") <> "" Then
	        Temp2 = Replace(Temp2, "<$log_Secret$>", "该日志是加密日志，需要输入正确密码才可以查看！")
	        Temp2 = Replace(Temp2, "<$log_hiddenIcon$>", "<img src=""images/icon_lock2.gif"" style=""margin:0px 0px -3px 2px;"" alt=""加密日志"" />")
	        if log_View("Title_Hidden") = true then
			Temp2 = Replace(Temp2, "<$Show_Title$>", log_View("log_Title"))
			else
	        Temp2 = Replace(Temp2, "<$Show_Title$>", "[加密日志]")
			end if
	    Else
	        Temp2 = Replace(Temp2, "<$log_Secret$>", "该日志是私密日志，只有管理员或发布者可以查看！")
	        Temp2 = Replace(Temp2, "<$log_hiddenIcon$>", "<img src=""images/icon_lock1.gif"" style=""margin:0px 0px -3px 2px;"" alt=""私密日志"" />")
	        if log_View("Title_Hidden") = true then
			Temp2 = Replace(Temp2, "<$Show_Title$>", log_View("log_Title"))
			else
	        Temp2 = Replace(Temp2, "<$Show_Title$>", "[私密日志]")
			end if
	    End If
    End If

    If Len(log_View("log_tag"))>0 Then
        Temp2 = Replace(Temp2, "<$log_tag$>", "<p>Tags: "&getTags.filterHTML(log_View("log_tag"))&"</p>")
    Else
        Temp2 = Replace(Temp2, "<$log_tag$>", "")
    End If

    If log_View("log_comorder") Then comDesc = "Desc" Else comDesc = "Asc" End If

    If log_View("log_edittype") = 1 Then
        Temp2 = Replace(Temp2, "<$log_Intro$>", UnCheckStr(UBBCode(log_View("log_Intro"), Mid(log_View("log_ubbFlags"), 1, 1), Mid(log_View("log_ubbFlags"), 2, 1), Mid(log_View("log_ubbFlags"), 3, 1), Mid(log_View("log_ubbFlags"), 4, 1), Mid(log_View("log_ubbFlags"), 5, 1))))
        If log_View("log_Intro")<>HtmlEncode(log_View("log_Content")) Then
        
            If blog_postFile = 2 and log_View("log_IsShow") and not getCate.cate_Secret Then
          	   Temp2 = Replace(Temp2, "<$log_readMore$>", "<p><a href="""&Alias(LogID)&""" class=""more"">查看更多...</a></p>")
			else
          	   Temp2 = Replace(Temp2, "<$log_readMore$>", "<p><a href=""article.asp?id="&LogID&""" class=""more"">查看更多...</a></p>")           
  			End If
  			
        Else
            Temp2 = Replace(Temp2, "<$log_readMore$>", "")
        End If
    Else
        Temp2 = Replace(Temp2, "<$log_Intro$>", UnCheckStr(log_View("log_Intro")))
        If log_View("log_Intro")<>log_View("log_Content") Then
            If blog_postFile = 2 and log_View("log_IsShow") and not getCate.cate_Secret Then
             	   Temp2 = Replace(Temp2, "<$log_readMore$>", "<p><a href="""&Alias(LogID)&""" class=""more"">查看更多...</a></p>")
         	else
             	   Temp2 = Replace(Temp2, "<$log_readMore$>", "<p><a href=""article.asp?id="&LogID&""" class=""more"">查看更多...</a></p>")
         	end if           
        Else
            Temp2 = Replace(Temp2, "<$log_readMore$>", "")
        End If
    End If

    Temp2 = Replace(Temp2, "<$log_CommNums$>", log_View("log_CommNums"))
    Temp2 = Replace(Temp2, "<$log_QuoteNums$>", log_View("log_QuoteNums"))

    SaveArticle = SaveToFile(Temp2, "cache/" & LogID & ".asp")
    
    if memoryCache = true then
		Application.Lock
		Application(CookieName&"_introCache_"&LogID) = Temp2
		Application.UnLock
	end if
End Sub

'======================================================
'模板文件保存到内存里
'======================================================

Sub LoadTemplateFile(Path)
    Dim cache
End Sub

%>