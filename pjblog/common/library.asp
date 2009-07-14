<%
'***************PJblog2 模块与类处理*******************
' PJblog2 Copyright 2005
' Update:2005-10-20
'**************************************************

'**********************************************
'BLOG日历
'**********************************************

Function Calendar(C_Year, C_Month, C_Day, update)
    Dim C_YM, S_Date, E_Date, isDo, Dclass, RS_Month, Link_TF, i, DayCount, DayStr, NotCM, TS_Date, TE_Date
    If C_Year = Empty Then C_Year = Year(Now())
    If C_Month = Empty Then C_Month = Month(Now())
    If C_Day = Empty Then C_Day = 0
    C_Year = CInt(C_Year)
    C_Month = CInt(C_Month)
    C_Day = CInt(C_Day)
    C_YM = C_Year & "-" & C_Month
    Dim PY, PM, NY, NM
    PM = C_Month -1
    If PM<1 Then
        PM = 12
        PY = C_Year -1
    Else
        PY = C_Year
    End If
    NM = C_Month + 1
    If NM>12 Then
        NM = 1
        NY = C_Year + 1
    Else
        NY = C_Year
    End If
    Calendar = "<div id=""Calendar_Body""><div id=""Calendar_Top""><div id=""LeftB"" onclick=""location.href='"&GetbaseUrl&"default.asp?log_Year="&PY&"&amp;log_Month="&PM&"'""></div><div id=""RightB"" onclick=""location.href='"&GetbaseUrl&"default.asp?log_Year="&NY&"&amp;log_Month="&NM&"'""></div>"&C_Year&"年"&C_Month&"月</div>"
    Calendar = Calendar & "<div id=""Calendar_week""><ul class=""Week_UL""><li><font color=""#FF0000"">日</font></li><li>一</li><li>二</li><li>三</li><li>四</li><li>五</li><li>六</li></ul></div>"

    '--->计算当前月份的日期
    i = Weekday(C_YM & "-" & 1) -1
    TS_Date = DateSerial(C_Year, C_Month, 1 - i)
    TE_Date = DateAdd("d", 42, TS_Date)
    S_Date = Year(TS_Date)&"-"&Month(TS_Date)&"-"&Day(TS_Date)
    E_Date = Year(TE_Date)&"-"&Month(TE_Date)&"-"&Day(TE_Date)

    '--->保存日志日历缓存
    Dim Link_Count, Link_Days, CalendarArray, doUpdate, upTime
    upTime = Year(Now())&"-"&Month(Now())
    doUpdate = False

    If Not IsArray(Application(CookieName&"_blog_Calendar")) Then '判断日期更新条件
        doUpdate = True
    ElseIf Application(CookieName&"_blog_Calendar")(1)<>upTime Then
        doUpdate = True
    ElseIf upTime<>C_Year&"-"&C_Month Then
        doUpdate = True
    ElseIf update = 2 Then
        doUpdate = True
    End If

    If doUpdate Then
        ReDim Link_Days(4, 0)
        Link_Count = 0
        SQL = "SELECT C.log_id,C.log_title,C.log_PostTime,C.log_IsShow FROM blog_Content as C,blog_Category as A where C.log_PostTime Between #"&S_Date&" 00:00:00# And #"&E_Date&" 23:59:59# and C.log_IsDraft=false and C.log_CateID=A.cate_ID and A.cate_Secret=false ORDER BY C.log_PostTime"
        Set RS_Month = Conn.Execute(SQL)
        SQLQueryNums = SQLQueryNums + 1
        Dim the_Day, TempTitle, TempCount, TempSplit
        the_Day = 0
        TempCount = 0
        TempTitle = ""
        Do While Not RS_Month.EOF
            If Day(RS_Month("log_PostTime"))<>the_Day Then
                the_Day = Day(RS_Month("log_PostTime"))
                ReDim PreServe Link_Days(4, Link_Count)
                Link_Days(0, Link_Count) = Year(RS_Month("log_PostTime"))
                Link_Days(1, Link_Count) = Month(RS_Month("log_PostTime"))
                Link_Days(2, Link_Count) = Day(RS_Month("log_PostTime"))
                Link_Days(3, Link_Count) = "default.asp?log_Year="&Year(RS_Month("log_PostTime"))&"&amp;log_Month="&Month(RS_Month("log_PostTime"))&"&amp;log_Day="&Day(RS_Month("log_PostTime"))
                TempCount = 1
                If RS_Month("log_IsShow") Then
                    TempTitle = Chr(13) & " - " & RS_Month("log_title")
                Else
                    TempTitle = Chr(13) & " - [私密日志]"
                End If
                Link_Days(4, Link_Count) = "当天共写了" & TempCount &"篇日志" & TempTitle
                Link_Count = Link_Count + 1
            Else
                TempCount = TempCount + 1
                If RS_Month("log_IsShow") Then
                    Link_Days(4, Link_Count -1) = Link_Days(4, Link_Count -1) & Chr(10) & " - " & RS_Month("log_title")
                Else
                    Link_Days(4, Link_Count -1) = Link_Days(4, Link_Count -1) & Chr(10) & " - [私密日志]"
                End If

                TempSplit = Split(Link_Days(4, Link_Count -1), Chr(13))
                TempSplit(0) = "当天共写了" & TempCount &"篇日志" & Chr(13)
                If UBound(TempSplit)>0 Then Link_Days(4, Link_Count -1) = TempSplit(0) & TempSplit(1)
            End If
            RS_Month.MoveNext
        Loop
        Set RS_Month = Nothing
        'response.write "<script>alert('动态日历')</script>"
        If upTime = C_Year&"-"&C_Month Then
            CalendarArray = Array(Link_Days, upTime)
            Application.Lock
            Application(CookieName&"_blog_Calendar") = CalendarArray
            Application.UnLock
            'response.write "<script>alert('写日历缓存')</script>"
        End If
    Else
        Link_Days = Application(CookieName&"_blog_Calendar")(0)
        Link_Count = UBound(Link_Days, 2) + 1
        'response.write "<script>alert('静态日历')</script>"
    End If

    If update = 2 Then Exit Function

    Dim DayEnd, Calendar_Count
    Calendar_Count = 0
    DayEnd = False
    DayCount = 0
    Dclass = ""
    DayStr = ""
    isDo = 0
    NotCM = 1
    Do Until Month(S_Date)<>C_Month And NotCM = 7
        If DayCount>6 Then
            Calendar = Calendar & "<div class=""Calendar_Day""><ul class=""Day_UL"">"&DayStr&"</ul></div>"
            DayCount = 0
            DayStr = ""
        End If
        If Calendar_Count = Link_Count Then
            Calendar_Count = Link_Count -1
            DayEnd = True
        End If
        If Month(S_Date) = C_Month Then NotCM = 0
        If Month(S_Date)<>C_Month Then
            Dclass = "class=""otherday"""
            NotCM = NotCM + 1
        ElseIf Year(S_Date) = Year(Now()) And Month(S_Date) = Month(Now()) And Day(S_Date) = Day(Now()) Then
            Dclass = "class=""today"""
        Else
            Dclass = ""
        End If
        If Link_Count>0 Then
            If Link_Days(1, Calendar_Count) = Month(S_Date) And Link_Days(2, Calendar_Count) = Day(S_Date) And DayEnd = False Then
                If Month(S_Date)<>C_Month Then
                    Dclass = "class=""otherday"""
                ElseIf Day(S_Date) = C_Day Then
                    Dclass = "class=""click"""
                ElseIf C_Year = Year(Now()) And C_Month = Month(Now()) And Day(S_Date) = Day(Now()) Then
                    Dclass = "class=""DayD"""
                Else
                    Dclass = "class=""haveD"""
                End If
                DayStr = DayStr&"<li class=""DayA""><a "&Dclass&" href="""&Link_Days(3, Calendar_Count)&""" title="""&Link_Days(4, Calendar_Count)&""">"&Day(S_Date)&"</a></li>"
                Calendar_Count = Calendar_Count + 1
            Else
                DayStr = DayStr&"<li class=""DayA""><a "&Dclass&">"&Day(S_Date)&"</a></li>"
            End If
        Else
            DayStr = DayStr&"<li class=""DayA""><a "&Dclass&">"&Day(S_Date)&"</a></li>"
        End If
        DayCount = DayCount + 1
        S_Date = DateAdd("d", 1, S_Date)
    Loop
    Calendar = Calendar & "</div>"
End Function


'**********************************************
'用户面板
'**********************************************

Function userPanel()
    userPanel = ""
    If memName<>Empty Then userPanel = userPanel&" <b>"&memName&"</b>，欢迎你!<br/>你的权限: "&stat_title&"<br/><br/>"
    If stat_Admin = True Then userPanel = userPanel + "<a href=""control.asp"" target=""_blank"" class=""sideA"" accesskey=""3"">系统管理</a>"
    If stat_AddAll = True Or stat_Add = True Then userPanel = userPanel + "<a href=""blogpost.asp"" class=""sideA"" accesskey=""N"">发表新日志</a>"
    If (stat_AddAll = True Or stat_Add = True) And (stat_EditAll Or stat_Edit) Then
        If IsEmpty(session(CookieName&"_draft_"&memName)) Then
            If stat_EditAll Then
                session(CookieName&"_draft_"&memName) = conn.Execute("select count(log_ID) from blog_Content where log_IsDraft=true")(0)
                SQLQueryNums = SQLQueryNums + 1
            ElseIf stat_Edit Then
                session(CookieName&"_draft_"&memName) = conn.Execute("select count(log_ID) from blog_Content where log_IsDraft=true and log_Author='"&memName&"'")(0)
                SQLQueryNums = SQLQueryNums + 1
            Else
	            session(CookieName&"_draft_"&memName) = 0
            End If
        End If
        
        If session(CookieName&"_draft_"&memName) > 0 Then
            userPanel = userPanel + "<a href=""default.asp?display=draft"" class=""sideA"" accesskey=""D""><strong>编辑草稿 ["&session(CookieName&"_draft_"&memName)&"]</strong></a>"
        Else
            userPanel = userPanel + "<a href=""default.asp?display=draft"" class=""sideA"" accesskey=""D"">编辑草稿</a>"
        End If
    End If
    If memName<>Empty Then
        userPanel = userPanel&"<a href=""member.asp?action=edit"" class=""sideA"" accesskey=""M"">修改个人资料</a><a href=""login.asp?action=logout"" class=""sideA"" accesskey=""Q"">退出系统</a>"
    Else
        userPanel = userPanel&"<a href=""login.asp"" class=""sideA"" accesskey=""L"">登录</a><a href=""register.asp"" class=""sideA"" accesskey=""U"">用户注册</a>"
        If blog_PasswordProtection Then
        	userPanel = userPanel&"<a href=""javascript:void(0)"" class=""sideA"" accesskey=""P"" onclick=""PasswordProtection();"">忘记密码</a>"
        End If
    End If
End Function


'**********************************************
'输出日志统计信息
'**********************************************

Function info_code(Str)
    Dim vOnline
    vOnline = getOnline
    Str = Replace(Str, "$blog_LogNums$", blog_LogNums)
    Str = Replace(Str, "$blog_CommNums$", blog_CommNums)
    Str = Replace(Str, "$blog_TbCount$", blog_TbCount)
    Str = Replace(Str, "$blog_MessageNums$", blog_MessageNums)
    Str = Replace(Str, "$blog_MemNums$", blog_MemNums)
    Str = Replace(Str, "$blog_VisitNums$", blog_VisitNums)
    Str = Replace(Str, "$blog_OnlineNums$", vOnline)
    info_code = Str
End Function

'**********************************************
'获取在线人数
'**********************************************

Function getOnline
    getOnline = 1
    If Len(Application(CookieName&"_onlineCount"))>0 Then
        If DateDiff("s", Application(CookieName&"_userOnlineCountTime"), Now())>60 Then
            Application.Lock()
            Application(CookieName&"_online") = Application(CookieName&"_onlineCount")
            Application(CookieName&"_onlineCount") = 1
            Application(CookieName&"_onlineCountKey") = randomStr(2)
            Application(CookieName&"_userOnlineCountTime") = Now()
            Application.Unlock()
        Else
            If Session(CookieName&"userOnlineKey")<>Application(CookieName&"_onlineCountKey") Then
                Application.Lock()
                Application(CookieName&"_onlineCount") = Application(CookieName&"_onlineCount") + 1
                Application.Unlock()
                Session(CookieName&"userOnlineKey") = Application(CookieName&"_onlineCountKey")
            End If
        End If
    Else
        Application.Lock
        Application(CookieName&"_online") = 1
        Application(CookieName&"_onlineCount") = 1
        Application(CookieName&"_onlineCountKey") = randomStr(2)
        Application(CookieName&"_userOnlineCountTime") = Now()
        Application.Unlock
    End If
    getOnline = Application(CookieName&"_online")
End Function


'**********************************************
'侧边模版处理
'**********************************************

Sub Side_Module_Replace()
    '日历处理
    Dim Cal_code
    Cal_code = Calendar(log_Year, log_Month, log_Day, 1)
    side_html_default = Replace(side_html_default, "$calendar_code$", Cal_code)
    side_html = Replace(side_html, "$calendar_code$", Cal_code)
    side_html_static = Replace(side_html_static, "$calendar_code$", Cal_code)
    
    '用户面板处理
    Dim user_code
    user_code = userPanel
    side_html_default = Replace(side_html_default, "$user_code$", user_code)
    side_html = Replace(side_html, "$user_code$", user_code)
    side_html_static = Replace(side_html_static, "$user_code$", user_code)
    
    '归档面板处理
    Dim archive_code
    archive_code = Archive(1)
    side_html_default = Replace(side_html_default, "$archive_code$", archive_code)
    side_html = Replace(side_html, "$archive_code$", archive_code)
    side_html_static = Replace(side_html_static, "$archive_code$", archive_code)
    
    '树形分类处理
    Dim Category_code
    Category_code = CategoryList(1)
    side_html_default = Replace(side_html_default, "$Category_code$", Category_code)
    side_html = Replace(side_html, "$Category_code$", Category_code)
    side_html_static = Replace(side_html_static, "$Category_code$", Category_code)
    
    '显示统计信息
    side_html_default = info_code(side_html_default)
    side_html = info_code(side_html)
    side_html_static = info_code(side_html_static)
    
    '处理最新评论内容
    Dim Comment_code
    Comment_code = NewComment(1)
    side_html_default = Replace(side_html_default, "$comment_code$", Comment_code)
    side_html = Replace(side_html, "$comment_code$", Comment_code)
    side_html_static = Replace(side_html_static, "$comment_code$", Comment_code)
    
    '处理友情链接内容
    Dim Link_Code
    Link_Code = Bloglinks(1)
    side_html_default = Replace(side_html_default, "$Link_Code$", Link_Code)
    side_html = Replace(side_html, "$Link_Code$", Link_Code)
    side_html_static = Replace(side_html_static, "$Link_Code$", Link_Code)
End Sub




'==============================================================
' Blog Class
'==============================================================

'*******************************************
'  分类读取Class
'*******************************************

Class Category
    Public cate_ID
    Public cate_Name
    Public cate_Part
    Public cate_Order
    Public cate_Intro
    Public cate_OutLink
    Public cate_URL
    Public cate_icon
    Public cate_count
    Public cate_Lock
    Public cate_local
    Public cate_Secret
    Private LastID
    Private Loaded

    Private Sub Class_Initialize()
        cate_ID = 0
        cate_Name = ""
        cate_Part = ""
        cate_Order = 0
        cate_Intro = ""
        cate_OutLink = False
        cate_URL = ""
        cate_icon = ""
        cate_count = ""
        cate_Lock = False
        cate_local = ""
        cate_Secret = False
        LastID = -99
        Loaded = False
    End Sub

    Private Sub Class_Terminate()

    End Sub


    Public Sub Reload
        CategoryList(2) '更新分类缓存
    End Sub

    Public Function Load(ID)
        Dim blog_Cate, blog_CateArray, Category_Len, i
        If Int(ID) = LastID Then Exit Function
        If Not IsArray(Application(CookieName&"_blog_Category")) Then Reload
        blog_CateArray = Application(CookieName&"_blog_Category")
        If UBound(blog_CateArray, 1) = 0 Then Exit Function
        Category_Len = UBound(blog_CateArray, 2)
        For i = 0 To Category_Len
            If Int(blog_CateArray(0, i)) = Int(ID) Then
                cate_ID = blog_CateArray(0, i)
                cate_Name = blog_CateArray(1, i)
                cate_Order = blog_CateArray(2, i)
                cate_Intro = blog_CateArray(3, i)
                cate_OutLink = blog_CateArray(4, i)
                cate_URL = blog_CateArray(5, i)
                cate_icon = blog_CateArray(6, i)
                cate_count = blog_CateArray(7, i)
                cate_Lock = blog_CateArray(8, i)
                cate_local = blog_CateArray(9, i)
                cate_Secret = blog_CateArray(10, i)
                cate_Part = blog_CateArray(11, i)
                LastID = Int(ID)
                Loaded = True
                Exit Function
            End If
        Next
    End Function

End Class


'*******************************************
'  Tag Class
'*******************************************

Class Tag

    Private Sub Class_Initialize()
        If Not IsArray(Application(CookieName&"_blog_Tags")) Then Reload
    End Sub

    Private Sub Class_Terminate()

    End Sub

    Public Sub Reload
        Tags(2) '更新Tag缓存
    End Sub


    Public Function insert(tagName) '插入标签,返回ID号
        If checkTag(tagName) Then
            conn.Execute("update blog_tag set tag_count=tag_count+1 where tag_name='"&tagName&"'")
            insert = conn.Execute("select top 1 tag_id from blog_tag where tag_name='"&tagName&"'")(0)
        Else
            conn.Execute("insert into blog_tag (tag_name,tag_count) values ('"&tagName&"',1)")
            insert = conn.Execute("select top 1 tag_id from blog_tag order by tag_id desc")(0)
        End If
    End Function


    Public Function Remove(tagID) '清除标签
        If checkTagID(tagID) Then
            conn.Execute("update blog_tag set tag_count=tag_count-1 where tag_id="&tagID)
        End If
    End Function

    Public Function filterHTML(Str) '过滤标签
        If IsEmpty(Str) Or IsNull(Str) Or Len(Str) = 0 Then
            Exit Function
            filterHTML = Str
        Else
            Dim log_Tag, log_TagItem
            For Each log_TagItem IN Arr_Tags
                log_Tag = Split(log_TagItem, "||")
                Str = Replace(Str, "{"&log_Tag(0)&"}", "<a href=""default.asp?tag="&Server.URLEncode(log_Tag(1))&""">"&log_Tag(1)&"</a><a href=""http://technorati.com/tag/"&log_Tag(1)&""" rel=""tag"" style=""display:none"">"&log_Tag(1)&"</a> ")
            Next
            Dim re
            Set re = New RegExp
            re.IgnoreCase = True
            re.Global = True
            re.Pattern = "\{(\d)\}"
            Str = re.Replace(Str, "")
            filterHTML = Str
        End If
    End Function

    Public Function filterEdit(Str) '过滤标签进行编辑
        If IsEmpty(Str) Or IsNull(Str) Or Len(Str) = 0 Then
            Exit Function
            filterEdit = Str
        Else
            Dim log_Tag, log_TagItem
            For Each log_TagItem IN Arr_Tags
                log_Tag = Split(log_TagItem, "||")
                Str = Replace(Str, "{"&log_Tag(0)&"}", log_Tag(1)&" ")
            Next
            Dim re
            Set re = New RegExp
            re.IgnoreCase = True
            re.Global = True
            re.Pattern = "\{(\d)\}"
            Str = re.Replace(Str, "")
            filterEdit = Left(Str, Len(Str) -1)
        End If
    End Function

    Private Function checkTag(tagName) '检测是否存在此标签（根据名称）
        checkTag = False
        Dim log_Tag, log_TagItem
        For Each log_TagItem IN Arr_Tags
            log_Tag = Split(log_TagItem, "||")
            If LCase(log_Tag(1)) = LCase(tagName) Then
                checkTag = True
                Exit Function
            End If
        Next
    End Function

    Private Function checkTagID(tagID) '检测是否存在此标签（根据ID）
        checkTagID = False
        Dim log_Tag, log_TagItem
        For Each log_TagItem IN Arr_Tags
            log_Tag = Split(log_TagItem, "||")
            If Int(log_Tag(0)) = Int(tagID) Then
                checkTagID = True
                Exit Function
            End If
        Next
    End Function

    Public Function getTagID(tagName) '获得Tag的ID
        getTagID = 0
        Dim log_Tag, log_TagItem
        For Each log_TagItem IN Arr_Tags
            log_Tag = Split(log_TagItem, "||")
            If LCase(log_Tag(1)) = LCase(ClearHTML(tagName)) Then
                getTagID = log_Tag(0)
                Exit Function
            End If
        Next
    End Function

End Class

'*******************************************
'  Tag Gravatar
' 设置Gravatar头像信息
' <img alt="Gravatar Icon" src="http://www.gravatar.com/avatar/email md5?d=identicon&s=80&r=g"/>
'*******************************************
Class Gravatar
	Public Gravatar_d, Gravatar_s, Gravatar_r, Gravatar_EmailMd5
	Private Sub Class_Initialize()
        Gravatar_d = "identicon" ' 默认图片，如d=http%3A%2F%2Fexample.com%2Fimages%2Fexample.jpg(其中“%3A”代“:”，“%2F”代“/”)，也可以用三个特殊参数：identicons、monsterids、wavatars
        Gravatar_s = "40" ' 图片大小，单位是px，默认是80，可以取1~512之间的整数
        Gravatar_r = "g" ' 限制等级，默认为g，(G 普通级、PG 辅导级、R 和 X 为限制级)
        Gravatar_EmailMd5 = "" ' 邮箱的MD5值
    End Sub

    Private Sub Class_Terminate()

    End Sub
    
    Public Function outPut()
    	outPut = "http://www.gravatar.com/avatar/" & Gravatar_EmailMd5 & "?d=" & Gravatar_d & "&s=" & Gravatar_s & "&r=" & Gravatar_r
    End Function
    
End Class
%>
