<%
'==================================
'  日志类文件
'    更新时间: 2006-6-2
'==================================

'--------------------输出文件头------------------

Sub outCardHead(title) ' Output WML Head
    response.Write "<head><meta forua=""true"" http-equiv=""Cache-Control"" content=""max-age=0"" /></head>"
    response.Write "<card id=""MainCard"" title="""&toUnicode(title)&""">"
    If Not blog_wap Then
        response.Write "<p>"&toUnicode("Blog不允许通过wap访问.")&"</p>"
        response.Write "</card>"
        response.Write "</wml>"
        response.End
    End If
End Sub

'--------------------输出wap 首页----------------

Sub outDefault 'Output Default
    outCardHead ("欢迎光临")
    response.Write "<p><a href=""wap.asp"">"&toUnicode(HTMLDecode(SiteName))&"</a><br/>&nbsp;</p>"
    response.Write "<p><a href=""wap.asp?do=showLog"">"&toUnicode("最新日志")&"</a></p>"
    response.Write "<p><a href=""wap.asp?do=showComment"">"&toUnicode("最新评论")&"</a></p>"
    Dim Arr_Category, Category_Len, i,Category_code
    CategoryList(3)
    Category_code = ""
    Arr_Category = Application(CookieName&"_blog_Category")
    If UBound(Arr_Category, 1) = 0 Then Exit Sub
    Category_Len = UBound(Arr_Category, 2)
    For i = 0 To Category_Len
        If Not Arr_Category(4, i) Then
            If CBool(Arr_Category(10, i)) Then
                If stat_ShowHiddenCate Or stat_Admin Then Category_code = Category_code&("&nbsp;- <a href=""wap.asp?do=showLog&amp;cateID="&Arr_Category(0, i)&""">"&toUnicode(Arr_Category(1, i))&"</a>&nbsp;("&Arr_Category(7, i)&")<br/>")
            Else
                Category_code = Category_code&("&nbsp;- <a href=""wap.asp?do=showLog&amp;cateID="&Arr_Category(0, i)&""">"&toUnicode(Arr_Category(1, i))&"</a>&nbsp;("&Arr_Category(7, i)&")<br/>")
            End If
        End If
    Next
    response.Write "<p>"&Category_code&"</p>"
    response.Write "<p><a href=""wap.asp?do=showTag"">"&toUnicode("标签云集")&"</a></p>"
    outControl
    outCardFoot
    Conn.Close
    Set Conn = Nothing
End Sub

'--------------------输出最新文章----------------

Sub outLog
    outCardHead ("欢迎光临")
    Dim wPageSize
    wPageSize = blog_wapNum
    response.Write "<p><a href=""wap.asp"">"&toUnicode(HTMLDecode(SiteName))&"</a><br/>&nbsp;</p>"
    Dim dbRow, i, webLog, strSQL, CanRead, Log_Num
    Dim CT, ViewTag, getCate
    CT = "最新日志"
    ViewTag = checkstr(Request.QueryString("tag"))
    Set getCate = New Category
    If IsInteger(cateID) = True Then
        getCate.load(cateID)
        CT = "分类: "&getCate.cate_Name
        If getCate.cate_Secret Then
            If Not stat_ShowHiddenCate And Not stat_Admin Then
                response.Write "<p>"&toUnicode("找不到任何日志")&"</p>"
                outCardFoot
                Exit Sub
            End If
        End If
    End If

    If Len(ViewTag)>0 Then
        Dim getTag, getTID
        Set getTag = New tag
        getTID = getTag.getTagID(ViewTag)
        If getTID<>0 Then
            SQLFiltrate = SQLFiltrate & " log_tag LIKE '%{"&getTID&"}%' AND "
            Url_Add = Url_Add & "tag="&Server.URLEncode(ViewTag)&"&amp;"
            CT = "Tag: "&ViewTag
        End If
        Set getTag = Nothing
    End If
    strSQL = "log_ID,log_Title,log_CateID,log_Author,log_PostTime,log_IsShow,log_IsTop"
    If stat_ShowHiddenCate Or stat_Admin Then
        SQL = "SELECT "&strSQL&" FROM blog_Content "&SQLFiltrate&" log_IsDraft=false ORDER BY log_PostTime DESC"
    Else
        SQL = "SELECT "&strSQL&" FROM blog_Content As T,blog_Category As C "&SQLFiltrate&" T.log_CateID=C.cate_ID and C.cate_Secret=false and log_IsDraft=false ORDER BY log_PostTime DESC"
    End If

    Set webLog = Server.CreateObject("Adodb.Recordset")
    webLog.Open SQL, Conn, 1, 1
    SQLQueryNums = SQLQueryNums + 1

    If webLog.EOF Or webLog.BOF Then
        response.Write "<p>"&toUnicode("找不到任何日志")&"</p>"
        outCardFoot
        ReDim dbRow(0, 0)
        Exit Sub
    Else
        webLog.PageSize = wPageSize
        webLog.AbsolutePage = CurPage
        Log_Num = webLog.RecordCount
        dbRow = webLog.getrows(wPageSize)
    End If
    webLog.Close
    Set webLog = Nothing
    Conn.Close
    Set Conn = Nothing

    If UBound(dbRow, 1)<>0 Then
        response.Write "<p>"&toUnicode(CT)&"</p>"
        For i = 0 To UBound(dbRow, 2)
            CanRead = False
            If stat_Admin = True Then CanRead = True
            If dbRow(5, i) Then CanRead = True
            If dbRow(5, i) = False And dbRow(2, i) = memName Then CanRead = True
            If CanRead Then
                response.Write "<p>&nbsp;&nbsp;"&(i + wPageSize * (CurPage -1) + 1)&". <a href=""wap.asp?do=showLogDetail&amp;id="&dbRow(0, i)&""">"&toUnicode(dbRow(1, i))&"</a>"
                If dbRow(5, i) = False Then response.Write toUnicode(" [私]")
                response.Write "</p>"
            Else
                response.Write "<p>&nbsp;&nbsp;"&(i + wPageSize * (CurPage -1) + 1)&". "&toUnicode("[私密日志]")&"</p>"
            End If
        Next
        Call getPage (Log_Num, CurPage, wPageSize, Url_Add)
    End If

    outControl
    outCardFoot
End Sub

Sub outLogDetail
    outCardHead ("欢迎光临")
    response.Write "<p><a href=""wap.asp"">"&toUnicode(HTMLDecode(SiteName))&"</a><br/>&nbsp;</p>"
    Dim logID, lArticle, loadArticle
    logID = CLng(request.QueryString("id"))

    Set lArticle = New logArticle
    loadArticle = lArticle.getLog(logID)
    If loadArticle(0) = 0 Then
        Dim getCate
        Set getCate = New Category
        getCate.load(Int(lArticle.categoryID))
        If (lArticle.logAuthor = memName And lArticle.logIsShow = False) Or stat_Admin Or lArticle.logIsShow = True Then
        Else
            response.Write "<p>"&toUnicode("没有权限查看该日志.")&"</p>"
            outCardFoot
        End If

        If (Not getCate.cate_Secret) Or (lArticle.logAuthor = memName And getCate.cate_Secret) Or stat_Admin Or (getCate.cate_Secret And stat_ShowHiddenCate) Then
        Else
            response.Write "<p>"&toUnicode("没有权限查看该日志.")&"</p>"
            outCardFoot
        End If

        response.Write "<p>"&"<b>"&toUnicode("标题:")&"</b> "&toUnicode(lArticle.logTitle)
        If lArticle.logIsShow = False Then response.Write toUnicode(" [私]")
        response.Write "</p>"
        response.Write "<p>"&"<b>"&toUnicode("作者:")&"</b> "&toUnicode(lArticle.logAuthor)&"</p>"
        response.Write "<p>"&"<b>"&toUnicode("日期:")&"</b> "&toUnicode(DateToStr(lArticle.logPubTime, "Y-m-d H:I A"))&"</p>"
        response.Write "<p>"&"<b>"&toUnicode("分类:")&"</b> <a href=""wap.asp?do=showLog&amp;cateID="&lArticle.categoryID&""">"&toUnicode(getCate.cate_Name)&"</a></p>"
        If lArticle.logEditType = 0 Then
            response.Write "<p>"&"<b>"&toUnicode("内容:")&"</b> "&HtmlToText(UnCheckStr(lArticle.logMessage))&"</p>"
        Else
            response.Write "<p>"&"<b>"&toUnicode("内容:")&"</b> "&HtmlToText(UBBCode(HTMLEncode(lArticle.logMessage), 0, 0, 0, 1, 1))&"</p>"
        End If
        response.Write "<p> + <a href=""#CommentCard"">"&toUnicode("查看当前日志评论")&"</a> ("&lArticle.logCommentCount&")</p>"
    Else
        response.Write "<p>"&toUnicode(loadArticle(1))&"</p><p><a href=""wap.asp"">"&toUnicode("返回首页")&"</a></p>"
        outCardFoot
        Exit Sub
    End If
    outControl
    outCardFoot

    If stat_CommentAdd Then
        response.Write "<card id=""postCommentCard"">"
        response.Write "<p><b>"&toUnicode("标题:")&"</b> <a href=""#MainCard"">"&toUnicode(lArticle.logTitle)&"</a><br/></p>"
        response.Write "<p><br/><b>发表评论:</b></p><p>"
        If memName<>Empty Then
            response.Write toUnicode("昵称:")&"<b><i>"&memName&"</i></b><br/>"
        Else
            response.Write toUnicode("昵称:")&"<input emptyok=""false"" name=""userName"" size=""10"" maxlength=""20"" type=""text"" value="""" /><br/>"&toUnicode("密码:")&"<input emptyok=""false"" name=""password"" size=""10"" maxlength=""20"" type=""password"" value="""" /><br/>"
        End If
        response.Write toUnicode("评论内容: ")&"<br/> <input emptyok=""false"" name=""message""  maxlength="""&blog_commLength&""" format=""false"" value="""" /><br/>"
        response.Write "<anchor>"&toUnicode("发表")&"<go href=""wap.asp?do=postComment"" method=""post"">"
        If memName<>Empty Then
            response.Write "<postfield name=""userName"" value="""&memName&""" />"
        Else
            response.Write "<postfield name=""userName"" value=""$(userName)"" />"
        End If
        response.Write "<postfield name=""password"" value=""$(password)"" />"
		response.Write "<postfield name=""validate"" value=""0000"" />"
        response.Write "<postfield name=""message"" value=""$(message)"" />"
        response.Write "<postfield name=""id"" value="""&logID&""" />"
        response.Write "</go></anchor>"
        response.Write "&nbsp;<a href=""#MainCard"">"&toUnicode("返回")&"</a></p>"
    Else
        response.Write "<card id=""postCommentCard"">"
        response.Write "<p><b>"&toUnicode("标题:")&"</b> <a href=""#MainCard"">"&toUnicode(lArticle.logTitle)&"</a></p>"
        response.Write "<p><br/>你没有权限发表评论</p>"
    End If
    outCardFoot

    response.Write "<card id=""CommentCard"">"
    Dim blog_Comment, comDesc, dbRow, i
    If lArticle.logCommentOrder Then comDesc = "Desc" Else comDesc = "Asc" End If
    SQL = "SELECT comm_ID,comm_Content,comm_Author,comm_PostTime,comm_DisSM,comm_DisUBB,comm_DisIMG,comm_AutoURL,comm_PostIP,comm_AutoKEY FROM blog_Comment WHERE blog_ID="&logID&" ORDER BY comm_PostTime "&comDesc
    Set blog_Comment = conn.Execute (SQL)
    If blog_Comment.EOF And blog_Comment.BOF Then
        response.Write "<p>"&toUnicode("暂无评论")&"</p><p><a href=""#MainCard"">"&toUnicode("返回")&"</a></p>"
        outCardFoot
        Exit Sub
    Else
        dbRow = blog_Comment.getrows()
    End If
    response.Write "<p><b>"&toUnicode("标题:")&"</b> <a href=""#MainCard"">"&toUnicode(lArticle.logTitle)&"</a></p>"
    response.Write"<p><b>"&toUnicode("评论内容:")&"</b></p>"

    If UBound(dbRow, 1)<>0 Then
        For i = 0 To UBound(dbRow, 2)
            response.Write"<p>== <b>"&toUnicode(dbRow(2, i))&"</b> <small>"&DateToStr(dbRow(3, i), "Y-m-d H:I A")&"</small> ----</p>"
            response.Write"<p>"&HtmlToText(UBBCode(HtmlEncode(dbRow(1, i)), dbRow(4, i), dbRow(5, i), dbRow(6, i), dbRow(7, i), dbRow(9, i)))&" </p>"
            response.Write"<p> </p>"
        Next
    End If
    outCardFoot

    blog_Comment.Close
    Set blog_Comment = Nothing
    Set lArticle = Nothing
    Conn.Close
    Set Conn = Nothing
End Sub


Sub Articleadd() '发表日志
    outCardHead ("欢迎光临")
    response.Write "<p><a href=""wap.asp"">"&toUnicode(HTMLDecode(SiteName))&"</a><br/>&nbsp;</p>"
    If stat_AddAll = True Or stat_Add = True Then
        response.Write "<p><br/><b>发表日志:</b></p><p>"
%>
        分类：<select name="log_CateID" id="log_CateID">
          <option value="" selected="selected" style="color:#333">请选择分类</option>
          <%
Dim Arr_Category, Category_Len, i
Arr_Category = Application(CookieName&"_blog_Category")
If UBound(Arr_Category, 1) = 0 Then Exit Sub
Category_Len = UBound(Arr_Category, 2)

For i = 0 To Category_Len
    If Not Arr_Category(4, i) Then
        If CBool(Arr_Category(10, i)) Then
            If stat_ShowHiddenCate And stat_Admin Then Response.Write("<option value='"&Arr_Category(0, i)&"'>"&Arr_Category(1, i)&"&nbsp;["&Arr_Category(7, i)&"]</option>")
        Else
            Response.Write("<option value='"&Arr_Category(0, i)&"'>"&Arr_Category(1, i)&"["&Arr_Category(7, i)&"]</option>")
        End If
    End If
Next

%>
        </select><br/>
<%
response.Write toUnicode("标题: ")&"<input emptyok=""false"" name=""Artile_Title"" size=""10"" maxlength=""50"" type=""text"" value="""" /><br/>"
response.Write toUnicode("tags: ")&"<input emptyok=""false"" name=""tags"" size=""10"" maxlength=""50"" type=""text"" value="""" /><br/>"
response.Write toUnicode("内容: ")&"<input emptyok=""false"" name=""Article_content"" size=""10""  maxlength=""1000"" value="""" /><br/>"
response.Write "<anchor>"&toUnicode("发表")&"<go href=""wap.asp?do=postArtile"" method=""post"">"
response.Write "<postfield name=""log_CateID"" value=""$(log_CateID)"" />"
response.Write "<postfield name=""Artile_Title"" value=""$(Artile_Title)"" />"
response.Write "<postfield name=""tags"" value=""$(tags)"" />"
response.Write "<postfield name=""Article_content"" value=""$(Article_content)"" />"
response.Write "<postfield name=""username"" value="""&memName&""" />"
response.Write "</go></anchor>"
response.Write "&nbsp;<a href=""#MainCard"">"&toUnicode("返回")&"</a></p>"
Else
    response.Write "<p><br/>你没有权限发表日志</p>"
End If
outCardFoot
End Sub

'--------------------输出导航----------------

Sub getPage (lN, cP, wP, uA)
    Dim getPages, URL
    If lN Mod CInt(wP) = 0 Then
        getPages = Int(lN / wP)
    Else
        getPages = Int(lN / wP) + 1
    End If
    URL = Request.ServerVariables("Script_Name")&uA
    response.Write "<p> "
    If Int(cP)<>1 Then
        response.Write "<a href="""&Url&"do=showLog&amp;page="&(cP -1)&""">"& toUnicode("上一页") &"</a>" & " "
    Else
        response.Write toUnicode("上一页") & " "
    End If

    If getPages<>Int(cP) Then
        response.Write "<a href="""&Url&"do=showLog&amp;page="&(cP + 1)&""">"& toUnicode("下一页") &"</a>"
    Else
        response.Write toUnicode("下一页")
    End If
    response.Write "</p>"
End Sub

'--------------------输出最新评论----------------

Sub outComment
    outCardHead ("欢迎光临")
    response.Write "<p><a href=""wap.asp"">"&toUnicode(HTMLDecode(SiteName))&"</a><br/>&nbsp;</p>"
    Dim blog_Comment, comDesc, dbRow, i
    SQL = "SELECT top 10 comm_ID,comm_Content,comm_Author,comm_PostTime,comm_DisSM,comm_DisUBB,comm_DisIMG,comm_AutoURL,comm_AutoKEY,blog_ID FROM blog_Comment ORDER BY comm_PostTime desc"
    Set blog_Comment = conn.Execute (SQL)
    If blog_Comment.EOF And blog_Comment.BOF Then
        response.Write "<p>"&toUnicode("暂无评论")&"</p><p><a href=""wap.asp"">"&toUnicode("返回首页")&"</a></p>"
        outCardFoot
        Exit Sub
    Else
        dbRow = blog_Comment.getrows()
    End If
    blog_Comment.Close
    Set blog_Comment = Nothing
    Conn.Close
    Set Conn = Nothing

    If UBound(dbRow, 1)<>0 Then
        For i = 0 To UBound(dbRow, 2)
            response.Write"<p>== <b>"&toUnicode(dbRow(2, i))&"</b> <small>"&DateToStr(dbRow(3, i), "Y-m-d H:I A")&"</small> ----</p>"
            response.Write"<p>"&HtmlToText(UBBCode(HtmlEncode(dbRow(1, i)), dbRow(4, i), dbRow(5, i), dbRow(6, i), dbRow(7, i), dbRow(8, i)))&" </p>"
            response.Write"<p> + <a href=""wap.asp?do=showLogDetail&amp;id="&dbRow(9, i)&"#CommentCard"">"&toUnicode("查看完整评论")&"</a></p>"
            response.Write"<p> </p>"
        Next
    End If
    outCardFoot
End Sub

'--------------------输出Tag列表----------------

Sub outTagList
    outCardHead ("标签云集")
    response.Write "<p><a href=""wap.asp"">"&toUnicode(HTMLDecode(SiteName))&"</a><br/>&nbsp;</p><p>"
    Dim log_Tag, log_TagItem
    For Each log_TagItem IN Arr_Tags
        log_Tag = Split(log_TagItem, "||")

%>
	          &nbsp;- <a href="wap.asp?do=showLog&amp;tag=<%=Server.URLEncode(log_Tag(1))%>"><%=toUnicode(log_Tag(1))%></a>&nbsp;(<%=log_Tag(2)%>)<br/>
			<%
Next
response.Write "</p><p><br/><a href=""wap.asp"">"&toUnicode("返回首页")&"</a></p>"
outControl
outCardFoot
Conn.Close
Set Conn = Nothing
End Sub

'--------------------登录框----------------

Sub outLogin
    outCardHead ("登录Blog")
    response.Write "<p><a href=""wap.asp"">"&toUnicode(HTMLDecode(SiteName))&"</a><br/>&nbsp;</p><p>"
    response.Write toUnicode("用户名: ")&"<input emptyok=""false"" name=""userName"" size=""10"" maxlength=""20"" type=""text"" value="""" /><br/>"
    response.Write toUnicode("密　码: ")&"<input emptyok=""false"" name=""Password"" size=""10"" maxlength=""32"" type=""password"" /><br/>"
    response.Write "<anchor>"&toUnicode("登录")&"<go href=""wap.asp?do=CheckUser"" method=""post"">"
    response.Write "<postfield name=""userName"" value=""$(userName)"" />"
    response.Write "<postfield name=""Password"" value=""$(Password)"" />"
    response.Write "<postfield name=""validate"" value=""0000"" />"
    response.Write "</go></anchor>"
    response.Write "&nbsp;<a href=""wap.asp"">"&toUnicode("返回")&"</a></p>"
    outCardFoot
End Sub

'--------------------用户登录----------------

Sub outLoginUser
    Dim loginUser
    response.Write "<head><meta forua=""true"" http-equiv=""Cache-Control"" content=""max-age=0"" /></head>"
    response.Write "<card id=""splashscreen"" ontimer=""wap.asp"" title="""&toUnicode("登入信息")&""">"

    If Not blog_wapLogin Then
        response.Write "<timer value=""5000""/>"
        response.Write "<p>"&toUnicode("Blog不允许wap登录")&"</p>"
        response.Write "<p><a href=""wap.asp"">"&toUnicode("返回首页")&"</a></p>"
    Else
        Session("GetCode") = Request.Form("validate")
        loginUser = login(Request.Form("userName"), Request.Form("Password"))
        If Not loginUser(3) Then
            response.Write "<timer value=""5000""/>"
            response.Write "<p>"&toUnicode("登入失败！")&"</p>"
            response.Write "<p><br/><a href=""wap.asp?do=Login"">"&toUnicode("重新登入")&"</a><br/><a href=""wap.asp"">"&toUnicode("返回首页")&"</a></p>"
        Else
            response.Write "<timer value=""30""/>"
            response.Write "<p>"&toUnicode("登入成功！")&toUnicode(Request.Form("userName"))&toUnicode("欢迎你回来.")&"</p>"
            response.Write "<p>"&toUnicode("3秒后自动返回首页返回首页")&"</p>"
            response.Write "<p><br/><a href=""wap.asp"">"&toUnicode("返回首页")&"</a></p>"
        End If
    End If
    outCardFoot
    Conn.Close
    Set Conn = Nothing
End Sub

'--------------------用户登出----------------

Sub outLogout
    logout(True)
    response.Write "<head><meta forua=""true"" http-equiv=""Cache-Control"" content=""max-age=0"" /></head>"
    response.Write "<card id=""splashscreen"" ontimer=""wap.asp"" title="""&toUnicode("登出信息")&""">"
    response.Write "<timer value=""30""/>"
    response.Write "<p>"&toUnicode("欢迎你已成功退出.")&"</p>"
    response.Write "<p>"&toUnicode("3秒后自动返回首页返回首页")&"</p>"
    response.Write "<p><br/><a href=""wap.asp"">"&toUnicode("返回首页")&"</a></p>"
    response.Write "</card>"
    Conn.Close
    Set Conn = Nothing
End Sub


'--------------------控制面板----------------

Sub outControl
    response.Write "<p>&nbsp;<br/>"
    If memName<>Empty Then response.Write "<b>"&toUnicode(memName)&"</b>"&toUnicode("，欢迎你!")&"<br/>"
    If stat_AddAll = True Or stat_Add = True Then response.Write "<a href=""wap.asp?do=Articleadd"">"&toUnicode("发表新日志")&"</a>"
    If request.QueryString("do") = "showLogDetail" And stat_CommentAdd And blog_wapComment Then response.Write "<br/><a href=""#postCommentCard"">"&toUnicode("发表评论")&"</a>"
    If memName<>Empty Then
        response.Write "<br/><a href=""wap.asp?do=Logout"">"&toUnicode("登出")&"</a>"
    Else
        If blog_wapLogin Then response.Write "<br/><a href=""wap.asp?do=Login"">"&toUnicode("登录")&"</a>"
    End If
    response.Write "</p>"
End Sub

'--------------------输出文件未----------------

Sub outCardFoot 'Output Foot
    response.Write "<p><br/>"&toUnicode("────────────")&"</p>"
    response.Write "<p><a href=""wap.asp"">"&toUnicode(HTMLDecode(SiteName))&"</a></p>"
    response.Write "<p><a href=""http://www.pjhome.net/wap.asp"">PJBlog3&nbsp;v"&blog_version&"</a>&nbsp;Inside.</p>"
    response.Write "<p>Processed&nbsp;In&nbsp;"&FormatNumber(Timer() - StartTime, 3, -1)&"&nbsp;ms</p>"
    response.Write "<do type=""prev"" label="""&toUnicode("返回")&"""><prev/></do>"
    response.Write "</card>"
End Sub


Sub outPostComment
    Dim postComment
    postComment = wap_CommentPost

    response.Write "<head><meta forua=""true"" http-equiv=""Cache-Control"" content=""max-age=0"" /></head>"
    response.Write "<card id=""splashscreen"" ontimer=""wap.asp"" title="""&toUnicode("发表评论")&""">"
    response.Write "<p>"&postComment&"</p>"

    outCardFoot
    Conn.Close
    Set Conn = Nothing
End Sub

Function wap_CommentPost
    wap_CommentPost = ""
    Dim username, post_logID, post_Message, LastMSG, FlowControl, ReInfo, password
    username = Trim(CheckStr(request.Form("userName")))
    password = Trim(CheckStr(request.Form("password")))
    post_logID = CLng(CheckStr(request.Form("id")))
    post_Message = CheckStr(request.Form("message"))

    Set LastMSG = conn.Execute("select top 1 comm_Content from blog_Comment where blog_ID="&post_logID&" order by comm_ID desc")
    If LastMSG.EOF Then
        FlowControl = False
    Else
        If LastMSG("comm_Content") = post_Message Then FlowControl = True
    End If

    If filterSpam(post_Message, "spam.xml") And stat_Admin = False Then
        wap_CommentPost = "<b>"&toUnicode("评论中包含被屏蔽的字符")&"</b><br/><a href=""wap.asp?do=showLogDetail&amp;id="&post_logID&"#postCommentCard"">"&toUnicode("返回")&"</a>"
        Exit Function
    End If

    If FlowControl Then
        wap_CommentPost = "<b>"&toUnicode("禁止恶意灌水！")&"</b><br/><a href=""wap.asp?do=showLogDetail&amp;id="&post_logID&"#postCommentCard"">"&toUnicode("返回")&"</a>"
        Exit Function
    End If

    If DateDiff("s", Request.Cookies(CookieName)("memLastPost"), Now())<blog_commTimerout Then
        wap_CommentPost = "<b>"&toUnicode("发言太快,请 "&blog_commTimerout&" 秒后再发表评论")&"</b><br/><a href=""wap.asp?do=showLogDetail&amp;id="&post_logID&"#postCommentCard"">"&toUnicode("返回")&"</a>"
        Exit Function
    End If

    If Len(username)<1 Then
        wap_CommentPost = "<b>"&toUnicode("请输入你的昵称.")&"</b><br/><a href=""wap.asp?do=showLogDetail&amp;id="&post_logID&"#postCommentCard"">"&toUnicode("返回")&"</a>"
        Exit Function
    End If

    If IsValidUserName(username) = False Then
        wap_CommentPost = "<b>"&toUnicode("非法用户名！")&"</b><br/><a href=""wap.asp?do=showLogDetail&amp;id="&post_logID&"#postCommentCard"">"&toUnicode("返回")&"</a>"
        Exit Function
    End If

    If Not stat_CommentAdd Then
        wap_CommentPost = "<b>"&toUnicode("你没有权限发表评论")&"</b><br/><a href=""wap.asp?do=showLogDetail&amp;id="&post_logID&"#postCommentCard"">"&toUnicode("返回")&"</a>"
        Exit Function
    End If

    If Not blog_wapComment Then
        wap_CommentPost = "<b>"&toUnicode("Blog不开放Wap发表评论")&"</b><br/><a href=""wap.asp?do=showLogDetail&amp;id="&post_logID&"#postCommentCard"">"&toUnicode("返回")&"</a>"
        Exit Function
    End If

    If Len(post_Message)<1 Then
        wap_CommentPost = "<b>"&toUnicode("不允许发表空评论")&"</b><br/><a href=""wap.asp?do=showLogDetail&amp;id="&post_logID&"#postCommentCard"">"&toUnicode("返回")&"</a>"
        Exit Function
    End If

    If Len(post_Message)>blog_commLength Then
        wap_CommentPost = "<b>"&toUnicode("评论超过最大字数限制")&"</b><br/><a href=""wap.asp?do=showLogDetail&amp;id="&post_logID&"#postCommentCard"">"&toUnicode("返回")&"</a>"
        Exit Function
    End If

    Dim checkMem
    If memName = Empty Then'匿名评论
        If Len(password)>0 Then
            Dim loginUser
			Session("GetCode") = Request.Form("validate")
            loginUser = login(Request.Form("username"), Request.Form("password"))
			If Not loginUser(3) Then
				wap_CommentPost = "<b>"&toUnicode("登录失败，请检查用户名和密码")&"</b><br/><a href=""wap.asp?do=showLogDetail&amp;id="&post_logID&"#postCommentCard"">"&toUnicode("返回")&"</a>"
                Exit Function
            End If
        Else
            Set checkMem = Conn.Execute("select top 1 mem_id from blog_Member where mem_Name='"&username&"'")
            If Not checkMem.EOF Then
				wap_CommentPost = "<b>"&toUnicode("该用户已经存在，请登录后或输入用户密码再发表评论")&"</b><br/><a href=""wap.asp?do=showLogDetail&amp;id="&post_logID&"#postCommentCard"">"&toUnicode("返回")&"</a>"
                Exit Function
            End If
        End If
    Else
 			If Not request.Cookies(CookieName)("memName") = username Then
                ReInfo(0) = "评论发表错误信息"
				wap_CommentPost = "<b>"&toUnicode("请输入正确的用户名")&"</b><br/><a href=""wap.asp?do=showLogDetail&amp;id="&post_logID&"#postCommentCard"">"&toUnicode("返回")&"</a>"
                Exit Function
            End If
    End If

    If Conn.Execute("select log_DisComment from blog_Content where log_ID="&post_logID)(0) Then
        wap_CommentPost = "<b>"&toUnicode("该日志不允许发表任何评论")&"</b><br/><a href=""wap.asp?do=showLogDetail&amp;id="&post_logID&"#postCommentCard"">"&toUnicode("返回")&"</a>"
        Exit Function
    End If

    Dim post_DisSM, post_DisURL, post_DisKEY, post_disImg, post_DisUBB
    post_DisSM = 0
    post_DisURL = 1
    post_DisKEY = 1

    post_disImg = 1
    post_DisUBB = 0

    Dim AddComm
    AddComm = Array(Array("blog_ID", post_logID), Array("comm_Content", post_Message), Array("comm_Author", username), Array("comm_DisSM", post_DisSM), Array("comm_DisUBB", post_DisUBB), Array("comm_DisIMG", post_disImg), Array("comm_AutoURL", post_DisURL), Array("comm_PostIP", getIP), Array("comm_AutoKEY", post_DisKEY))
    DBQuest "blog_Comment", AddComm, "insert"
    'Conn.ExeCute("INSERT INTO blog_Comment(blog_ID,comm_Content,comm_Author,comm_DisSM,comm_DisUBB,comm_DisIMG,comm_AutoURL,comm_PostIP,comm_AutoKEY) VALUES ("&post_logID&",'"&post_Message&"','"&username&"',"&post_DisSM&","&post_DisUBB&","&post_disImg&","&post_DisURL&",'"&getIP()&"',"&post_DisKEY&")")
    Conn.Execute("update blog_Content set log_CommNums=log_CommNums+1 where log_ID="&post_logID)
    Conn.Execute("update blog_Info set blog_CommNums=blog_CommNums+1")
    Response.Cookies(CookieName)("memLastpost") = DateToStr(now(),"Y-m-d H:I:S")
    getInfo(2)
    NewComment(2)

    If memName<>Empty Then conn.Execute("update blog_Member set mem_PostComms=mem_PostComms+1 where mem_Name='"&memName&"'")
    wap_CommentPost = "<b>"&toUnicode("你成功地对该日志发表了评论")&"</b><br/><a href=""wap.asp?do=showLogDetail&amp;id="&post_logID&"#CommentCard"">"&toUnicode("返回")&"</a>"
    PostArticle post_logID, False
    call newEtag
End Function

Sub outwap_AritclePost
    Dim postAritcle
    postAritcle = wap_AritclePost

    response.Write "<head><meta forua=""true"" http-equiv=""Cache-Control"" content=""max-age=0"" /></head>"
    response.Write "<card id=""splashscreen"" ontimer=""wap.asp"" title="""&toUnicode("发表日志")&""">"
    response.Write "<p>"&postAritcle&"</p>"

    outCardFoot
    Conn.Close
    Set Conn = Nothing
End Sub

Function wap_AritclePost '提交日志
    wap_AritclePost = ""
    If stat_AddAll = True Or stat_Add = True Then
        Dim lArticle, postLog
        Set lArticle = New logArticle
        lArticle.categoryID = request.Form("log_CateID")
        lArticle.logTitle = request.Form("Artile_Title")
        lArticle.logAuthor = memName
        lArticle.logIsDraft = request.Form("log_IsDraft")
        lArticle.logMessage = request.Form("Article_content")
        lArticle.logTags = request.Form("tags")
        postLog = lArticle.postLog
        Set lArticle = Nothing
        If postLog(0)<0 Then
            wap_AritclePost = "<b>"&toUnicode("发表日志失败")&"</b><br/><a href=""wap.asp"">"&toUnicode("返回")&"</a>"
        End If
        If postLog(0)>= 0 Then
            wap_AritclePost = "<b>"&toUnicode("发表日志成功")&"</b><br/><a href=""wap.asp"">"&toUnicode("返回")&"</a>"
            Dim Article_code
            If Session(CookieName&"_LastDo") = "DelArticle" Or Session(CookieName&"_LastDo") = "AddArticle" Or Session(CookieName&"_LastDo") = "EditArticle" Then NewArticle(2)
            Article_code = NewArticle(0)
            side_html_default = Replace(side_html_default, "<$NewLog$>", Article_code)
            side_html = Replace(side_html, "<$NewLog$>", Article_code)
        End If
    Else
        response.Write "<p><br/>你没有权限发表日志</p>"
        Exit Function
    End If
End Function

Function toUnicode(Str) 'To Unicode
    Dim i, a
    For i = 1 To Len (Str)
        a = Mid(Str, i, 1)
        toUnicode = toUnicode & "&#x" & Hex(Ascw(a)) & ";"
    Next
End Function

Function toUnicodeJS(Str) 'To Unicode
    Dim i, a, ignoreStr
    ignoreStr = "1234567890`~!@#$%^&*()_+=-{}][|\"":;'?><,./qwertyuioplkjhgfdsazxcvbnmQWERTYUIOPLKJHGFDSAZXCVBNM"
    For i = 1 To Len (Str)
        a = Mid(Str, i, 1)
        If InStr(ignoreStr, a)>0 Then
            toUnicodeJS = toUnicodeJS & a
        Else
            toUnicodeJS = toUnicodeJS & "&#x" & Hex(Ascw(a)) & ";"
        End If
    Next
End Function

Function NewArticle(ByVal action)
    Dim blog_Article
    If Not IsArray(Application(CookieName&"_blog_Article")) Or action = 2 Then
        Dim book_Articles, book_Article
        Set book_Articles = Conn.Execute("SELECT top 10 C.log_ID,C.log_Author,C.log_IsShow,C.log_PostTime,C.log_title,L.cate_ID,L.cate_Secret FROM blog_Content AS C,blog_Category AS L where L.cate_ID=C.log_CateID and L.cate_Secret=false and C.log_IsDraft=false order by log_ID Desc")
        SQLQueryNums = SQLQueryNums + 1
        TempVar = ""
        Do While Not book_Articles.EOF
            If book_Articles("cate_Secret") Then
                book_Article = book_Article&TempVar&book_Articles("log_ID")&"|,|"&book_Articles("log_Author")&"|,|"&book_Articles("log_PostTime")&"|,|"&"[私密分类日志]"
            ElseIf book_Articles("log_IsShow") Then
                book_Article = book_Article&TempVar&book_Articles("log_ID")&"|,|"&book_Articles("log_Author")&"|,|"&book_Articles("log_PostTime")&"|,|"&book_Articles("log_title")
            Else
                book_Article = book_Article&TempVar&book_Articles("log_ID")&"|,|"&book_Articles("log_Author")&"|,|"&book_Articles("log_PostTime")&"|,|"&"[私密日志]"
            End If
            TempVar = "|$|"
            book_Articles.MoveNext
        Loop
        Set book_Articles = Nothing
        blog_Article = Split(book_Article, "|$|")
        Application.Lock
        Application(CookieName&"_blog_Article") = blog_Article
        Application.UnLock
    Else
        blog_Article = Application(CookieName&"_blog_Article")
    End If

    If action<>2 Then
        Dim Article_Items, Article_Item
        For Each Article_Items IN blog_Article
            Article_Item = Split(Article_Items, "|,|")
            NewArticle = NewArticle&"<a class=""sideA"" href=""default.asp?id="&Article_Item(0)&""" title="""&Article_Item(1)&" 于 "&Article_Item(2)&" 发表该日志"&Chr(10)&CCEncode(CutStr(Article_Item(3), 25))&""">"&CCEncode(CutStr(Article_Item(3), 25))&"</a>"
        Next
    End If
End Function

%>

<script runat="server" Language="javascript">
	function HtmlToText(str) {
	    //filter HTMLToUBBAndText
		str = str.replace(/\r/g,"");
		str = str.replace(/on(load|click|dbclick|mouseover|mousedown|mouseup)="[^"]+"/ig,"");
		if (blog_wapImg) str = str.replace(/<img[^>]+src="([^"]+)"[^>]*>/ig,"\n[img]$1[/img]\n");
		if (blog_wapHTML) {
			str = str.replace(/<a[^>]+href="([^"]+)"[^>]*>(.*?)<\/a>/ig,"[url=$1]$2[/url]");
			str = str.replace(/<([\/]?)b>/ig,"[$1b]");
			str = str.replace(/<([\/]?)strong>/ig,"[$1b]");
			str = str.replace(/<p>(.*?)<\/p>/ig,"$1<br/>");
			str = str.replace(/<([\/]?)i>/ig,"[$1i]");
			str = str.replace(/<([\/]?)u>/ig,"[$1u]");
		}
		str = str.replace(/<br\/>/ig,"\n");
		str = str.replace(/<[^>]*?>/g,"");
		str = str.replace(/&/g,"&amp;");
		str = str.replace(/&amp;(.*?[^\s]);/g,"&$1;");
		str = str.replace(/\n+/g,"<br/>");

		if (blog_wapHTML) {
		//filter UBBToText
			str=str.replace(/(\[i\])(.*?)(\[\/i\])/ig,"<i>$2</i>")
			str=str.replace(/(\[u\])(.*?)(\[\/u\])/ig,"<u>$2</u>")
			str=str.replace(/(\[b\])(.*?)(\[\/b\])/ig,"<b>$2</b>")		
			if (blog_wapURL) 
			  str=str.replace(/(\[url=(.[^\[]*)\])(.*?)(\[\/url\])/ig,"<a href=\"$2\">$3</a>");
			 else
			  str=str.replace(/(\[url=(.[^\[]*)\])(.*?)(\[\/url\])/ig,"[<i><u>$3</u></i>]");
		}
		if (blog_wapImg) str=str.replace(/\[img\](.+?)\[\/img\]/ig,"<img src=\"$1\" alt=\"Image\"/>");
	
		return toUnicodeJS(str).replace(/&#x20;/g," ");
	}
</script>
