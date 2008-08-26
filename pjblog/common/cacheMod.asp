<%
'***************PJblog2 动态文章保存*******************
' PJblog2 Copyright 2005
' Update:2005-10-18
'**************************************************

Sub PostArticle(LogID)
    If Not blog_postFile Then Exit Sub
    Dim SaveArticle, LoadTemplate, Temp, TempStr, log_View, preLogC, nextLogC
    '读取日志模块
    LoadTemplate = LoadFromFile("Template/Article.asp")
    If LoadTemplate(0) = 0 Then '读取成功后写入信息
        '读取分类信息
        Temp = LoadTemplate(1)
        '读取日志内容
        SQL = "SELECT TOP 1 * FROM blog_Content WHERE log_ID=" & LogID
        SQLQueryNums = SQLQueryNums + 1
        Set log_View = conn.Execute(SQL)
        Dim blog_Cate, blog_CateArray, comDesc
        Dim getCate, getTags
        Set getCate = New Category
        Set getTags = New tag
        getCate.load(Int(log_View("log_CateID"))) '获取分类信息

        Temp = Replace(Temp, "<$Cate_icon$>", getCate.cate_icon)
        Temp = Replace(Temp, "<$Cate_Title$>", getCate.cate_Name)
        Temp = Replace(Temp, "<$log_CateID$>", log_View("log_CateID"))
        Temp = Replace(Temp, "<$LogID$>", LogID)
        Temp = Replace(Temp, "<$log_Title$>", HtmlEncode(log_View("log_Title")))
        Temp = Replace(Temp, "<$log_Author$>", log_View("log_Author"))
        Temp = Replace(Temp, "<$log_PostTime$>", DateToStr(log_View("log_PostTime"), "Y-m-d"))
        Temp = Replace(Temp, "<$log_weather$>", log_View("log_weather"))
        Temp = Replace(Temp, "<$log_level$>", log_View("log_level"))
        Temp = Replace(Temp, "<$log_Author$>", log_View("log_Author"))
        Temp = Replace(Temp, "<$log_IsShow$>", log_View("log_IsShow"))
        Temp = Replace(Temp, "<$log_tag$>", getTags.filterHTML(log_View("log_tag")))
        If log_View("log_comorder") Then comDesc = "Desc" Else comDesc = "Asc" End If
        Temp = Replace(Temp, "<$comDesc$>", comDesc)
        Temp = Replace(Temp, "<$log_DisComment$>", log_View("log_DisComment"))
        TempStr = "&lt;%if stat_EditAll or (stat_Edit and memName="""&log_View("log_Author")&""") then %&gt;　<a href=""blogedit.asp?id="&log_View("log_ID")&""" title=""编辑该日志"" accesskey=""E""><img src=""images/icon_edit.gif"" alt="""" border=""0"" style=""margin-bottom:-2px""/></a>&lt;%end if%&gt;"&Chr(10)
        TempStr = TempStr & "&lt;%if stat_DelAll or (stat_Del and memName="""&log_View("log_Author")&""")  then %&gt;　<a href=""blogedit.asp?action=del&amp;id="&log_View("log_ID")&""" onclick=""if (!window.confirm('是否要删除该日志')) return false"" title=""删除该日志"" accesskey=""K""><img src=""images/icon_del.gif"" alt="""" border=""0"" style=""margin-bottom:-2px""/></a>&lt;%end if%&gt;"
        Temp = Replace(Temp, "<$EditAndDel$>", HTMLDecode(TempStr))
        If log_View("log_edittype") = 1 Then
            Temp = Replace(Temp, "<$ArticleContent$>", UnCheckStr(UBBCode(HtmlEncode(log_View("log_Content")), Mid(log_View("log_ubbFlags"), 1, 1), Mid(log_View("log_ubbFlags"), 2, 1), Mid(log_View("log_ubbFlags"), 3, 1), Mid(log_View("log_ubbFlags"), 4, 1), Mid(log_View("log_ubbFlags"), 5, 1))))
        Else
            Temp = Replace(Temp, "<$ArticleContent$>", UnCheckStr(log_View("log_Content")))
        End If
        If Len(log_View("log_Modify"))>0 Then
            Temp = Replace(Temp, "<$log_Modify$>", log_View("log_Modify")&"<br/>")
        Else
            Temp = Replace(Temp, "<$log_Modify$>", "")
        End If

        Temp = Replace(Temp, "<$log_FromUrl$>", log_View("log_FromUrl"))
        Temp = Replace(Temp, "<$log_From$>", log_View("log_From"))
        Temp = Replace(Temp, "<$trackback$>", SiteURL&"trackback.asp?tbID="&LogID)
        Temp = Replace(Temp, "<$log_CommNums$>", log_View("log_CommNums"))
        Temp = Replace(Temp, "<$log_QuoteNums$>", log_View("log_QuoteNums"))
        Temp = Replace(Temp, "<$log_IsDraft$>", log_View("log_IsDraft"))

        Set preLogC = Conn.Execute("SELECT TOP 1 log_Title,log_ID FROM blog_Content WHERE log_PostTime<#"&DateToStr(log_View("log_PostTime"), "Y-m-d H:I:S")&"# and (log_IsShow=true or log_Readpw<>'') and log_IsDraft=false ORDER BY log_PostTime DESC")
        Set nextLogC = Conn.Execute("SELECT TOP 1 log_Title,log_ID FROM blog_Content WHERE log_PostTime>#"&DateToStr(log_View("log_PostTime"), "Y-m-d H:I:S")&"# and (log_IsShow=true or log_Readpw<>'') and log_IsDraft=false ORDER BY log_PostTime ASC")

        Dim BTemp
        BTemp = ""
        If Not preLogC.EOF Then
            BTemp = BTemp & " | <a href=""?id="&preLogC("log_ID")&""" title=""上一篇日志: "&preLogC("log_Title")&""" accesskey="",""><img border=""0"" src=""images/Cprevious.gif"" alt=""""/>上一篇</a>"
        Else
            BTemp = BTemp & " | <img border=""0"" src=""images/Cprevious1.gif"" alt=""这是最新一篇日志""/>上一篇"
        End If
        If Not nextLogC.EOF Then
            BTemp = BTemp & " | <a href=""?id="&nextLogC("log_ID")&""" title=""下一篇日志: "&nextLogC("log_Title")&""" accesskey="".""><img border=""0"" src=""images/Cnext.gif"" alt=""""/>下一篇</a>"
        Else
            BTemp = BTemp & " | <img border=""0"" src=""images/Cnext1.gif"" alt=""这是最后一篇日志""/>下一篇"
        End If
        Temp = Replace(Temp, "<$log_Navigation$>", BTemp)

        SaveArticle = SaveToFile(Temp, "post/" & LogID & ".asp")
        Set getCate = Nothing
    End If
End Sub
%>
