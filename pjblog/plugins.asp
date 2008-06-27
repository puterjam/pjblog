<%'---- ASPCode For AboutMeForPJBlog ----%>
<%'---- ASPCode For NewLogForPJBlog ----%>
<%
Function NewArticle(ByVal action)
    Dim blog_Article
    If Not IsArray(Application(CookieName&"_blog_Article")) Or action = 2 Then
        Dim book_Articles, book_Article
        Set book_Articles = Conn.Execute("SELECT top 10 C.log_ID,C.log_Author,C.log_IsShow,C.log_PostTime,C.log_title,L.cate_ID,L.cate_Secret FROM blog_Content AS C,blog_Category AS L where L.cate_ID=C.log_CateID and L.cate_Secret=false and C.log_IsDraft=false order by log_PostTime Desc")
        SQLQueryNums = SQLQueryNums + 1
        TempVar = ""
        Do While Not book_Articles.EOF
            If book_Articles("cate_Secret") Then
                book_Article = book_Article&TempVar&book_Articles("log_ID")&"|,|"&book_Articles("log_Author")&"|,|"&book_Articles("log_PostTime")&"|,|"&"[隐藏分类日志]"
            ElseIf book_Articles("log_IsShow") Then
                book_Article = book_Article&TempVar&book_Articles("log_ID")&"|,|"&book_Articles("log_Author")&"|,|"&book_Articles("log_PostTime")&"|,|"&book_Articles("log_title")
            Else
                book_Article = book_Article&TempVar&book_Articles("log_ID")&"|,|"&book_Articles("log_Author")&"|,|"&book_Articles("log_PostTime")&"|,|"&"[隐藏日志]"
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

'处理最新日志内容
Dim Article_code
If Session(CookieName&"_LastDo") = "DelArticle" Or Session(CookieName&"_LastDo") = "AddArticle" Or Session(CookieName&"_LastDo") = "EditArticle" Then NewArticle(2)
Article_code = NewArticle(0)
side_html_default = Replace(side_html_default, "<$NewLog$>", Article_code)
side_html = Replace(side_html, "<$NewLog$>", Article_code)

%>
<%'---- ASPCode For GuestBookForPJBlog ----%>
<%'---- ASPCode For GuestBookForPJBlogSubItem1 ----%>

<%
Function NewMessage(ByVal action)
    Dim blog_Message
    If Not IsArray(Application(CookieName&"_blog_Message")) Or action = 2 Then
        Dim book_Messages, book_Message
        Set book_Messages = Conn.Execute("SELECT top 10 * FROM blog_book order by book_PostTime Desc")
        SQLQueryNums = SQLQueryNums + 1
        TempVar = ""
        Do While Not book_Messages.EOF
            If book_Messages("book_HiddenReply") Then
                book_Message = book_Message&TempVar&book_Messages("book_ID")&"|,|"&book_Messages("book_Messager")&"|,|"&book_Messages("book_PostTime")&"|,|"&"[隐藏留言]"
            Else
                book_Message = book_Message&TempVar&book_Messages("book_ID")&"|,|"&book_Messages("book_Messager")&"|,|"&book_Messages("book_PostTime")&"|,|"&book_Messages("book_Content")
            End If
            TempVar = "|$|"
            book_Messages.MoveNext
        Loop
        Set book_Messages = Nothing
        blog_Message = Split(book_Message, "|$|")
        Application.Lock
        Application(CookieName&"_blog_Message") = blog_Message
        Application.UnLock
    Else
        blog_Message = Application(CookieName&"_blog_Message")
    End If

    If action<>2 Then
        Dim Message_Items, Message_Item
        For Each Message_Items IN blog_Message
            Message_Item = Split(Message_Items, "|,|")
            NewMessage = NewMessage&"<a class=""sideA"" href=""LoadMod.asp?plugins=GuestBookForPJBlog#book_"&Message_Item(0)&""" title="""&Message_Item(1)&" 于 "&Message_Item(2)&" 发表留言"&Chr(10)&CCEncode(CutStr(Message_Item(3), 25))&""">"&CCEncode(CutStr(Message_Item(3), 25))&"</a>"
        Next
    End If
End Function

'处理最新留言内容
Dim Message_code
If Session(CookieName&"_LastDo") = "DelMessage" Or Session(CookieName&"_LastDo") = "AddMessage" Then NewMessage(2)
Message_code = NewMessage(0)
side_html_default = Replace(side_html_default, "<$NewMsg$>", Message_code)
side_html = Replace(side_html, "<$NewMsg$>", Message_code)

%>
