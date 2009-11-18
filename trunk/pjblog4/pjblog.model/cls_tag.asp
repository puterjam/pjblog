<%
Class Tag

    Private Sub Class_Initialize()
        If Not IsArray(Application(Sys.CookieName&"_blog_Tags")) Then Reload
    End Sub

    Private Sub Class_Terminate()

    End Sub

    Public Sub Reload
        Cache.TagsCache(2) '更新Tag缓存
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
            filterHTML = Str
			Exit Function
        Else
            Dim log_Tag, log_TagItem
            For Each log_TagItem IN Cache.TagsCache(1)
                log_Tag = Split(log_TagItem, "||")
                Str = Replace(Str, "{"&log_Tag(0)&"}", "<a href=""../default.asp?tag="&Server.URLEncode(log_Tag(1))&""">"&log_Tag(1)&"</a>")
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
            For Each log_TagItem IN Cache.TagsCache(1)
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
        For Each log_TagItem IN Cache.TagsCache(1)
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
        For Each log_TagItem IN Cache.TagsCache(1)
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
        For Each log_TagItem IN Cache.TagsCache(1)
            log_Tag = Split(log_TagItem, "||")
            If LCase(log_Tag(1)) = LCase(ClearHTML(tagName)) Then
                getTagID = log_Tag(0)
                Exit Function
            End If
        Next
    End Function
End Class
%>