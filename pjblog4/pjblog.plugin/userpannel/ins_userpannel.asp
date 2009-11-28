<%
' *****************************************************
' 	用户控制面板
' *****************************************************
Function userPanel
    userPanel = ""
    If memName <> Empty Then userPanel = userPanel&" <b>"&memName&"</b>，欢迎你!<br/>你的权限: "&stat_title&"<br/><br/>"
    If stat_Admin = True Then userPanel = userPanel & "<a href=""../pjblog.express/control.asp"" target=""_blank"" class=""sideA"" accesskey=""3"">系统管理</a>"
    If stat_AddAll = True Or stat_Add = True Then userPanel = userPanel & "<a href=""../pjblog.express/blogpost.asp"" class=""sideA"" accesskey=""N"">发表新日志</a>"
    If (stat_AddAll = True Or stat_Add = True) And (stat_EditAll Or stat_Edit) Then
        If IsEmpty(session(Sys.CookieName&"_draft_"&memName)) Then
            If stat_EditAll Then
                session(Sys.CookieName&"_draft_"&memName) = conn.Execute("select count(log_ID) from blog_Content where log_IsDraft=true")(0)
            ElseIf stat_Edit Then
                session(Sys.CookieName&"_draft_"&memName) = conn.Execute("select count(log_ID) from blog_Content where log_IsDraft=true and log_Author='"&memName&"'")(0)
            Else
	            session(Sys.CookieName&"_draft_"&memName) = 0
            End If
        End If
        
        If session(Sys.CookieName&"_draft_"&memName) > 0 Then
            userPanel = userPanel & "<a href=""default.asp?display=draft"" class=""sideA"" accesskey=""D""><strong>编辑草稿 ["&session(Sys.CookieName&"_draft_"&memName)&"]</strong></a>"
        Else
            userPanel = userPanel & "<a href=""default.asp?display=draft"" class=""sideA"" accesskey=""D"">编辑草稿</a>"
        End If
    End If
    If memName<>Empty Then
        userPanel = userPanel&"<a href=""member.asp?action=edit"" class=""sideA"" accesskey=""M"">修改个人资料</a><a href=""../pjblog.express/login.asp?action=logout"" class=""sideA"" accesskey=""Q"">退出系统</a>"
    Else
        userPanel = userPanel&"<a href=""../pjblog.express/login.asp"" class=""sideA"" accesskey=""L"">登录</a><a href=""register.asp"" class=""sideA"" accesskey=""U"">用户注册</a>"
    End If
End Function
plus.open("sys.userpannel")
Dim te
te = Trim(Asp.UnCheckStr(plus.getSingleTemplate("SideUserPannel")))
te = Replace(te, "<#user#>", userPanel)
%>
	PluginTempValue = ['SideUserPannel', '<%=outputSideBar(te)%>'];
	PluginOutPutString.push(PluginTempValue);