<!--#include file = "../include.asp" -->
<%
'分类管理
Dim Plugin_Category, Plugin_Category_String, Plugin_loop_i
Plugin_Category = Cache.CategoryCache(1)
If UBound(Plugin_Category, 1) = 0 Then
	Plugin_Category_String = ""
Else
	Plugin_Category_String = ""
	For Plugin_loop_i = 0 To UBound(Plugin_Category, 2)
		If Int(Plugin_Category(9, Plugin_loop_i)) = 0 Or Int(Plugin_Category(9, Plugin_loop_i)) = 2 Then
			Plugin_Category_String = Plugin_Category_String & "<img src=""" & Plugin_Category(6, Plugin_loop_i) & """ border=""0"" style=""margin:3px 4px -4px 0px;"" alt=""" & Plugin_Category(3, Plugin_loop_i) & """/>"
			If Plugin_Category(4, Plugin_loop_i) Then
				Plugin_Category_String = Plugin_Category_String & "<a class=""CategoryA"" href=""" & Plugin_Category(5, Plugin_loop_i) & """ title=""" & Plugin_Category(3, Plugin_loop_i) & """>" & Plugin_Category(1, Plugin_loop_i) & "</a><br/>"
			Else
				Plugin_Category_String = Plugin_Category_String & "<a class=""CategoryA"" href=""cate_" & Plugin_Category(11, Plugin_loop_i) & "_1.html"" title=""" & Plugin_Category(3, Plugin_loop_i) & """>" & Plugin_Category(1, Plugin_loop_i) & "</a><br/>"
			End If
		End If
	Next
End If
%>
	PluginTempValue = ['plugin_Category', '<%=outputSideBar(Plugin_Category_String)%>'];
	PluginOutPutString.push(PluginTempValue);
<%
' 用户控制面板
Function userPanel
    userPanel = ""
    If memName <> Empty Then userPanel = userPanel&" <b>"&memName&"</b>，欢迎你!<br/>你的权限: "&stat_title&"<br/><br/>"
    If stat_Admin = True Then userPanel = userPanel & "<a href=""control.asp"" target=""_blank"" class=""sideA"" accesskey=""3"">系统管理</a>"
    If stat_AddAll = True Or stat_Add = True Then userPanel = userPanel & "<a href=""blogpost.asp"" class=""sideA"" accesskey=""N"">发表新日志</a>"
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
        userPanel = userPanel&"<a href=""member.asp?action=edit"" class=""sideA"" accesskey=""M"">修改个人资料</a><a href=""login.asp?action=logout"" class=""sideA"" accesskey=""Q"">退出系统</a>"
    Else
        userPanel = userPanel&"<a href=""login.asp"" class=""sideA"" accesskey=""L"">登录</a><a href=""register.asp"" class=""sideA"" accesskey=""U"">用户注册</a>"
        If blog_PasswordProtection Then
        	userPanel = userPanel&"<a href=""javascript:void(0)"" class=""sideA"" accesskey=""P"" onclick=""PasswordProtection();"">忘记密码</a>"
        End If
    End If
End Function
%>
	PluginTempValue = ['plugin_UserPannel', '<%=outputSideBar(userPanel)%>'];
	PluginOutPutString.push(PluginTempValue);
<%
Sys.Close
%>