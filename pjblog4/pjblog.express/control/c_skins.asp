<%
Public Sub c_skins
	Dim fso, cxml, i
	Dim FileItems, FolderItems
	Dim SplitName
	
	Dim PluginName, AuthorName, PluginVersion, PJBlogSuitVersion, ModDate
	
%>
	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
		<tr>
			<th class="CTitle"><%=control.categoryTitle%></th>
		</tr>
		<tr>
			<td class="CPanel">
			<div class="SubMenu2">
				<b>风格设置: </b> <a href="?Fmenu=Skins">设置外观</a> <b>模板: </b> <a href="?Fmenu=Skins&Smenu=CodeEdit">编辑静态页面模板</a>  <br/> 
				<b>插件相关: </b> <a href="?Fmenu=Skins&Smenu=module">设置模块</a> | <a href="?Fmenu=Skins&Smenu=Plugins">已装插件管理</a> | <a href="?Fmenu=Skins&Smenu=PluginsInstall">安装模块插件</a>
            </div>
<%
	If Request.QueryString("Smenu") = "PluginsInstall" Then
		'On Error Resume Next
		Set fso = New cls_fso
		Set cxml = New xml
			FolderItems = fso.FolderItem("../pjblog.plugin/")' 获取总的文件夹数和名
			SplitName = Split(FolderItems, "|")
%>
			<table cellpadding="3" cellspacing="0" width="100%">
            	<tr>
                	<td>&nbsp;</td>
                    <td>插件名</td>
                    <td>作者</td>
                    <td>当前版本号</td>
                    <td>适用博客版本</td>
                    <td>最后修改</td>
                    <td>&nbsp;</td>
                </tr>
<%
			For i = 1 To UBound(SplitName)
				If fso.FileExists("../pjblog.plugin/" & SplitName(i) & "/install.xml") Then
					cxml.FilePath = "../pjblog.plugin/" & SplitName(i) & "/install.xml"
					If cxml.open Then
					PluginName = cxml.GetNodeText(cxml.FindNode("//Plugin/Name"))
						If Err Then Err.Clear : PluginName = ""
					AuthorName = cxml.GetNodeText(cxml.FindNode("//Plugin/Author/Name"))
						If Err Then Err.Clear : AuthorName = ""
					PluginVersion = cxml.GetNodeText(cxml.FindNode("//Plugin/Version"))
						If Err Then Err.Clear : PluginVersion = ""
					PJBlogSuitVersion = cxml.GetNodeText(cxml.FindNode("//Plugin/PJBlogSuitVersion"))
						If Err Then Err.Clear : PJBlogSuitVersion = ""
					ModDate = cxml.GetNodeText(cxml.FindNode("//Plugin/ModDate"))
						If Err Then Err.Clear : ModDate = ""
%>
				<tr>
                	<td class="SecTd"><%=i%></td>
                    <td class="SecTd"><%=PluginName%></td>
                    <td class="SecTd"><%=AuthorName%></td>
                    <td class="SecTd"><%=PluginVersion%></td>
                    <td class="SecTd"><%=PJBlogSuitVersion%> (<%If CheckVersion(PJBlogSuitVersion) Then Response.Write("<font color=green>可用</font>") Else Response.Write("<font color=red>不可用</font>")%>)</td>
                    <td class="SecTd"><%=ModDate%></td>
                    <td class="SecTd"><a href="../pjblog.logic/control/log_plugin.asp?action=install&folder=<%=cee.encode(SplitName(i))%>">安装插件</a></td>
                </tr>
<%
					End If
				End If
			Next
%>
			</table>
<%
		Set cxml = Nothing
		Set fso = Nothing
	ElseIf Request.QueryString("Smenu") = "CodeEdit" Then
%>
<iframe src="log_styleEdit.asp?action=default" width="100%" height="800"></iframe>
<%
	Else
	
	End If
%>
            </td>
        </tr>
    </table>
<%
Control.getMsg
End Sub

Public Function CheckVersion(ByVal Source)
	Dim SplitSource, SourceLeft, SourceRight
	SplitSource = Split(Source, "-")
	SourceLeft = Trim(SplitSource(0))
	SourceRight = Trim(SplitSource(1))
	If SourceLeft = "*" And SourceRight = "*" Then
		CheckVersion = True
	ElseIf SourceLeft = "*" And SourceRight <> "*" Then
		If CheckSingleversion(SourceRight) Then CheckVersion = True Else CheckVersion = False
	ElseIf SourceRight = "*" And SourceLeft <> "*" Then
		If Not CheckSingleversion(SourceLeft) Then CheckVersion = True Else CheckVersion = False
	Else
		If CheckSingleversion(SourceRight) And (Not CheckSingleversion(SourceLeft)) Then CheckVersion = True Else CheckVersion = False
	End If
End Function

Public Function CheckSingleversion(ByVal Source)
	CheckSingleversion = True
	Dim SplitVersion, i, SplitSource
	SplitSource = Split(Source, ".")
	SplitVersion = Split(Sys.version, ".")
	For i = 0 To UBound(SplitVersion)
		If Int(SplitSource(i)) < Int(SplitVersion(i)) Then
			CheckSingleversion = False
			Exit For
		End If
	Next
End Function
%>