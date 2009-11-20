<%
Public Sub c_skins
	Dim fso, cxml, i, Rs, PluginInstalledArray
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
	' -----------------------------------------------------
	'	未安装插件
	' -----------------------------------------------------
	If Request.QueryString("Smenu") = "PluginsInstall" Then
		PluginInstalledArray = ""
		On Error Resume Next
		Set Rs = Conn.Execute("Select * From blog_Module")
		If Not (Rs.Bof And Rs.Eof) Then
			Do While Not Rs.Eof
				PluginInstalledArray = PluginInstalledArray & Rs("plugin_folder").value & ","
			Rs.MoveNext
			Loop
		End If
		If Len(PluginInstalledArray) > 0 Then PluginInstalledArray = Mid(PluginInstalledArray, 1, Len(PluginInstalledArray) - 1)
		Set Rs = Nothing
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
			If CheckExe(SplitName(i), PluginInstalledArray) Then
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
			End If
			Next
%>
			</table>
<%
		Set cxml = Nothing
		Set fso = Nothing
	' -----------------------------------------------------
	'	已安装插件
	' -----------------------------------------------------
	ElseIf Request.QueryString("Smenu") = "Plugins" Then
%>
	<table cellpadding="3" cellspacing="0" width="100%">
    	<tr>
        	<td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>插件标识</td>
            <td>版本号</td>
            <td>适用号</td>
            <td>最后修改时间</td>
            <td>作者</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
        </tr>
<%
		Set Rs = Conn.Execute("Select * From blog_Module")
		If Rs.Bof Or Rs.Eof Then
			Response.Write("<tr><td colspan=""9"" align=""center"">没有插件被安装</td></tr>")
		Else
			Do While Not Rs.Eof
%>
		<tr>
        	<td><%=Rs("plugin_ID").value%></td>
            <td><%=Rs("plugin_name").value%></td>
            <td><%=Rs("plugin_Mark").value%></td>
            <td><%=Rs("plugin_version").value%></td>
            <td><%=Rs("plugin_PJBlogSuitVersion").value%></td>
            <td><%=Rs("plugin_ModDate").value%></td>
            <td><%If Len(Rs("plugin_author_url").value) > 0 Then%><a href="<%=Rs("plugin_author_url").value%>" target="_blank"><%=Rs("plugin_author_name").value%></a><%Else%><%=Rs("plugin_author_name").value%><%End If%></td>
            <td>
				<%
					If Rs("plugin_stop").value Then
						If Rs("plugin_system").value Then
							Response.Write("<font color=""#cccccc"">停用中</font>")
						Else
							Response.Write("停用中")
						End If
					Else
						If Rs("plugin_system").value Then
							Response.Write("<font color=""#cccccc"">启用中</font>")
						Else
							Response.Write("启用中")
						End If
					End If
				%>
            </td>
            <td>
            	<%
					If Len(Rs("plugin_setPath").value) > 0 Then
						Response.Write("<a href=""?Fmenu=Skins&Smenu=BaseSetting&SetPath=" & Rs("plugin_setPath").value & "&mark=" & Rs("plugin_Mark").value & "&Folder=" & Rs("plugin_folder").value & "&name=" & Rs("plugin_name").value & """>基本设置</a>")
					Else
						Response.Write("<font color=""#cccccc"">基本设置</font>")
					End If
				%>
                &nbsp;&nbsp;
                <%
					If Len(Rs("plugin_backPath").value) > 0 Then
						Response.Write("<a href=""" & Rs("plugin_backPath").value & """>高级设置</a>")
					Else
						Response.Write("<font color=""#cccccc"">高级设置</font>")
					End If
				%>
            </td>
            <td><a href="../pjblog.logic/control/log_plugin.asp?action=UnInstall&folder=<%=cee.encode(Rs("plugin_folder").value) & "&mark=" & cee.encode(Rs("plugin_Mark").value)%>">卸载该插件</a>  插件升级</td>
        </tr>
<%
			Rs.MoveNext
			Loop
		End If
		Set Rs = Nothing
%>
	</table>
<%
	' -----------------------------------------------------
	'	插件基本设置
	' -----------------------------------------------------
	ElseIf Request.QueryString("Smenu") = "BaseSetting" Then
%>
	<div align=""center"">
    	<form action="../pjblog.logic/control/log_plugin.asp?action=BaseSetting" method="post" style="margin:0px">
        	<table border="0" cellpadding="3" cellspacing="1" class="TablePanel" style="margin:6px">
<%
		Dim SettingPath, SettingMark, SettingFolder, SettingName, SettingTemps, SettingTemp, TempElement, NameType
		Dim KeyName, KeyValue, KeyType, KeyDescription, SrttingTitle, KeySize, GetTrue, KeyRows, KeyCols, optionelements, optionelement
		Dim optionvalue, optionText
		SettingPath = Asp.CheckStr(Request.QueryString("SetPath"))
		SettingMark = Asp.CheckStr(Request.QueryString("mark"))
		SettingFolder = Asp.CheckStr(Request.QueryString("Folder"))
		SettingName	= Asp.CheckStr(Request.QueryString("name"))
		If Len(SettingPath) = 0 Or Len(SettingMark) = 0 Or Len(SettingFolder) = 0 Then
			Response.Write("<tr><td>缺少参数,无法运行.</td></tr>")
		Else
			Set cxml = New xml
			Set Rs = Conn.Execute("Select * From blog_modsetting Where xml_name='" & SettingMark & "'")
			If Rs.Bof Or Rs.Eof Then
				Response.Write("<tr><td>没有数据</td></tr>")
			Else
				cxml.FilePath = "../pjblog.plugin/" & SettingFolder & "/" & SettingPath
				If cxml.open Then
					Response.Write("<input type=""hidden"" name=""PluginsName"" value=""" & SettingMark & """/>")
					Response.Write("<tr><td colspan=""2"" align=""left"" style=""background:#e5e5e5;padding:6px""><div style=""font-weight:bold;font-size:14px;"">" & SettingName & " 的基本设置</div></td></tr>")
					Do While Not Rs.Eof
						KeyName = Rs("xml_key").value
						KeyValue = Rs("xml_value").value
						Set SettingTemps = cxml.FindNodes("Key")
						GetTrue = False
						For Each SettingTemp In SettingTemps
							NameType = cxml.GetAttribute(SettingTemp, "name")
							If NameType = KeyName Then
								Set TempElement = SettingTemp
								GetTrue = True
								Exit For
							End If
						Next
						If GetTrue Then
							KeyType = Lcase(cxml.GetAttribute(TempElement, "type"))
								If Err Then Err.Clear
							KeyDescription = cxml.GetAttribute(TempElement, "description")
								If Err Then Err.Clear
							KeySize = cxml.GetAttribute(TempElement, "size")
								If Err Then Err.Clear
							KeyRows = cxml.GetAttribute(TempElement, "rows")
								If Err Then Err.Clear
							KeyCols = cxml.GetAttribute(TempElement, "cols")
								If Err Then Err.Clear
							If KeyType = "select" Then
								Response.Write("<tr><td align=""right"" width=""200"" valign=""top"" style=""padding-top:6px"">" & KeyDescription & "</td><td width=""300""><select name=""" & KeyName & """>")
								' ------------------------------------------
								'	option
								' ------------------------------------------
								Set optionelements = TempElement.getElementsByTagName("option")
								For Each optionelement In optionelements
									optionvalue = cxml.GetAttribute(optionelement, "value")
									If Err Then Err.Clear : optionvalue = ""
									optionText = cxml.GetNodeText(optionelement)
									If Err Then Err.Clear : optionText = ""
									If KeyValue = optionvalue Then
										Response.Write("<option value=""" & optionvalue & """ selected>" & optionText & "</option>")
									Else
										Response.Write("<option value=""" & optionvalue & """>" & optionText & "</option>")
									End If
								Next
								' ------------------------------------------
								'	option
								' ------------------------------------------
								Response.Write("</select></td></tr>")
							ElseIf KeyType = "textarea" Then
								 Response.Write "<tr><td align=""right"" width=""200"" valign=""top"" style=""padding-top:6px"">" & KeyDescription & "</td><td width=""300""><textarea name=""" & KeyName & """ rows=""" & KeyRows & """ cols=""" & KeyCols & """>" & KeyValue & "</textarea></td></tr>"
							Else
								Response.Write("<tr><td align=""right"" width=""200"" valign=""top"" style=""padding-top:6px"">" & KeyDescription & "</td><td width=""300""><input name=""" & KeyName & """ type=""" & KeyType & """ size=""" & KeySize & """ value=""" & KeyValue & """/></td></tr>")
							End If
						End If
					Rs.MoveNext
					Loop
				Else
					Response.Write("<tr><td>打开配置文件失败</td></tr>")
				End If
			End If
			Set Rs = Nothing
		End If
%>
			<%="<tr><td colspan=""2"" align=""center""><input type=""submit"" name=""Submit"" value=""保存设置"" class=""button""/><input type=""button"" value=""放弃返回"" class=""button"" onclick=""history.go(-1)""/></td></tr>"%>
            </table>
        </form>
    </div>
<%
	' -----------------------------------------------------
	'	模板编辑器
	' -----------------------------------------------------
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

Public Function CheckExe(ByVal Str, ByVal Arrays)
	Dim SplitArray, i
	CheckExe = True
	SplitArray = Split(Arrays, ",")
	For i = 0 To UBound(SplitArray)
		If SplitArray(i) = Str Then
			CheckExe = False
			Exit For
		End If
	Next
End Function
%>