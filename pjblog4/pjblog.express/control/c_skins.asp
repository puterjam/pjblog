<%
Public Sub c_skins
	Dim fso, cxml, i, Rs, PluginInstalledArray
	Dim FileItems, FolderItems
	Dim SplitName
	
	Dim PluginName, AuthorName, PluginVersion, PJBlogSuitVersion, ModDate, PluginInfo
	
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
			<div style="margin:10px 20px 10px 20px">
            <table cellpadding="3" cellspacing="0"class="CeeTable" width="100%">
                <thead>
                    <tr>
                        <th>&nbsp;</th>
                        <th class="act">未安装插件</th>
                        <th>描述</th>
                    </tr>
                </thead>
                <tfoot>
                    <tr>
                        <th>&nbsp;</th>
                        <th>未安装插件</th>
                        <th>描述</th>
                    </tr>
                </tfoot>
                <tbody>
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
					PluginInfo = cxml.GetNodeText(cxml.FindNode("//Plugin/Info"))
						If Err Then Err.Clear : ModDate = ""
%>

				<tr class="active">
                    <td><input type="checkbox" value="" name="SelectID"></td>
                    <td><%=PluginName%></td>
                    <td><%=PluginInfo%></td>
                </tr>
                <tr class="active second">
                    <td>&nbsp;</td>
                    <td class="act"><a href="../pjblog.logic/control/log_plugin.asp?action=install&folder=<%=cee.encode(SplitName(i))%>">安装</a> | 检测</td>
                    <td>
                        <div class="left">
                        版本 <%=PluginVersion%> | 
                        作者 <%=AuthorName%> | 
                        发布时间 <%=ModDate%> | 
                        使用版本 <%=PJBlogSuitVersion%>
                        </div>
                        <div class="cright cs">
                            <a href="javascript:;">删除插件</a>
                        </div>
                        <div class="clear"></div>
                    </td>
                </tr>
<%
					End If
				End If
			End If
			Next
%>
			</tbody>
			</table>
            </div>
<%
		Set cxml = Nothing
		Set fso = Nothing
	' -----------------------------------------------------
	'	已安装插件
	' -----------------------------------------------------
	ElseIf Request.QueryString("Smenu") = "Plugins" Then
%>
	<div style="margin:10px 20px 10px 20px">
	<table cellpadding="3" cellspacing="0"class="CeeTable" width="100%">
        <thead>
        	<tr>
        		<th>&nbsp;</th>
            	<th class="act">已安装插件</th>
                <th>描述</th>
            </tr>
        </thead>
        <tfoot>
        	<tr>
        		<th>&nbsp;</th>
            	<th>已安装插件</th>
                <th>描述</th>
            </tr>
        </tfoot>
        <tbody>
<%
		Set Rs = Conn.Execute("Select * From blog_Module")
		If Rs.Bof Or Rs.Eof Then
			Response.Write("<tr><td colspan=""9"" align=""center"">没有插件被安装</td></tr>")
		Else
			Do While Not Rs.Eof
%>
		<tr class="active">
        	<td><input type="checkbox" value="" name="SelectID"></td>
            <td><%=Rs("plugin_name").value%></td>
            <td><%=Rs("plugin_info").value%></td>
        </tr>
        <tr class="active second">
        	<td>&nbsp;</td>
            <td class="act"><%If Rs("plugin_stop").value Then%>启用<%Else%>停用<%End If%> | <a href="../pjblog.logic/control/log_plugin.asp?action=UnInstall&folder=<%=cee.encode(Rs("plugin_folder").value) & "&mark=" & cee.encode(Rs("plugin_Mark").value)%>" onclick="return confirm('危险操作\n卸载插件将删除所有该插件的安装信息!!')">卸载</a> | 升级</td>
            <td>
            	<div class="left">版本 <%=Rs("plugin_version").value%> | 
                作者 <%If Len(Rs("plugin_author_url").value) > 0 Then%><a href="<%=Rs("plugin_author_url").value%>" target="_blank"><%=Rs("plugin_author_name").value%></a><%Else%><%=Rs("plugin_author_name").value%><%End If%> | 
                适用版本 <%=Rs("plugin_PJBlogSuitVersion").value%> | 
                发布时间 <%=Rs("plugin_ModDate").value%></div>
                <div class="cright cs">
                	<%
						If Len(Rs("plugin_setPath").value) > 0 Then
							Response.Write("<a href=""?Fmenu=Skins&Smenu=BaseSetting&SetPath=" & Rs("plugin_setPath").value & "&mark=" & Rs("plugin_Mark").value & "&Folder=" & Rs("plugin_folder").value & "&name=" & Rs("plugin_name").value & """>基本设置</a>")
						Else
							Response.Write("<font color=""#cccccc"">基本设置</font>")
						End If
					%>
                     | 
                    <%
						If Len(Rs("plugin_backPath").value) > 0 Then
							Response.Write("<a href=""" & Rs("plugin_backPath").value & """>高级设置</a>")
						Else
							Response.Write("<font color=""#cccccc"">高级设置</font>")
						End If
					%>
                </div>
                <div class="clear"></div>
            </td>
        </tr>
		
<%
			Rs.MoveNext
			Loop
		End If
		Set Rs = Nothing
%>
		</tbody>
	</table>
    </div>
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
	' -----------------------------------------------------
	'	主题增加插件
	' -----------------------------------------------------
	ElseIf Request.QueryString("Smenu") = "AddPlus" Then
		Dim f1 : f1 = Asp.CheckStr(Request.QueryString("f1")) ' 路径做标识
		Dim Plus, Plus_Str
'		Set Plus = Conn.Execute("Select * From blog_Module")
'		If Plus.Bof Or Plus.Eof Then
'			Plus_Str = ""
'		Else
'			Plus_Str = ""
'			Do While Not Plus.Eof
'				Plus_Str = Plus_Str & "<input type=""checkbox"" name=""plusMark"" value="""" />"
'			Plus.MoveNext
'			Loop
'		End If
'		Set Plus = Nothing
%>
		<div>
        <h5>※ 已支持的插件</h5>
        <form action="../pjblog.logic/control/log_template.asp?action=ImportPlus" method="post" id="import">
		<table cellpadding="3" cellspacing="0" width="100%">
        	<tr>
            	<td>插件模板名</td>
                <td>插件模板标识</td>
                <td>加载插件路径</td>
                <td>插件所在文件夹</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
            </tr>
<%
		Dim Ass, AsCount, ms
		AsCount = Int(Sys.doGet("Select Count(*) From blog_tempPlugin Where tp_templateName='" & f1 & "'")(0))
		AsCount = AsCount - 1
		ReDim Ass(AsCount)
		Set Plus = Conn.Execute("Select * From blog_tempPlugin Where tp_templateName='" & f1 & "'")
		If Plus.Bof Or Plus.Eof Then
			Response.Write("<tr><td align=""center"" colspan=""7"" style=""border-top:1px dotted #ccc"">该主题暂无插件支持</td></tr>")
		Else
			ms = 0
			Do While Not Plus.Eof
				Ass(ms) = Plus("tp_pluginSingleMark")
				Response.Write("<tr>")
				Response.Write("<input type=""hidden"" value=""" & Plus("tp_PluginMark") & "." & Plus("tp_pluginSingleMark") & """ name=""key"">")
				Response.Write("<td>" & Plus("tp_pluginSingleName") & "</td>")
				Response.Write("<td>" & Plus("tp_pluginSingleMark") & "</td>")
				Response.Write("<td>" & Plus("tp_pluginSinglePath") & "<input type=""hidden"" name=""pluginSinglePath"" value=""" & Plus("tp_pluginSinglePath") & """></td>")
				Response.Write("<td>" & Plus("tp_pluginPath") & "<input type=""hidden"" name=""pluginPath"" value=""" & Plus("tp_pluginPath") & """></td>")
				If Len(Trim(Plus("tp_pluginSingleTempValue").value)) > 0 Then
					Response.Write("<td><a href=""javascript:;"" onclick=""center.PlusCode(" & Plus("tp_ID") & ")"">编辑插件模板代码</a></td>")
				Else
					Response.Write("<td>&nbsp;</td>")
				End If
				
				If Len(Plus("tp_plugintag").value) > 0 Then
					Response.Write("<td><a href=""javascript:;"" onclick=""center.GetPlusTag('" & Plus("tp_plugintag") & "')"">获取页面标签</a></td>")
				Else
					Response.Write("<td><span style=""color:#ccc"">获取页面标签</span></td>")
				End If
				Response.Write("<td><a href=""../pjblog.logic/control/log_template.asp?action=delPlus&id=" & Plus("tp_ID") & """>删除</a></td>")
				Response.Write("</tr>")
				ms = ms + 1
			Plus.MoveNext
			Loop
		End If
		Set Plus = Nothing
		Response.Write("<tr><td align=""right"" colspan=""7"" style=""border-top:1px dotted #ccc; padding-right:30px;""><a href=""javascript:;"" onclick=""CheckForm.Theme.ImportPlus('import')"">将以上插件导入到本主题中</a></td><td>&nbsp;</td></tr>")
%>
		</table>
        </form>
        
        <h5>※ 为此主题增加更多插件支持</h5>
        <table cellpadding="3" cellspacing="0" width="100%">
        	<tr>
            	<td>插件模板名</td>
                <td>插件模板标识</td>
                <td>加载插件路径</td>
                <td>插件所在文件夹</td>
                <td>&nbsp;</td>
            </tr>
<%
	Dim addi
		Set fso = New cls_fso
		Set cxml = New xml
		Dim tp_pluginSingleMark, tp_pluginSinglePath, tp_pluginSingleTempPath, tp_pluginPath, tp_pluginSingleName, tp_PluginMark, tp_templateName, Plugin_Mark
		Dim Config_Template, obj_mode, temp, subtemp
			'FileItems, FolderItems
			FolderItems = fso.FolderItem("../pjblog.plugin")
			SplitName = Split(FolderItems, "|")
			If Int(SplitName(0)) > 0 Then
				For addi = 1 To UBound(SplitName)
					If fso.FileExists("../pjblog.plugin/" & SplitName(addi) & "/install.xml") Then
						cxml.FilePath = "../pjblog.plugin/" & SplitName(addi) & "/install.xml"
						If cxml.open Then
							Plugin_Mark = cxml.GetNodeText(cxml.FindNode("//Plugin/Mark"))
							If Err Then Err.Clear : Plugin_Mark = ""
							If Int(Sys.doGet("Select Count(*) From blog_Module Where plugin_Mark='" & Plugin_Mark & "'")(0)) > 0 Then
							Set Config_Template = cxml.FindNode("//Plugin/Config/Template")
							If Lcase(cxml.GetAttribute(Config_Template, "do")) = "true" Then
								Set obj_mode = Config_Template.selectSingleNode("mode")
									If Lcase(cxml.GetAttribute(obj_mode, "do")) = "true" Then
										Set temp = obj_mode.getElementsByTagName("item")
											On Error Resume Next
											For Each subtemp In temp
												tp_pluginSingleMark = cxml.GetNodeText(subtemp.selectSingleNode("mark"))
													If Err Then Err.Clear : tp_pluginSingleMark = "&nbsp;"
												tp_pluginSinglePath = cxml.GetNodeText(subtemp.selectSingleNode("include"))
													If Err Then Err.Clear : tp_pluginSinglePath = "&nbsp;"
												tp_pluginSingleName = cxml.GetNodeText(subtemp.selectSingleNode("name"))
													If Err Then Err.Clear : tp_pluginSingleName = "&nbsp;"
												tp_pluginSingleTempPath = cxml.GetNodeText(subtemp.selectSingleNode("templatePath"))
													If Err Then Err.Clear : tp_pluginSingleTempPath = "&nbsp;"
												If Not CheckAddPlus(Ass, tp_pluginSingleMark) Then
													Response.Write("<tr>")
													Response.Write("<td style=""border-top:1px dotted #ccc"">" & tp_pluginSingleName & "</td>")
													Response.Write("<td style=""border-top:1px dotted #ccc"">" & tp_pluginSingleMark & "</td>")
													Response.Write("<td style=""border-top:1px dotted #ccc"">" & tp_pluginSinglePath & "</td>")
													Response.Write("<td style=""border-top:1px dotted #ccc"">" & tp_pluginSingleTempPath & "</td>")
													Response.Write("<td style=""border-top:1px dotted #ccc""><a href=""javascript:;"" onclick=""CheckForm.Theme.AddPlus('" & f1 & "', '" & SplitName(addi) & "', '" & tp_pluginSingleMark & "', this)"">添加</a></td>")
													Response.Write("</tr>")
												End If
											Next
									End If
							End If
							End If
						End If
					End If
				Next
			End If
		Set cxml = Nothing
		Set fso = Nothing
%>
		</table>
        </div>
<%
	' -----------------------------------------------------
	'	主题选择
	' -----------------------------------------------------
	Else
		Set cxml = New xml
		Set fso = New cls_fso
			'FileItems, FolderItems
			FolderItems = fso.FolderItem("../pjblog.template")
%>
	<table cellpadding="3" cellspacing="0" width="100%">
<%
		Dim TempFolderArray, tempi, tempj, m, cleft, cright, ztrue
		Dim SkinName, SkinDesigner, pubDate, DesignerURL, DesignerMail, version, smallpic, bigpic
		TempFolderArray = Split(FolderItems, "|")
		If Int(TempFolderArray(0)) = 0 Then ' 是否有文件夹
			Response.Write("<tr><td align=""center"">找不到模板, 请上传模板!</td></tr>")
		Else
			tempi = 0
			m = Int(TempFolderArray(0))
			On Error Resume Next
			For tempj = 1 To Ubound(TempFolderArray)
				If fso.FileExists("../pjblog.template/" & TempFolderArray(tempj) & "/skin.xml") Then ' 是否文件存在
					cxml.FilePath = "../pjblog.template/" & TempFolderArray(tempj) & "/skin.xml"
					If cxml.open Then ' 是否正常加载
						tempi = tempi + 1
						If tempi Mod 3 = 1 Then
							cleft = "<tr>"
							cright = ""
						ElseIf tempi Mod 3 = 0 Then
							cleft = ""
							cright = "</tr>"
						Else
							cleft = ""
							cright = ""
						End If
						' ---------------------------------------------
						' 	获取信息
						' ---------------------------------------------
						SkinName = cxml.GetNodeText(cxml.FindNode("//SkinSet/SkinName"))
							If Err Then Err.Clear : SkinName = ""
						SkinDesigner = cxml.GetNodeText(cxml.FindNode("//SkinSet/SkinDesigner"))
							If Err Then Err.Clear : SkinDesigner = ""
						pubDate = cxml.GetNodeText(cxml.FindNode("//SkinSet/pubDate"))
							If Err Then Err.Clear : pubDate = ""
						DesignerURL = cxml.GetNodeText(cxml.FindNode("//SkinSet/DesignerURL"))
							If Err Then Err.Clear : DesignerURL = ""
						DesignerMail = cxml.GetNodeText(cxml.FindNode("//SkinSet/DesignerMail"))
							If Err Then Err.Clear : DesignerMail = ""
						version = cxml.GetNodeText(cxml.FindNode("//SkinSet/version"))
							If Err Then Err.Clear : version = ""
						smallpic = cxml.GetNodeText(cxml.FindNode("//SkinSet/smallpic"))
							If Err Then Err.Clear : smallpic = ""
						bigpic = cxml.GetNodeText(cxml.FindNode("//SkinSet/bigpic"))
							If Err Then Err.Clear : bigpic = ""
							
					ztrue = (blog_DefaultSkin = TempFolderArray(tempj))
%>
    	<%=cleft%><td align="center">
        
        <table width="50%" cellpadding="3" cellspacing="0" style=" padding:3px; border:1px solid #ABA9C5;<%If ztrue Then Response.Write("background:#FFFFCC; border-color:#500000")%>">
        	<tr>
            	<td><img src="../pjblog.template/<%=TempFolderArray(tempj)%>/<%=smallpic%>" width="277"; height="170" style="border:1px dotted #ccc" onerror="this.src='../images/control/skin.jpg'"></td>
            </tr>
            <tr>
            	<td align="left">
                <div style="padding:0px 6px 0px 6px">
                	<h5 style="color:#004000"><%=SkinName%></h5>
                    <ul style="list-style: square">
                    	<li>作者 : <%=SkinDesigner%></li>
                        <li>发布时间 : <%=pubDate%></li>
                        <li>作者网址 : <%=DesignerURL%></li>
                        <li>作者邮箱 : <%=DesignerMail%></li>
                        <li>风格版本 : <%=version%></li>
                    </ul>
                    <p style="margin-bottom:8px;">
                    	<%If ztrue Then%>
                    		<a href="javascript:;" style="background:url(../images/notify.gif) no-repeat 0px 0px; padding-left:17px; line-height:20px;">当前主题</a> 
                        <%Else%>
                        	<a href="javascript:;" style="background:url(../images/notify.gif) no-repeat 0px 0px; padding-left:17px; line-height:20px;" onclick="CheckForm.Theme.Choose('<%=TempFolderArray(tempj)%>', this)">应用主题</a> 
                        <%End If%>
                        <a href="?Fmenu=Skins&Smenu=AddPlus&f1=<%=TempFolderArray(tempj)%>" style="background:url(../images/icon_apply.gif) no-repeat 0px 0px; padding-left:17px; line-height:20px;">为该主题添加插件</a> 
						<%If Not ztrue Then%>
                        	<a href="javascript:;" style="background:url(../images/hidden.gif) no-repeat 0px 0px; padding-left:20px; line-height:20px;">删除模板</a>
						<%Else%>
                            <a href="javascript:;" style="background:url(../images/icon_edit.gif) no-repeat 0px 0px; padding-left:17px; line-height:20px;" onclick="CheckForm.Theme.Choose('<%=TempFolderArray(tempj)%>', this)">选择样式</a>
						<%End If%>
                    </p>
                </div>
                </td>
            </tr>
        </table>
        
        </td><%=cright%>
<%
					End If
				End If
			Next
		End If
		Set fso = Nothing
		Set cxml = Nothing
%>
    </table>
<%
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

Public Function CheckAddPlus(Arrays, Str)
	CheckAddPlus = False
	Dim i
	For Each i In Arrays
		If i = Str Then
			CheckAddPlus = True
			Exit For
		End If
	Next
End Function
%>