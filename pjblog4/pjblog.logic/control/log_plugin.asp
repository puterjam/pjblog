<!--#include file = "../../include.asp" -->
<!--#include file = "../../pjblog.model/cls_xml.asp" -->
<!--#include file = "../../pjblog.model/cls_Stream.asp" -->
<!--#include file = "../../pjblog.data/control/cls_category.asp" -->
<!--#include file = "../../pjblog.data/control/cls_plugin.asp" -->
<!--#include file = "../../pjblog.model/cls_fso.asp" -->
<%
Dim doPlugin
Set doPlugin = New do_Plugin
Set doPlugin = Nothing

Class do_Plugin

	Private cxml, local, InstallFolder, this, UnInstallMark, sxml
	
	Private Mark, Name, Info, Version, PJBlogSuitVersion, pubDate, ModDate, Author_Name, Author_Url, Author_Email, Author_pubUrl, BaseenterPole, BackenterPole, objSetting
	
	Private Config_Template, obj_BaseenterPole, obj_BackenterPole
	Private obj_mode, temp, temp_mark,temp_include, subtemp, WriteTemp, WriteInclude
	Private Config_InstallNav, Nav_Title, Nav_Intro, Nav_Url, Nav_Local, NavTemp
	Private Config_SQL, Sql_Install, Sql_Install_items, Sql_Install_item, textSql
	Private plugin_ID, plugin_Mark, plugin_version, plugin_PJBlogSuitVersion, plugin_pubDate, plugin_ModDate, plugin_author_name, plugin_author_url, plugin_author_email, plugin_info, plugin_navid, plugin_setPath, plugin_backPath, plugin_folder, plugin_name, plugin_pubUrl, plugin_stop, plugin_system
	Private SplitArray, j, Navid, Ok

	Public Property Get Action
		Action = Request.QueryString("action")
	End Property

	' ***********************************************
	'	类初始化
	' ***********************************************
	Private Sub Class_Initialize
		local = "../../pjblog.plugin/"
		InstallFolder = cee.decode(Trim(Request.QueryString("folder")))
		UnInstallMark = cee.decode(Trim(Request.QueryString("mark")))
		Select Case Action
			Case "install" Call install
			Case "UnInstall" Call UnInstall
			Case "BaseSetting" Call BaseSetting
		End Select
    End Sub 
     
	' ***********************************************
	'	类终结化
	' ***********************************************
    Private Sub Class_Terminate
		Sys.Close
    End Sub
	
	Private Sub install
		On Error Resume Next
		Set cxml = New xml
			cxml.FilePath = local & InstallFolder & "/install.xml"
			If cxml.open Then
				Mark = cxml.GetNodeText(cxml.FindNode("//Plugin/Mark"))
					If Err Then Err.Clear : Mark = ""
				If Sys.doGet("Select Count(*) From blog_Module Where plugin_Mark='" & Mark & "'")(0) > 0 Then
					Session(Sys.CookieName & "_ShowMsg") = True
					Session(Sys.CookieName & "_MsgText") = "插件已安装"
					RedirectUrl(RedoUrl)
				End If
				Name = cxml.GetNodeText(cxml.FindNode("//Plugin/Name"))
					If Err Then Err.Clear : Name = ""
				Info = cxml.GetNodeText(cxml.FindNode("//Plugin/Info"))
					If Err Then Err.Clear : Info = ""
				Version = cxml.GetNodeText(cxml.FindNode("//Plugin/Version"))
					If Err Then Err.Clear : Version = ""
				PJBlogSuitVersion = cxml.GetNodeText(cxml.FindNode("//Plugin/PJBlogSuitVersion"))
					If Err Then Err.Clear : PJBlogSuitVersion = ""
				pubDate = cxml.GetNodeText(cxml.FindNode("//Plugin/pubDate"))
					If Err Then Err.Clear : pubDate = ""
				ModDate = cxml.GetNodeText(cxml.FindNode("//Plugin/ModDate"))
					If Err Then Err.Clear : ModDate = ""
				Author_Name = cxml.GetNodeText(cxml.FindNode("//Plugin/Author/Name"))
					If Err Then Err.Clear : Author_Name = ""
				Author_Url = cxml.GetNodeText(cxml.FindNode("//Plugin/Author/Url"))
					If Err Then Err.Clear : Author_Url = ""
				Author_Email = cxml.GetNodeText(cxml.FindNode("//Plugin/Author/Email"))
					If Err Then Err.Clear : Author_Email = ""
				Author_pubUrl = cxml.GetNodeText(cxml.FindNode("//Plugin/Author/pubUrl"))
					If Err Then Err.Clear : Author_pubUrl = ""
				' ----------------------------------------------------------------
				'	判断安装SQL数据
				' ----------------------------------------------------------------	
				Set Config_SQL = cxml.FindNode("//Plugin/Config/InstallSQL")
				If Lcase(cxml.GetAttribute(Config_SQL, "do")) = "true" Then
					Set Sql_Install = Config_SQL.selectSingleNode("install")
					Set Sql_Install_items = Sql_Install.getElementsByTagName("sql")
					For Each Sql_Install_item In Sql_Install_items
						textSql = cxml.GetNodeText(Sql_Install_item)
						If Err Then Err.Clear : textSql = ""
						If Len(textSql) > 0 Then Conn.Execute(textSql)
					Next
				End If
				' ----------------------------------------------------------------
				'	判断模板数据
				' ----------------------------------------------------------------	
				Set Config_Template = cxml.FindNode("//Plugin/Config/Template")
				If Lcase(cxml.GetAttribute(Config_Template, "do")) = "true" Then
					' 	BaseenterPole
					Set obj_BaseenterPole = Config_Template.selectSingleNode("BaseenterPole")
					If Lcase(cxml.GetAttribute(obj_BaseenterPole, "do")) = "true" Then
						BaseenterPole = cxml.GetNodeText(obj_BaseenterPole)
						If Err Then Err.Clear : BaseenterPole = ""
						If Len(BaseenterPole) > 0 Then
							Set sxml = New xml
								sxml.FilePath = local & InstallFolder & "/" & BaseenterPole
								If sxml.open Then
									Dim sxmlkeyName, sxmlKeyvalue, BaseElements, BaseElement
									Set BaseElements = sxml.FindNodes("Key")
									For Each BaseElement In BaseElements
										sxmlkeyName = sxml.GetAttribute(BaseElement, "name")
											If Err Then Err.Clear : sxmlkeyName = ""
										sxmlKeyvalue = sxml.GetAttribute(BaseElement, "value")
											If Err Then Err.Clear : sxmlKeyvalue = ""
										Plugin.SetInsert sxmlkeyName, sxmlKeyvalue, Mark
									Next
								End If
							Set sxml = Nothing
						End If
					Else
						BaseenterPole = ""
					End If
					'	BackenterPole
					Set obj_BackenterPole = Config_Template.selectSingleNode("BackenterPole")
					If Lcase(cxml.GetAttribute(obj_BackenterPole, "do")) = "true" Then
						BackenterPole = cxml.GetNodeText(obj_BackenterPole)
						If Err Then Err.Clear : BackenterPole = ""
					Else
						BackenterPole = ""
					End If
					' ----------------------------------------------------------------
					'	判断写入文件数据
					' ----------------------------------------------------------------	
'					Set obj_mode = Config_Template.selectSingleNode("mode")
'					If Lcase(cxml.GetAttribute(obj_mode, "do")) = "true" Then
'						Set temp = obj_mode.getElementsByTagName("item")
'						For Each subtemp In temp
'							temp_mark = cxml.GetNodeText(subtemp.selectSingleNode("mark"))
'								If Err Then Err.Clear : temp_mark = ""
'							temp_include = cxml.GetNodeText(subtemp.selectSingleNode("include"))
'								If Err Then Err.Clear : temp_include = ""
'							If Len(temp_mark) > 0 And Len(temp_include) > 0 Then
'								WriteInclude = "<!--" & chr(35) & "include file=""../pjblog.plugin/" & InstallFolder & "/" & temp_include & """ -->"
'								WriteTemp = Plugin.WritePluginAsp(temp_mark, WriteInclude)
'							End If
'						Next
'					End If
				End If
				' ----------------------------------------------------------------
				'	判断导航数据
				' ----------------------------------------------------------------	
				Navid = ""
				Set Config_InstallNav = cxml.FindNode("//Plugin/Config/InstallNav")
				If Lcase(cxml.GetAttribute(Config_InstallNav, "do")) = "true" Then
					Set temp = Config_InstallNav.getElementsByTagName("item")
					For Each subtemp In temp
						Nav_Title = cxml.GetNodeText(subtemp.selectSingleNode("title"))
							If Err Then Err.Clear : Nav_Title = ""
						Nav_Intro = cxml.GetNodeText(subtemp.selectSingleNode("intro"))
							If Err Then Err.Clear : Nav_Intro = ""
						Nav_Url = cxml.GetNodeText(subtemp.selectSingleNode("url"))
							If Err Then Err.Clear : Nav_Url = ""
						Nav_Local = cxml.GetNodeText(subtemp.selectSingleNode("local"))
							If Err Then Err.Clear : Nav_Local = ""
						If Len(Nav_Title) > 0 And Len(Nav_Url) > 0 And Len(Nav_Local) > 0 Then
							Category.cate_Order = 99
							Category.cate_icon = "../images/icons/9.gif"
							Category.cate_Name = Nav_Title
							Category.cate_Intro = Nav_Intro
							Category.cate_Folder = ""
							Category.cate_URL = Nav_Url
							Category.cate_local = Int(Nav_Local)
							Category.cate_Secret = False
							Category.cate_count = 0
							Category.cate_OutLink = True
							Category.cate_Lock = False
							NavTemp = Category.Add
							If NavTemp(0) Then Navid = Navid & NavTemp(1) & ","
						End If
					Next
					If Len(Navid) > 0 Then Navid = Mid(Navid, 1, Len(Navid) - 1)
				Else
					Navid = ""
				End If
				
				'插入数据库
				Plugin.plugin_Mark = Mark
				Plugin.plugin_version = Version
				Plugin.plugin_PJBlogSuitVersion = PJBlogSuitVersion
				Plugin.plugin_pubDate = pubDate
				Plugin.plugin_ModDate = ModDate
				Plugin.plugin_author_name = Author_Name
				Plugin.plugin_author_url = Author_Url
				Plugin.plugin_author_email = Author_Email
				Plugin.plugin_info = Info
				Plugin.plugin_navid = Navid
				Plugin.plugin_setPath = BaseenterPole
				Plugin.plugin_backPath = BackenterPole
				Plugin.plugin_folder = InstallFolder
				Plugin.plugin_name = Name
				Plugin.plugin_pubUrl = Author_pubUrl
				Plugin.plugin_stop = False
				Plugin.plugin_system = False
				Ok = Plugin.install
			Else
				Ok = Array(False, "打开插件配置文件失败")
			End If
		Set cxml = Nothing
		
		Session(Sys.CookieName & "_ShowMsg") = True
		If (OK(0)) Then
			Session(Sys.CookieName & "_MsgText") = "插件安装成功, 请返回已安装插件进行设置!"
		Else
			Session(Sys.CookieName & "_MsgText") = OK(1)
		End If
		RedirectUrl(RedoUrl)
	End Sub
	
	Private Sub UnInstall
		If Len(UnInstallMark) = 0 Then
			Session(Sys.CookieName & "_ShowMsg") = True
			Session(Sys.CookieName & "_MsgText") = "参数无效"
			RedirectUrl(RedoUrl)
		End If
		'On Error Resume Next
		'Conn.BeginTrans
		Set cxml = New xml
			cxml.FilePath = local & InstallFolder & "/install.xml"
			If cxml.open Then
				' ----------------------------------------------------------------
				'	判断安装SQL数据
				' ----------------------------------------------------------------	
				Set Config_SQL = cxml.FindNode("//Plugin/Config/InstallSQL")
				If Lcase(cxml.GetAttribute(Config_SQL, "do")) = "true" Then
					Set Sql_Install = Config_SQL.selectSingleNode("Uninstall")
					Set Sql_Install_items = Sql_Install.getElementsByTagName("sql")
					For Each Sql_Install_item In Sql_Install_items
						textSql = cxml.GetNodeText(Sql_Install_item)
						If Err Then Err.Clear : textSql = ""
						If Len(textSql) > 0 Then Conn.Execute(textSql)
					Next
				End If
				' ----------------------------------------------------------------
				'	判断写入文件数据
				' ----------------------------------------------------------------
				Set Config_Template = cxml.FindNode("//Plugin/Config/Template")
				If Lcase(cxml.GetAttribute(Config_Template, "do")) = "true" Then	
					Set obj_mode = Config_Template.selectSingleNode("mode")
					If Lcase(cxml.GetAttribute(obj_mode, "do")) = "true" Then
						Set temp = obj_mode.getElementsByTagName("item")
						For Each subtemp In temp
							temp_mark = cxml.GetNodeText(subtemp.selectSingleNode("mark"))
								If Err Then Err.Clear : temp_mark = ""
							If Len(temp_mark) > 0 Then
								WriteTemp = Plugin.RemovePluginAsp(UnInstallMark & "." & temp_mark)
							End If
						Next
					End If
					' ----------------------------------------------------------------
					'	清理基本设置
					' ----------------------------------------------------------------
					Set objSetting = Config_Template.selectSingleNode("BaseenterPole")
					If Lcase(cxml.GetAttribute(objSetting, "do")) = "true" Then
						If len(UnInstallMark) > 0 Then
							Sys.doGet("Delete From blog_modsetting Where xml_name='" & UnInstallMark & "'")
						End If
					End If
				End If

				' ----------------------------------------------------------------
				'	从数据库中删除
				' ----------------------------------------------------------------
				Set this = Server.CreateObject("Adodb.RecordSet")
				this.open "Select * From blog_Module Where plugin_Mark='" & UnInstallMark & "'", Conn, 1, 3
				If this.Bof Or this.Eof Then
					Ok = Array(False, "卸载失败")
				Else
					Navid = this("plugin_navid").value
					this.Delete
					Ok = Array(True, "删除数据库数据成功")
				End If
				this.Close
				Set this = Nothing
				If OK(0) Then
					If Len(Navid) > 0 Then
						SplitArray = Split(Navid, ",")
						For j = 0 To UBound(SplitArray)
							If Asp.IsInteger(SplitArray(j)) Then Sys.doGet("Delete From blog_Category Where cate_ID=" & SplitArray(j))
						Next
					End If
				End If
			Else
				Ok = Array(False, "打开插件配置文件失败")
			End If
			Sys.doRecDel(Array(Array("blog_tempPlugin", "tp_PluginMark", "'" & UnInstallMark & "'")))
			'If Err.Number > 0 Then
				'Conn.RollBackTrans
				'Ok = Array(False, Err.Description)
			'Else
				'Conn.CommitTrans
				'Ok = Array(True, "卸载插件成功")
			'End If
		Set cxml = Nothing
		Session(Sys.CookieName & "_ShowMsg") = True
		If (OK(0)) Then
			Session(Sys.CookieName & "_MsgText") = "卸载插件成功!"
		Else
			Session(Sys.CookieName & "_MsgText") = OK(1)
		End If
		RedirectUrl(RedoUrl)
	End Sub
	
	Private Sub BaseSetting
		Dim ResMark
		ResMark = Request.Form("PluginsName")
		Set this = Server.CreateObject("Adodb.RecordSet")
		this.open "Select * From blog_modsetting Where xml_name='" & ResMark & "'", Conn, 1, 3
			If this.Bof Or this.Eof Then
				Session(Sys.CookieName & "_ShowMsg") = True
				Session(Sys.CookieName & "_MsgText") = "您的操作失去了更新的意义."
				RedirectUrl(RedoUrl)
			Else
				Do While Not this.Eof
					this("xml_value") = Asp.CheckStr(Request.Form(this("xml_key").value))
					this.update
				this.MoveNext
				Loop
				If Len(Trim(Asp.CheckStr(Request.Form("PluginSettingCache")))) > 0 Then
					If Asp.IsInteger(Trim(Asp.CheckStr(Request.Form("PluginSettingCache")))) Then
						If CBool(Int(Trim(Asp.CheckStr(Request.Form("PluginSettingCache"))))) Then
							Application.Lock()
							Application(Sys.CookieName & "_plugin_Setting_" & ResMark) = True
							Application(Sys.CookieName & "_mod_" & ResMark) = ""
							Application.UnLock()
						End If
					End If
				End If
				Session(Sys.CookieName & "_ShowMsg") = True
				Session(Sys.CookieName & "_MsgText") = "更新数据成功."
				RedirectUrl(RedoUrl)
			End If
		this.Close
		Set this = Nothing
	End Sub

End Class
%>