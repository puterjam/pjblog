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

	Private cxml, local, InstallFolder
	
	Private Mark, Name, Info, Version, PJBlogSuitVersion, pubDate, ModDate, Author_Name, Author_Url, Author_Email, Author_pubUrl, BaseenterPole, BackenterPole
	
	Private Config_Template, obj_BaseenterPole, obj_BackenterPole
	Private obj_mode, temp, temp_mark,temp_include, subtemp, WriteTemp, WriteInclude
	Private Config_InstallNav, Nav_Title, Nav_Intro, Nav_Url, Nav_Local, NavTemp
	
	Private plugin_ID, plugin_Mark, plugin_version, plugin_PJBlogSuitVersion, plugin_pubDate, plugin_ModDate, plugin_author_name, plugin_author_url, plugin_author_email, plugin_info, plugin_navid, plugin_setPath, plugin_backPath, plugin_folder, plugin_name, plugin_pubUrl, plugin_stop, plugin_system

	Public Property Get Action
		Action = Request.QueryString("action")
	End Property

	' ***********************************************
	'	类初始化
	' ***********************************************
	Private Sub Class_Initialize
		local = "../../pjblog.plugin/"
		InstallFolder = cee.decode(Trim(Request.QueryString("folder")))
		Select Case Action
			Case "install" Call install
		End Select
    End Sub 
     
	' ***********************************************
	'	类终结化
	' ***********************************************
    Private Sub Class_Terminate
		Sys.Close
    End Sub
	
	Private Sub install
		'On Error Resume Next
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
				'	判断模板数据
				' ----------------------------------------------------------------	
				Set Config_Template = cxml.FindNode("//Plugin/Config/Template")
				If Lcase(cxml.GetAttribute(Config_Template, "do")) = "true" Then
					' 	BaseenterPole
					Set obj_BaseenterPole = Config_Template.selectSingleNode("BaseenterPole")
					If Lcase(cxml.GetAttribute(obj_BaseenterPole, "do")) = "true" Then
						BaseenterPole = cxml.GetNodeText(obj_BaseenterPole)
						If Err Then Err.Clear : BaseenterPole = ""
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
					'	mode
					
					Set obj_mode = Config_Template.selectSingleNode("mode")
					If Lcase(cxml.GetAttribute(obj_mode, "do")) = "true" Then
						Set temp = obj_mode.getElementsByTagName("item")
						For Each subtemp In temp
							temp_mark = cxml.GetNodeText(subtemp.selectSingleNode("mark"))
								If Err Then Err.Clear : temp_mark = ""
							temp_include = cxml.GetNodeText(subtemp.selectSingleNode("include"))
								If Err Then Err.Clear : temp_include = ""
							If Len(temp_mark) > 0 And Len(temp_include) > 0 Then
								WriteInclude = "<!--" & chr(35) & "include file=""../pjblog.plugin/" & InstallFolder & "/" & temp_include & """ -->"
								WriteTemp = Plugin.WritePluginAsp(temp_mark, WriteInclude)
							End If
						Next
					End If
				End If
				' ----------------------------------------------------------------
				'	判断导航数据
				' ----------------------------------------------------------------	
				Dim Navid : Navid = ""
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
				Dim Ok
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

End Class
%>