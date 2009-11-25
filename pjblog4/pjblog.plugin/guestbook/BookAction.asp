<!--#include file = "../../include.asp" -->
<!--#include file = "../../pjblog.model/cls_template.asp" -->
<%
Dim Mud, page
Set Mud = New template
Mud.PluginPath = "../../../pjblog.template/" & Trim(blog_DefaultSkin) & "/"
Mud.Path = "template"
Mud.FileName = "guestbook.html"
Mud.TemplateContent = ""
Mud.open
Call General
Mud.Sets("KeyWords") = blog_KeyWords
Mud.Sets("Description") = blog_Description
Mud.PageUrl = "index_{$page}.html"
page = Trim(Asp.CheckStr(Request.QueryString("page")))
If Not Asp.IsInteger(page) Then page = 1
Mud.CurrentPage = Asp.CheckPage(page)
Mud.Buffer
Mud.Flush
' ***********************************************
'	全局变量
' ***********************************************
Public Sub General
	Mud.Sets("SiteName") = SiteName
	Mud.Sets("SiteURL") = SiteURL
	Mud.Sets("blogabout") = blogabout
	Mud.Sets("Skin") = blog_DefaultSkin
	Mud.Sets("blogTitle") = blog_Title
	Mud.Sets("master") = blog_master
	Mud.Sets("email") = blog_email
	Mud.Sets("cookie") = Sys.CookieName
	Mud.Sets("version") = Sys.version
	Mud.Sets("UpdateTime") = Sys.UpdateTime
	Mud.Sets("blogPerPage") = blogPerPage
	Mud.Sets("blogcommpage") = blogcommpage
End Sub
Sys.Close
%>