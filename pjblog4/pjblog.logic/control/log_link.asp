<!--#include file = "../../include.asp" -->
<!--#include file = "../../pjblog.data/control/cls_link.asp" -->
<%
Dim dolink
Set dolink = New do_Link
Set dolink = Nothing

Class do_Link

	Private OK
	
	Public Property Get Action
		Action = Request.QueryString("action")
	End property
	
	' ***********************************************
	'	类初始化
	' ***********************************************
	Private Sub Class_Initialize
		Select Case Action
			Case "edit" Call edit
			Case "session" Call SetSession
			Case "addNewClass" Call addNewClass
			Case "addNewLink" Call addNewLink
			Case "del" Call del
			Case "show" Call show(True)
			Case "unshow" Call show(False)
			Case "tomain" Call main(True)
			Case "unmain" Call main(False)
			Case "move" Call Moveto
			Case "showAll" Call showAll
		End Select
    End Sub 
     
	' ***********************************************
	'	类终结化
	' ***********************************************
    Private Sub Class_Terminate
		Sys.Close
    End Sub
	
	Private Sub edit
		Dim name, url, image, id, order
		name = Trim(Asp.CheckStr(Request.QueryString("name")))
		url = Trim(Asp.CheckStr(Request.QueryString("url")))
		image = Trim(Asp.CheckStr(Request.QueryString("image")))
		id = Trim(Asp.CheckStr(Request.QueryString("id")))
		order = Trim(Asp.CheckStr(Request.QueryString("order")))
		If Not Asp.IsInteger(order) Then order = 0
		If Not Asp.IsInteger(id) Then
			Response.Write("{Suc : false, Info : '非法参数'}")
			Exit Sub
		End If
		link.link_Name = name
		link.link_URL = url
		link.link_Image = image
		link.link_Order = Int(order)
		link.link_ID = Int(id)
		OK = link.edit
		If OK(0) Then
			Response.Write("{Suc : true, Info : '更新成功'}")
		Else
			Response.Write("{Suc : false, Info : '" & OK(1) & "'}")
		End If
	End Sub
	
	Private Sub SetSession
		Dim key, values
		key = Asp.CheckStr(Request.QueryString("key"))
		values = Asp.CheckStr(Request.QueryString("value"))
		Session(Sys.CookieName & key) = values
		Response.Write("{Suc : true}")
	End Sub
	
	Private Sub addNewClass
		Dim name
		name = Trim(Asp.CheckStr(Request.QueryString("name")))
		link.link_Name = name
		OK = link.addNewClass
		If OK(0) Then
			Session(Sys.CookieName & "_ShowMsg") = True
			Session(Sys.CookieName & "_MsgText") = "增加新分类成功!"
			Response.Write("{Suc : true, Info : '增加新分类成功!'}")
		Else
			Session(Sys.CookieName & "_ShowMsg") = True
			Session(Sys.CookieName & "_MsgText") = OK(1)
			Response.Write("{Suc : false, Info : '" & OK(1) & "'}")
		End If
	End Sub
	
	Private Sub addNewLink
		Dim order, name, url, image, cateID
		name = Trim(Asp.CheckStr(Request.QueryString("name")))
		url = Trim(Asp.CheckStr(Request.QueryString("url")))
		image = Trim(Asp.CheckStr(Request.QueryString("image")))
		order = Trim(Asp.CheckStr(Request.QueryString("order")))
		cateID = Int(Trim(Asp.CheckStr(Request.QueryString("cateID"))))
		If Not Asp.IsInteger(order) Then order = 0
		link.link_Name = name
		link.link_URL = url
		link.link_Image = image
		link.link_Order = Int(order)
		link.Link_ClassID = cateID
		OK = link.addNewLink
		If OK(0) Then
			Session(Sys.CookieName & "_ShowMsg") = True
			Session(Sys.CookieName & "_MsgText") = "增加新链接成功!"
			Response.Write("{Suc : true, Info : '添加成功!'}")
		Else
			Response.Write("{Suc : false, Info : '" & OK(1) & "'}")
		End If
	End Sub
	
	Private Sub del
		Dim id
		id = Trim(Asp.CheckStr(Request.QueryString("id")))
		If Not Asp.IsInteger(id) Then
			Response.Write("{Suc : false, Info : '非法参数'}")
			Exit Sub
		End If
		Ok = link.del(id)
		If Ok(0) Then
			Response.Write("{Suc : true, Info : '删除链接成功'}")
			Session(Sys.CookieName & "_ShowMsg") = True
			Session(Sys.CookieName & "_MsgText") = "删除链接成功!"
		Else
			Response.Write("{Suc : false, Info : '" & OK(1) & "'}")
		End If
	End Sub
	
	Private Sub show(ByVal Bool)
		Dim id
		id = Trim(Asp.CheckStr(Request.QueryString("id")))
		If Not Asp.IsInteger(id) Then
			Response.Write("{Suc : false, Info : '非法参数'}")
			Exit Sub
		End If
		Sys.doGet("Update blog_Links Set link_IsShow=" & Bool & " Where link_ID=" & Int(id))
		Response.Write("{Suc : true, Info : '操作成功!'}")
	End Sub
	
	Private Sub main(ByVal Bool)
		Dim id
		id = Trim(Asp.CheckStr(Request.QueryString("id")))
		If Not Asp.IsInteger(id) Then
			Response.Write("{Suc : false, Info : '非法参数'}")
			Exit Sub
		End If
		Sys.doGet("Update blog_Links Set link_IsMain=" & Bool & " Where link_ID=" & Int(id))
		Response.Write("{Suc : true, Info : '操作成功!'}")
	End Sub
	
	Private Sub Moveto
		Dim id, classid
		id = Trim(Asp.CheckStr(Request.QueryString("id")))
		If Not Asp.IsInteger(id) Then
			Response.Write("{Suc : false, Info : '非法参数'}")
			Exit Sub
		End If
		classid = Trim(Asp.CheckStr(Request.QueryString("classid")))
		If Not Asp.IsInteger(classid) Then
			Response.Write("{Suc : false, Info : '非法参数'}")
			Exit Sub
		End If
		Sys.doGet("Update blog_Links Set Link_ClassID=" & classid & " Where link_ID=" & Int(id))
		Session(Sys.CookieName & "_ShowMsg") = True
		Session(Sys.CookieName & "_MsgText") = "移动链接成功!"
		Response.Write("{Suc : true, Info : '操作成功!'}")
	End Sub
	
	Private Sub showAll
		Session(Sys.CookieName & "CateID") = ""
		Session(Sys.CookieName & "IsShow") = ""
		Session(Sys.CookieName & "IsMain") = ""
		Session(Sys.CookieName & "Order")  = ""
		Response.Write("{Suc : true, Info : '清理完成'}")
	End Sub
	
End Class
%>