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
			Case "editClass" Call editClass
			Case "getLinkCateArray" Call getLinkCateArray
		End Select
    End Sub 
     
	' ***********************************************
	'	类终结化
	' ***********************************************
    Private Sub Class_Terminate
		Sys.Close
    End Sub
	
	Private Sub edit
		Dim name, url, info, id, order
		name = Trim(Asp.CheckStr(Request.Form("link_name")))
		url = Trim(Asp.CheckStr(Request.Form("link_url")))
		info = Trim(Asp.CheckStr(Request.Form("link_info")))
		id = Trim(Asp.CheckStr(Request.QueryString("id")))
		order = Trim(Asp.CheckStr(Request.Form("link_Order")))
		If Not Asp.IsInteger(order) Then order = 0
		If Not Asp.IsInteger(id) Then
			Response.Write("{Suc : false, Info : '非法参数'}")
			Exit Sub
		End If
		If Len(name) = 0 Then
			Response.Write("{Suc : false, Info : '链接名不能为空'}")
			Exit Sub
		End If
		If Len(url) = 0 Then
			Response.Write("{Suc : false, Info : '链接地址不能为空'}")
			Exit Sub
		End If
		link.link_Name = name
		link.link_URL = url
		link.link_info = info
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
		Dim name, order
		name = Trim(Asp.CheckStr(Request.Form("link_Name")))
		order = Trim(Asp.CheckStr(Request.Form("link_Order")))
		If Len(name) > 0 And (Len(order) > 0) Then
		link.link_Name = name
		link.link_Order = order
		OK = link.addNewClass
		If Ok(0) Then Response.Write("{Suc : true}") Else Response.Write("{Suc : false, Info : '" & OK(1) & "'}")
		Else
			Response.Write("{Suc : false, Info : '内容不能为空'}")
		End If
	End Sub
	
	Private Sub addNewLink
		Dim order, name, url, cateID, linkinfo
		name = Trim(Asp.CheckStr(Request.Form("link_Name")))
		url = Trim(Asp.CheckStr(Request.Form("link_Url")))
		order = Trim(Asp.CheckStr(Request.Form("link_Order")))
		linkinfo = Trim(Asp.CheckStr(Request.Form("link_info")))
		cateID = Trim(Asp.CheckStr(Request.Form("Link_ClassID")))
		If Not Asp.IsInteger(order) Then order = 0
		If Not Asp.IsInteger(cateID) Then
			Response.Write("{Suc : false, Info : '请选择分类完成添加友情链接操作.'}")
		Else
			If Int(cateID) = 0 Then
				Response.Write("{Suc : false, Info : '选择的分类非法'}")
			Else
				link.link_Name = name
				link.link_URL = url
				link.link_info = linkinfo
				link.link_Order = Int(order)
				link.Link_ClassID = Int(cateID)
				OK = link.addNewLink
				If OK(0) Then
					Response.Write("{Suc : true}")
				Else
					Response.Write("{Suc : false, Info : '" & OK(1) & "'}")
				End If
			End If
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
	
	Private Sub editClass
		Dim link_order, link_name, link_ID
		link_ID = Trim(Asp.CheckStr(Request.QueryString("id")))
		If Not Asp.IsInteger(link_ID) Then
			Response.Write("{Suc : false, Info : '非法参数'}")
		Else
			link_name = Trim(Asp.CheckStr(Request.Form("link_Name")))
			link_order = Trim(Asp.CheckStr(Request.Form("link_Order")))
			If Len(link_name) > 0 And (Len(link_order) > 0) Then
				link.link_ID = link_ID
				link.link_Name = link_name
				link.link_Order = link_order
				OK = link.editClass
				If Ok(0) Then Response.Write("{Suc : true}") Else Response.Write("{Suc : false, Info : '" & OK(1) & "'}")
			Else
				Response.Write("{Suc : false, Info : '内容不能为空'}")
			End If
		End If
	End Sub
	
	Private Sub getLinkCateArray
		Dim Rs, Str
		Set Rs = Conn.Execute("Select * From blog_Links Where Link_Type=True Order By link_Order Asc")
		If Rs.Bof Or Rs.Eof Then
			Response.Write("{Suc : false}")
		Else
			Str = ""
			Do While Not Rs.Eof
				Str = Str & "{id:" & Rs("link_ID").value & ",name:'" & Rs("link_Name").value & "'},"
			Rs.MoveNext
			Loop
			Str = Mid(Str, 1, Len(Str) - 1)
			Str = "[" & Str & "]"
			Response.Write("{Suc : true, Arr : " & Str & "}")
		End If
		Set Rs = Nothing
	End Sub
	
End Class
%>