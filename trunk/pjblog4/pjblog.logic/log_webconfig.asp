<!--#include file = "../include.asp" -->
<!--#include file = "../pjblog.model/cls_fso.asp" -->
<!--#include file = "../pjblog.model/cls_Stream.asp" -->
<!--#include file = "../pjblog.model/cls_tag.asp" -->
<!--#include file = "../pjblog.data/cls_webconfig.asp" -->
<!--#include file = "../pjblog.model/cls_template.asp" -->
<%
Dim Config
Set Config = New log_webConfig
Set Config = Nothing

Class log_webConfig

	Private OK
	
	Public property Get Action
		Action = Request.QueryString("action")
	End property

	' ***********************************************
	'	类初始化
	' ***********************************************
	Private Sub Class_Initialize
		Select Case Lcase(action)
			Case "default" Call default
			Case "article" Call Article
			Case "category" Call Category
			Case "getarticleid" Call GetArticleID
			Case "getcategoryid" Call GetCategoryID
		End Select
    End Sub 
     
	' ***********************************************
	'	类终结化
	' ***********************************************
    Private Sub Class_Terminate
		Sys.Close
    End Sub
	
	Private Sub default
		OK = web.default
		Response.Write("{Suc : " & Lcase(OK(0)) & ", Info : '" & OK(1) & "'}")
	End Sub
	
	Private Sub Article
		Dim id
		id = Asp.CheckUrl(Asp.CheckStr(Trim(Request.QueryString("id"))))
		If Not Asp.IsInteger(id) Then
			Response.Write("{Suc : false, Path : '该日志不存在', Title : ''}")
			Exit Sub
		End If
		Ok = web.Article(id)
		Response.Write("{Suc : " & Lcase(OK(0)) & ", Path : '" & Trim(OK(1)) & "', Title : '" & Trim(OK(2)) & "'}")
	End Sub
	
	Private Sub GetArticleID
		Dim Rs, Str, Counts, id
		id = ""
		Set Rs = Server.CreateObject("Adodb.RecordSet")
		Rs.open "Select log_ID From blog_Content Where log_IsShow=True And log_IsDraft=False", Conn, 1, 1
		If Rs.Bof Or Rs.Eof Then
			Str = "{Suc : false, total : 0, id : []}"
		Else
			Counts = Rs.RecordCount
			Do While Not Rs.Eof
				id = id & Rs(0).value & ","
			Rs.MoveNext
			Loop
			If Len(id) > 0 Then id = Mid(id, 1, (Len(id) - 1))
			Str = "{Suc : true, total : " & Counts & ", id : [" & id & "]}"
		End If
		Rs.Close
		Set Rs = Nothing
		Response.Write(Str)
	End Sub
	
	Private Sub Category
		Dim CateID, Folder
		CateID = Asp.CheckStr(Trim(Request.QueryString("id")))
		Folder = Asp.CheckStr(Trim(Request.QueryString("folder")))
		Call doCategory(CateID, Folder)
	End Sub
	
	Private Sub doCategory(ByVal CateID, ByVal Folder)
		OK = web.category(CateID, Folder)
		If OK(0) Then
			Response.Write("{Suc : true, Info : '" & OK(1) & "', folder : '" & OK(2) & "'}")
		Else
			Response.Write("{Suc : false, Info : '" & OK(1) & "', folder : '" & OK(2) & "'}")
		End If
	End Sub
	
	
	
	Private Sub GetCategoryID
		Dim Rs, Str, Counts, id
		id = ""
		Set Rs = Server.CreateObject("Adodb.RecordSet")
		Rs.open "Select cate_ID, cate_Folder From blog_Category Where cate_OutLink=False And cate_Secret=False", Conn, 1, 1
		If Rs.Bof Or Rs.Eof Then
			Str = "{Suc : false, total : 0, id : []}"
		Else
			Counts = Rs.RecordCount
			Do While Not Rs.Eof
				id = id & "[" & Rs(0).value & ",'" & Rs(1).value & "'],"
			Rs.MoveNext
			Loop
			If Len(id) > 0 Then id = Mid(id, 1, (Len(id) - 1))
			Str = "{Suc : true, total : " & Counts & ", id : [" & id & "]}"
		End If
		Rs.Close
		Set Rs = Nothing
		Response.Write(Str)
	End Sub

End Class
%>