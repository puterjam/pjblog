<!--#include file = "../include.asp" -->
<!--#include file = "../pjblog.model/cls_fso.asp" -->
<!--#include file = "../pjblog.model/cls_Stream.asp" -->
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
			cASE "getarticleid" Call GetArticleID
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
			Str = "{Suc : false, total : 0, id : ''}"
		Else
			Counts = Rs.RecordCount
			Do While Not Rs.Eof
				id = id & Rs(0).value & ","
			Rs.MoveNext
			Loop
			If Len(id) > 0 Then id = Mid(id, 1, (Len(id) - 1))
			Str = "{Suc : true, total : " & Counts & ", id : '" & id & "'}"
		End If
		Rs.Close
		Set Rs = Nothing
		Response.Write(Str)
	End Sub

End Class
%>