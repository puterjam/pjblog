<!--#include file = "../include.asp" -->
<%
'*************** PJblog4 基本设置 *******************
' PJblog4 Copyright 2009
' Update: 2009-11-03
' Author: evio
' Owner : PuterJam
' Mode  : Class
' MoName: Ajax Form
'***************************************************
Dim Ajax
Set Ajax = New Sys_Ajax
Set Ajax = Nothing

Class Sys_Ajax

	Public Property Get Action
		Action = Request.QueryString("action")
	End Property

	' ***********************************************
	'	整站初始化类初始化
	' ***********************************************
	Private Sub Class_Initialize
		Select Case Action
			Case "CheckCode" Call CheckCode
			Case "downloadInfo" Call DownLoadInfo
		End Select
    End Sub 
     
	' ***********************************************
	'	整站初始化类终结化
	' ***********************************************
    Private Sub Class_Terminate
		Sys.Close
    End Sub
	
	Private Sub CheckCode
		Response.Write(Trim(Session("GetCode")))
	End Sub
	
	Private Sub DownLoadInfo
		Dim id, Rs, Str
		id = Trim(Asp.CheckStr(Request.QueryString("id")))
		If Not Asp.IsInteger(id) Then Exit Sub
		Set Rs = Conn.Execute("Select * From blog_Attment Where Att_ID=" & id)
		If Rs.Bof Or Rs.Eof Then
			Response.Write("{Suc : false, Info : '抱歉, 找不到该插件信息!'}")
		Else
			Str = ""
			Str = Str & "<div style=""font-size:14px; font-weight:bolder;""><a href=""../download.asp?id=" & id & """>点击下载该附件</a></div>"
			Str = Str & "<div style=""margin-top:8px;"">"
			Str = Str & "<div style=""line-height:20px;""><strong>上传日期:</strong> " & Asp.DateToStr(Rs("Att_Date"), "Y-m-d H:I:S") & "</div>"
			Str = Str & "<div style=""line-height:20px;""><strong>附件大小:</strong> " & Rs("Att_Size") & " Bytes</div>"
			Str = Str & "<div style=""line-height:20px;""><strong>附件类型:</strong> " & Ucase(Rs("Att_type")) & "</div>"
			Str = Str & "</div>"
			Response.Write("{Suc : true, Info : '" & Str & "'}")
		End If
		Set Rs = Nothing
	End Sub

End Class
%>