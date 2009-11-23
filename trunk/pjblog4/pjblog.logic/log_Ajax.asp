<!--#include file = "../include.asp" -->
<!--#include file = "../pjblog.model/cls_fso.asp" -->
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
			Case "GetArticledownloadInfo" Call GetArticledownloadInfo
			Case "delfile" Call DelFile
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
	
	Private Sub GetArticledownloadInfo
		Dim id, Rs, Str, varStr, i, s
		id = Trim(Asp.CheckStr(Request.QueryString("id")))
		If Not Asp.IsInteger(id) Then Exit Sub
		Set Rs = Conn.Execute("Select * From blog_Attment Where Blog_ID=" & id)
		If Rs.Bof Or Rs.Eof Then
			Response.Write("{Suc : false, Info : '抱歉, 找不到该日志附件信息, 或者该日志没有使用附件!'}")
		Else
			Str = ""
			Str = Str & "<div style=""border: 1px solid #7FCAE2; padding:10px 10px 10px 10px; display:block; background:#fff"">"
			i = 0
			Do While Not Rs.Eof
				i = i + 1
				If Asp.CheckExe(Rs("Att_type")) Then
					varStr = "../download.asp?id=" & Rs("Att_ID").value & " (" & Ucase(Rs("Att_type")) & " " & Rs("Att_Size").value & ")"
				Else
					s = Rs("Att_Path").value
					varStr = Mid(s, InStrRev(s, "/") + 1, Len(s)) & " (" & Rs("Att_Size").value & ")"
				End If
				Str = Str & "<div style=""line-height:20px;clear:both"" id=""upBox_" & Rs("Att_ID").value & """><div style=""margin-right:120px; float:left"">" & i & ". " & varStr & "</div><div style=""width:100px; float:right""><a href=""javascript:;"" onClick=""Upload.open(this, \'Message\', 0, \'../upload/month_" & Asp.DateToStr(Now(), "ym") & "\', 1)"">更新</a> <a href=""javascript:;"" onclick=""Upload.DelFile(" & Rs("Att_ID").value & ")"">删除</a></div></div>"
			Rs.MoveNext
			Loop
			Str = Str & "<div style=""clear:both; text-align:right; color:#bbb; font-size:10px; border-top:1px solid #ddd; line-height:22px;"">CopyRight @ 2009 PJBlog4 " & Sys.version & "&nbsp;&nbsp;&nbsp;&nbsp;<a href=""javascript:;"" onclick=""$(\'AttUpload\').removeChild($(\'ArticleAttmentBox\'))"" class=""close"">&nbsp;&nbsp;&nbsp;</a></div>"
			Str = Str & "</div>"
			Response.Write("{Suc : true, Info : '" & Str & "'}")
		End If
		Set Rs = Nothing
	End Sub
	
	Private Sub DelFile
		Dim id, Rs, OK, FilePath, fso
		id = Trim(Asp.CheckStr(Request.QueryString("id")))
		If Not Asp.IsInteger(id) Then Exit Sub
		Set Rs = Server.CreateObject("Adodb.RecordSet")
		Rs.open "Select * From blog_Attment Where Att_ID=" & id, Conn, 1, 3
		If Rs.Bof Or Rs.Eof Then
			Response.Write("{Suc : false, Info : '抱歉, 找不到该插件信息!'}")
			Exit Sub
		Else
			FilePath = Rs("Att_Path").value
			Set fso = New cls_fso
				OK = fso.DeleteFile(FilePath)
			Set fso = Nothing
			If OK Then Rs.Delete
		End If
		Set Rs = Nothing
		Response.Write("{Suc : true, Info : '删除附件成功<br /><font color=red>请将文本框中对应的UBB标签手动删除!</font>'}")
	End Sub

End Class
%>