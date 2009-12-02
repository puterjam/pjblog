<!--#include file = "../include.asp" -->
<!--#include file = "../pjblog.model/cls_upload.asp" -->
<!--#include file = "../pjblog.model/cls_fso.asp" -->
<%
Dim doUpload
Set doUpload = New Sys_Upload
Set doUpload = Nothing

Class Sys_Upload
	
	Private Upload
	
	Private Att_Path, Att_Size, Att_type, Att_Ext, Att_width, Att_height, Att_upName
	Private Arrays, OK
	
	Public Property Get Action
		Action = Request.QueryString("action")
	End Property

	Private Sub Class_Initialize()
		Set Upload = New UpLoadClass
			Upload.TotalSize = 10240000				' @ 上传总大小限制字节数
			Upload.Charset = "UTF-8"				' @ 接受字符集
			Upload.MaxSize = 1024000				' @ 每个上传文件的最大字节数
			Upload.FileType = "jpg/png/gif/rar"		' @ 允许上传的文件类型
			Upload.SavePath = "../upload/"			' @ 文件存放的路径，可以是相对路径
			'Upload.AutoSave = 2					' @ 设置 Open 方法处理文件的方式，对其他方法无效
													'		默认值：0
													'		可选值：
													'			0：取无重复的服务器时间字符串为文件名自动保存文件
													'			1：取源文件名自动保存文件
													'			2：不自动保存文件，Open之后请用Save/GetData方法保存文件
			Select Case Action
				Case "fast" Call FastUpLoad
				Case "update" Call Updates
				Case "singledel" Call SingleDel
			End Select
	End Sub
	
	Private Sub FastUpLoad
		Dim DataTotal, i, FormName, Str, intCount, ResponseString
		Upload.AutoSave = 0
		ResponseString = ""
		DataTotal = Upload.open
		For i = 1 To UBound(Upload.FileItem)
			FormName 	= 	Upload.FileItem(i)
			Att_Path 	= 	Upload.SavePath & Upload.Form(FormName)
			Att_Size 	= 	Upload.Form(FormName & "_Size")
			Att_type 	= 	Upload.Form(FormName & "_Type")
			Att_Ext 	= 	Upload.Form(FormName & "_Ext")
			Att_width 	= 	Upload.Form(FormName & "_Width")
			Att_height 	= 	Upload.Form(FormName & "_Height")
			Att_upName 	= 	Upload.Form(FormName & "_Name")
			If Upload.Form(FormName & "_Err") = 0 Then
				Str = InToBase
				ResponseString  = ResponseString & "{id : " & Str(1) & ", name : '" & Att_upName & "', path : '" & Att_Path & "', size : " & Att_Size & ", type : '" & Att_type & "', ext : '" & Att_Ext & "'},"
			End If
		Next
		ResponseString = Mid(ResponseString, 1, Len(ResponseString) - 1)
		ResponseString = "[" & ResponseString & "]"
		Response.Write("{Suc : true, Info : " & ResponseString & "}")
	End Sub
	
	Private Sub Class_Terminate()
	  Set Upload = Nothing
	  Sys.Close
	End Sub

	Private Function InToBase
		Arrays = Array(Array("Att_Path", Att_Path), Array("Att_Date", Now()), Array("Att_Size", Att_Size), Array("Att_type", Att_type), Array("Att_Ext", Att_Ext), Array("Att_width", Att_width), Array("Att_height", Att_height), Array("Att_upName", Att_upName))
		OK = Sys.doRecord("blog_Attment", Arrays, "insert", "Att_ID", "")
		InToBase = OK
	End Function
	
	Private Function upToBase(ByVal id)
		Arrays = Array(Array("Att_Path", Att_Path), Array("Att_Date", Now()), Array("Att_Size", Att_Size), Array("Att_type", Att_type), Array("Att_Ext", Att_Ext), Array("Att_width", Att_width), Array("Att_height", Att_height), Array("Att_upName", Att_upName))
		OK = Sys.doRecord("blog_Attment", Arrays, "update", "Att_ID", id)
		upToBase = OK
	End Function
	
	Private Function DelFromBase(ByVal id)
		OK = Sys.doRecDel(Array(Array("blog_Attment", "Att_ID", id)))
		DelFromBase = OK
	End Function
	
	Private Sub Updates
		Dim id, FormName, Str, ResponseString, i, oldPath, fso
		id = Asp.CheckStr(Request.QueryString("id"))
		If Not Asp.IsInteger(id) Then
			Response.Write("{Suc : false, Info : '非法参数'}")
		Else
			Upload.AutoSave = 0
			Upload.open
			ResponseString = ""
			For i = 1 To UBound(Upload.FileItem)
				FormName 	= 	Upload.FileItem(i)
				Att_Path 	= 	Upload.SavePath & Upload.Form(FormName)
				Att_Size 	= 	Upload.Form(FormName & "_Size")
				Att_type 	= 	Upload.Form(FormName & "_Type")
				Att_Ext 	= 	Upload.Form(FormName & "_Ext")
				Att_width 	= 	Upload.Form(FormName & "_Width")
				Att_height 	= 	Upload.Form(FormName & "_Height")
				Att_upName 	= 	Upload.Form(FormName & "_Name")
				If Upload.Form(FormName & "_Err") = 0 Then
					oldPath = Sys.doGet("Select Att_Path From blog_Attment Where Att_ID=" & id)(0)
					Set fso = New cls_fso
						fso.DeleteFile(oldPath)
					Set fso = Nothing
					Str = upToBase(id)
					ResponseString  = ResponseString & "{id : " & Str(1) & ", name : '" & Att_upName & "', path : '" & Att_Path & "', size : " & Att_Size & ", type : '" & Att_type & "', ext : '" & Att_Ext & "'},"
				End If
			Next
			ResponseString = Mid(ResponseString, 1, Len(ResponseString) - 1)
			ResponseString = "[" & ResponseString & "]"
			Response.Write("{Suc : true, Info : " & ResponseString & "}")
		End If
	End Sub
	
	Private Sub SingleDel
		On Error Resume Next
		Set Upload = Nothing
		Dim id, oldPath, fso, Rs
		id = Asp.CheckStr(Request.QueryString("id"))
		If Not Asp.IsInteger(id) Then
			Response.Write("{Suc : false, Info : '非法参数'}")
		Else
			Conn.BeginTrans
			Set Rs = Server.CreateObject("Adodb.RecordSet")
			Rs.open "Select * From blog_Attment Where Att_ID=" & id, Conn, 1, 3
			If Rs.Bof Or Rs.Eof Then
				Response.Write("{Suc : false, Info : '找不到该附件'}")
			Else
				Set fso = New cls_fso
					If fso.DeleteFile(Rs("Att_Path").value) Then
						Rs.Delete
						Response.Write("{Suc : true, Info : '删除附件成功'}")
					Else
						Response.Write("{Suc : false, Info : '删除附件失败'}")
					End If
				Set fso = Nothing
			End If
			Rs.Close
			Set Rs = Nothing
			If Conn.Errors.Count > 0 Then
				Conn.RollBackTrans
				Response.Write("{Suc : false, Info : '" & Err.Description & "'}")
			Else
				Conn.CommitTrans
			End If
		End If
	End Sub
	
End Class
%>