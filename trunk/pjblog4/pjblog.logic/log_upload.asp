<!--#include file = "../include.asp" -->
<!--#include file = "../pjblog.model/cls_upload.asp" -->
<%
Dim doUpload
Set doUpload = New Sys_Upload
Set doUpload = Nothing

Class Sys_Upload

	Private Upload, successful, thisFile, allFiles, upPath, path,file, tempCls

	Private Att_ID, Att_Path, Att_Date, Att_Size, Att_type, Blog_ID
	
	Private OK, Arrays, GoBase, back

	Private Sub Class_Initialize()
'		On Error Resume Next
		Server.ScriptTimeout = 9999999
		Set Upload = New SysUpLoad
			Upload.openProcesser = true  '打开进度条显示
			Upload.SingleSize = 512*1024*1024  '设置单个文件最大上传限制,按字节计；默认为不限制，本例为512M
			Upload.MaxSize = Int(UP_FileSize) '设置最大上传限制,按字节计；默认为不限制，本例为1G
			Upload.Exe = Lcase(UP_FileTypes)'"bmp|jpg|gif|png|rar|doc|pdf|txt"  '设置允许上传的扩展名
			Upload.GetData()
		If Upload.ErrorID > 0 Then
			upload.setApp "faild", 1, 0, Upload.description
		Else
			If Upload.files(-1).Count > 0 Then
				Dim str
				For Each file In Upload.files(-1)
					upPath = Request.Querystring("path")
					path = Server.Mappath(upPath)
					Set tempCls = Upload.files(file)
					upload.setApp "saving", Upload.TotalSize, Upload.TotalSize, tempCls.FileName
					successful = tempCls.SaveToFile(path,0)
					' -----------------------------------------------
					'	插入数据库 开始
					' -----------------------------------------------
					Att_Path = upPath & "/" & tempCls.FileName
					Att_Size = tempCls.Size
					Att_type = Mid(Att_Path, InStrRev(Att_Path, ".") + 1, Len(Att_Path))
					GoBase = InToBase
					' -----------------------------------------------
					'	插入数据库 结束
					' -----------------------------------------------
					If GoBase(0) Then back = GoBase(1) Else back = 0
					If CheckExe(Att_type) Then
						thisFile = "{name:'" & tempCls.FileName & "',size:" & tempCls.Size & ",path:'../download.asp?id=" & back & "',attid:" & back & "}"
					Else
						thisFile = "{name:'" & tempCls.FileName & "',size:" & tempCls.Size & ",path:'" & upPath & "/" & tempCls.FileName & "',attid:" & back & "}"
					End If
					allFiles = allFiles & thisFile & ","
					Set tempCls = Nothing
				Next
				upload.setApp "saved", Upload.TotalSize, Upload.TotalSize, allFiles
			Else
				upload.setApp "faild", 1, 0, "没有上传任何文件"
			End If
		End if
		If err Then upload.setApp "faild", 1, 0, Err.Description
	End Sub
	
	Private Sub Class_Terminate()
	  Set Upload = Nothing
	  Sys.Close
	End Sub

	Private Function InToBase
		Arrays = Array(Array("Att_Path", Att_Path), Array("Att_Date", Now()), Array("Att_Size", Att_Size), Array("Att_type", Att_type))
		OK = Sys.doRecord("blog_Attment", Arrays, "insert", "Att_ID", "")
		InToBase = OK
	End Function
	
	Private Function CheckExe(ByVal Str)
		CheckExe = True
		Dim Arr, i
		Arr = Array("jpg", "gif", "png", "bmp", "jpeg")
		For Each i In Arr
			If Lcase(Str) = i Then
				CheckExe = False
				Exit Function
			End If
		Next
	End Function
	
End Class
%>