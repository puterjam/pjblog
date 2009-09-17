<%
'**************** 使用实例 *******************
'Dim e, s, r
'Set e = New cls_FSO
'	s = "Response.write ""123"""
'	With e
'	文件操作
'		.CreateFile "default.asp", s
'		.DeleteFile("default.asp")
'		.MoveFile "default.asp", "1/2.asp"
'		.CopyFile "2.asp", "1/2.asp"
'		.FileRename "1/2.asp", "default.asp"
'		Response.Write(.ReadFileContent("1/default.asp"))
'		For each r In .GetFileInfo("1/default.asp")
'			Response.Write(r & "<br />")
'		Next
'	文件夹操作
'		.CreateFolder("2/")
'		.DeleteFolder "1", False
'		.MoveFolder "2", "3/2"
'		.CopyFolder "2", "3/2"
'		.FolderRename "2", "4"
'		Response.Write(.FolderItem("3"))
'		Response.Write(.FileItem("./"))
'	End With
'Set e = Nothing
'**************** 使用实例 *******************
Class cls_FSO

	Private fso, NewObject
	
	Private Sub Class_Initialize
        Set fso = Server.CreateObject("Scripting.FileSystemObject") 
    End Sub 
    
    Private Sub Class_Terminate
        Set fso = Nothing
    End Sub
	
' ***************************************
' 文件操作  开始
' ***************************************
	' 文件是否存在
	Public Function FileExists(ByVal FileDir)
		If fso.FileExists(Server.MapPath(FileDir)) Then
			FileExists = True
		Else
			FileExists = False
		End If
	End Function
	
	' 文件删除
	Public Function DeleteFile(ByVal FileDir)
		If FileExists(FileDir) Then
			fso.DeleteFile Server.MapPath(FileDir)
			DeleteFile = True
		Else
			DeleteFile = False
		End If
	End Function
	
	' 文件创建并写入内容
	Public Function CreateFile(ByVal FileDir, ByVal FileContent)
		If FileExists(FileDir) Then
			CreateFile = False
			Exit Function
		Else
			Set NewObject = fso.CreateTextFile(Server.MapPath(FileDir))
				NewObject.Write FileContent
			Set NewObject = Nothing
			CreateFile = True
		End If
	End Function
	
	' 文件移动
	Public Function MoveFile(ByVal Source, ByVal Destination)
		If FileExists(Source) Then
			fso.MoveFile Server.MapPath(Source), Server.MapPath(Destination)
			MoveFile = True
		Else
			MoveFile = False
		End If
	End Function
	
	' 复制文件
	Public Function CopyFile(ByVal Source, ByVal Destination)
		If FileExists(Source) Then
			fso.CopyFile Server.MapPath(Source), Server.MapPath(Destination)
			CopyFile = True
		Else
			CopyFile = False
		End If
	End Function
	
	' 重命名
	Public Function FileRename(ByVal FileDir, ByVal NewName)
		If FileExists(FileDir) Then
			Set NewObject = fso.GetFile(Server.MapPath(FileDir))
				NewObject.Name = NewName
			Set NewObject = Nothing
			FileRename = True
		Else
			FileRename = False
		End If
	End Function
	
	' 读取文件内容
	Public Function ReadFileContent(ByVal FileDir)
		If FileExists(FileDir) Then
            Set NewObject = fso.OpenTextFile(Server.MapPath(FileDir)) 
            	ReadFileContent = NewObject.ReadAll 
            Set NewObject = Nothing 
        Else 
            ReadFileContent=False 
        End If
	End Function
	
	' 获取文件的详细信息
	Public Function GetFileInfo(ByVal FileDir)
    	Dim File, FileInfo(10)
    	If FileExists(FileDir) Then
        	Set File = fso.GetFile(Server.MapPath(FileDir))
        		FileInfo(0) = File.Size
        		If FileInfo(0) > 1024 Then
          			FileInfo(0) = Round(FileInfo(0) / 1024,2)
        			If FileInfo(0) > 1024 Then
          				FileInfo(0)=Round(FileInfo(0) / 1024,2)
          				FileInfo(0)= FileInfo(0) & " MB"
        			Else
         				FileInfo(0)= FileInfo(0) & " KB"
        			End If
        		Else
          			FileInfo(0)= FileInfo(0) & " Byte"
        		End If
        	FileInfo(1) = LCase(Right(FileDir, 4))
			FileInfo(2) = File.DateCreated
			FileInfo(3) = File.Type
			FileInfo(4) = File.DateLastModified
			FileInfo(5) = File.Path
			FileInfo(6) = "" 'File.ShortPath 部分服务器不支持
			FileInfo(7) = File.Name
			FileInfo(8) = "" 'File.ShortName 部分服务器不支持
			FileInfo(9) = fso.getExtensionName(FileDir)
			FileInfo(10) = File.DateLastModified
    	End If
    	GetFileInfo = FileInfo
	End Function

' ***************************************
' 文件夹操作  开始
' ***************************************
	' 文件夹是否存在
	Public Function FolderExists(ByVal FolderDir)
		If fso.FolderExists(Server.MapPath(FolderDir)) Then
			FolderExists = True
		Else
			FolderExists = False
		End If
	End Function
	
	' 文件夹删除
	Public Function DeleteFolder(ByVal FolderDir, ByVal Force)
		If FolderExists(FolderDir) Then
			fso.DeleteFolder Server.MapPath(FolderDir), Force
			DeleteFolder = True
		Else
			DeleteFolder = False
		End If
	End Function
	
	' 文件夹创建
	Public Function CreateFolder(ByVal FolderDir)
		If FolderExists(FolderDir) Then
			CreateFolder = False
		Else
			fso.CreateFolder(Server.MapPath(FolderDir))
			CreateFolder = True
		End If
	End Function
	
	' 循环创建文件夹
	Public Sub CreateLoopFolder(ByVal FolderName)
		On Error Resume Next
		If Len(FolderName) = 0 Then Exit Sub
		FolderName = Replace(FolderName, "\", "/")
		If Mid(FolderName, Len(FolderName), 1) = "/" Then FolderName = Mid(FolderName, 1, (Len(FolderName) - 1))
		Dim TempFolder, TempArray, i
		TempFolder = ""
		TempArray = Split(FolderName, "/")
		For i = 0 To UBound(TempArray) 
			TempFolder = TempFolder & TempArray(i) & "/"
			CreateFolder(TempFolder)
			If Err <> 0 Then Err.Clear
		Next
	End Sub
	
	' 文件夹移动
	Public Function MoveFolder(ByVal Source, ByVal Destination)
		If FolderExists(Source) Then
			fso.MoveFolder Server.MapPath(Source), Server.MapPath(Destination)
			MoveFolder = True
		Else
			MoveFolder = False
		End If
	End Function
	
	'文件夹复制
	Public Function CopyFolder(ByVal Source, ByVal Destination)
		If FolderExists(Source) Then
			fso.CopyFolder Server.MapPath(Source), Server.MapPath(Destination)
			CopyFolder = True
		Else
			CopyFolder = False
		End If
	End Function
	
	' 重命名文件夹
	Public Function FolderRename(ByVal FolderDir, ByVal NewName)
		If FolderExists(FolderDir) Then
			Set NewObject = fso.GetFolder(Server.MapPath(FolderDir))
				NewObject.Name = NewName
			Set NewObject = Nothing
			FolderRename = True
		Else
			FolderRename = False
		End If
	End Function

' ***************************************
' 文件及文件夹的循环操作  开始
' ***************************************	
	Function FolderItem(ByVal FolderDir) 
    '//文件夹里的文件夹集合 
        If FolderExists(FolderDir) = False Then 
            FolderItem = False 
            Exit Function 
        End If 
        Dim FolderObj, FolderList, F 
        Set FolderObj = fso.GetFolder(Server.MapPath(FolderDir)) 
        Set FolderList = FolderObj.SubFolders 
        	FolderItem=FolderObj.SubFolders.Count'//文件夹总数 
        	For Each F In FolderList 
            	FolderItem = FolderItem & "|" & F.Name 
        	Next 
        Set FolderList=Nothing 
        Set FolderObj=Nothing 
    End Function 

    Function FileItem(ByVal FolderDir) 
    '//文件夹里的文件集合 
        If FolderExists(FolderDir) = False Then 
            FileItem = False 
            Exit Function 
        End If
        Dim FileObj,FileerList,F,FileList 
        Set FileObj = fso.GetFolder(Server.MapPath(FolderDir)) 
        Set FileList = FileObj.Files 
        	FileItem = FileObj.Files.Count'//文件总数 
        	For Each F In FileList 
            	FileItem=FileItem & "|" & F.Name 
        	Next 
        Set FileList=Nothing 
        Set FileObj=Nothing 
    End Function
	
	Public Function getFileIcons(Str)
		Dim FileIcon
		Str = Split(Str, " ")(0)
		If Lcase(Str) = "jpeg" Then Str = "jpg"
		If Str = "层叠样式表文档" Then Str = "css"
		If Lcase(Str) = "jscript" Then Str = "js"
		If FileExists("images/file/"&Str&".gif") Then FileIcon = Str Else FileIcon = "unknow"
		getFileIcons = "<img border=""0"" src=""images/file/"&FileIcon&".gif"" style=""margin:4px 3px -3px 0px""/>"
	End Function

	
	
End Class
%>