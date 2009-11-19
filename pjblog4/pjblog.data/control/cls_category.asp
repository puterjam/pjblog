<%
Dim Category
Set Category = New Sys_Category

Class Sys_Category

	Public cate_ID, cate_Order, cate_icon, cate_Name, cate_Intro, cate_Folder, cate_URL, cate_local, cate_Secret, cate_count, cate_OutLink, cate_Lock
	Private Str, Arrays, fso, Rs, i
	
	' ***********************************************
	'	类初始化
	' ***********************************************
	Private Sub Class_Initialize
    End Sub 
     
	' ***********************************************
	'	类终结化
	' ***********************************************
    Private Sub Class_Terminate
    End Sub
	
	Public Function Add
		Dim DirPath, FolderExsit, FolderJoin, FolderJoinArray, i
		FolderExsit = False
		Set fso = New cls_fso
			FolderJoin = fso.FolderItem("../../html")
			FolderJoinArray = Split(FolderJoin, "|")
			For i = 1 To UBound(FolderJoinArray)
				If Lcase(Trim(FolderJoinArray(i))) = Lcase(Trim(cate_Folder)) Then FolderExsit = True
			Next
			If FolderExsit Then
				Add = Array(False, "该目录已存在!")
			Else
				Arrays = Array(Array("cate_Order", cate_Order), Array("cate_icon", cate_icon), Array("cate_Name", cate_Name), Array("cate_Intro", cate_Intro), Array("cate_Folder", cate_Folder), Array("cate_URL", cate_URL), Array("cate_local", cate_local), Array("cate_Secret", cate_Secret), Array("cate_count", cate_count), Array("cate_OutLink", cate_OutLink), Array("cate_Lock", cate_Lock))
				Add = Sys.doRecord("blog_Category", Arrays, "insert", "cate_ID", "")
				DirPath = "../../html/" & cate_Folder
				If Not cate_OutLink Then fso.CreateFolder(DirPath)
				Call Data.NavList(2)
				Call Cache.CategoryCache(2)
			End If
		Set fso = Nothing
	End Function
	
	Public Function edit
		Dim Check, ReNameBool, files, fileSplit, OutLink
		OutLink = True
		' --------------------------- 判断该分类是否存在 ----------------------------
		Set Rs = Conn.Execute("Select * From blog_Category Where cate_ID=" & cate_ID)
			If Rs.Bof Or Rs.Eof Then
				Check = Array(False, cate_ID & "号分类不存在")
			Else
				Check = Array(True, Trim(Rs("cate_Folder").value))
				OutLink = Rs("cate_OutLink")
			End If
		Rs.Close
		Set Rs = Nothing
		' --------------------------- 判断文件夹是否存在 ----------------------------
		If Int(Sys.doGet("Select Count(*) From blog_Category Where cate_Folder='" & Trim(cate_Folder) & "' And cate_Folder <> '' And cate_OutLink=False")(0)) > 1 Then
			Check = Array(False, cate_Folder & "目录已存在!")
		End If
		If Check(0) Then
			If Not OutLink Then
				' --------------------------- 判断是否改动文件夹名 ----------------------------
				If Trim(Check(1)) <> Trim(cate_Folder) Then
					Set fso = New cls_fso
						' --------------------------- 判断文件夹名是否为空 ----------------------------
						If Len(Trim(cate_Folder)) = 0 Then
							If Len(Trim(Check(1))) > 0 Then
								files = fso.FileItem("../../html/" & Check(1))
								fileSplit = Split(files, "|")
								For i = 1 To UBound(fileSplit)
									fso.MoveFile "../../html/" & Check(1) & "/" & fileSplit(i), "../../html/" & fileSplit(i)
								Next
							End If
						Else
							fso.FolderRename "../../html/" & Check(1), cate_Folder
						End If
					Set fso = Nothing
				End If
			End If
			Arrays = Array(Array("cate_Order", cate_Order), Array("cate_icon", cate_icon), Array("cate_Name", cate_Name), Array("cate_Intro", cate_Intro), Array("cate_Folder", cate_Folder), Array("cate_URL", cate_URL), Array("cate_local", cate_local), Array("cate_Secret", cate_Secret), Array("cate_count", cate_count), Array("cate_OutLink", cate_OutLink))
			edit = Sys.doRecord("blog_Category", Arrays, "update", "cate_ID", cate_ID)
			Call Data.NavList(2)
			Call Cache.CategoryCache(2)
		Else
			edit = Array(False, Check(1))
		End If
	End Function
	
	Public Function del
		Dim Check, OutLink
		OutLink = True
		Set Rs = Conn.Execute("Select * From blog_Category Where cate_ID=" & cate_ID)
		If Rs.Bof Or Rs.Eof Then
			Check = Array(False, cate_ID & "号分类不存在")
		Else
			Check = Array(True, Rs("cate_Folder").value)
			OutLink = Rs("cate_OutLink")
		End If
		Set Rs = Nothing
		If Check(0) Then
			Set fso = New cls_fso
				If Not OutLink Then
					If Len(Trim(Check(1))) = 0 Then
						Set Rs = Conn.Execute("Select log_cname, log_ctype From blog_Content Where log_CateID=" & Int(cate_ID))
						If Not (Rs.Bof And Rs.Eof) Then
							Do While Not Rs.Eof
								fso.DeleteFile("../../html/" & Trim(Rs("log_cname").value) & "." & Trim(Rs("log_ctype").value))
							Rs.Movenext
							Loop
						End If
						Set Rs = Nothing
					Else
						fso.DeleteFolder "../../html/" & Check(1), True
					End If
				End If
			Set fso = Nothing
			Arrays = Array(Array("blog_Category", "cate_ID", cate_ID))
			del = Sys.doRecDel(Arrays)
			Call Data.NavList(2)
			Call Cache.CategoryCache(2)
		Else
			del = Array(False, Check(1))
		End If
	End Function
	
End Class
%>