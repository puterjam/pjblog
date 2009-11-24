<%
Dim plus
Set plus = New Sys_Plus
Class Sys_Plus

	Private PlusArray, PluginMark
	
	' ***********************************************
	'	类初始化
	' ***********************************************
	Private Sub Class_Initialize
		If Not IsArray(Application(Sys.CookieName & "_plus_" & blog_DefaultSkin)) Then
			Reload
		Else
			PlusArray = Application(Sys.CookieName & "_plus_" & blog_DefaultSkin)
		End If
    End Sub 
     
	' ***********************************************
	'	类终结化
	' ***********************************************
    Private Sub Class_Terminate
    End Sub
	
	Public Sub open(ByVal PluginNameMark)
		PluginMark = PluginNameMark
	End Sub
	
	Public Sub Reload
		Dim plusCount, plusArray, plusRs, i
		plusCount = Int(Sys.doGet("Select Count(*) From blog_tempPlugin Where tp_templateName='" & blog_DefaultSkin & "'")(0))
		If plusCount > 0 Then
			Set plusRs = Conn.Execute("Select tp_ID, tp_pluginSingleMark, tp_pluginSinglePath, tp_pluginSingleTempPath, tp_pluginPath, tp_pluginSingleName, tp_PluginMark, tp_pluginSingleTempValue From blog_tempPlugin Where tp_templateName='" & blog_DefaultSkin & "'")
			If plusRs.Bof Or plusRs.Eof Then
				ReDim plusArray(0, 0)
			Else
				plusArray = plusRs.GetRows
			End If
			Set plusRs = Nothing
		End If
		Application.Lock()
		Application(Sys.CookieName & "_plus_" & blog_DefaultSkin) = plusArray
		Application.UnLock()
	End Sub
	
	Public Function getPlusID(ByVal Str)
		Dim i
		getPluID = 0
		If Ubound(PlusArray, 1) > 0 Then
			For i = 0 To UBound(PlusArray, 2)
				If PlusArray(6, i) = PluginMark Then ' 外标识
					If PlusArray(1, i) = Str Then '  内标识
						getPlusID = PlusArray(0, i)
						Exit For
					End If
				End If
			Next
		End If
	End Function
	
	Public Function getSinglePath(ByVal Str)
		Dim i
		getSinglePath = ""
		If Ubound(PlusArray, 1) > 0 Then
			For i = 0 To UBound(PlusArray, 2)
				If PlusArray(6, i) = PluginMark Then ' 外标识
					If PlusArray(1, i) = Str Then '  内标识
						getSinglePath = PlusArray(2, i)
						Exit For
					End If
				End If
			Next
		End If
	End Function
	
	Public Function getSingleTempPath(ByVal Str)
		Dim i
		getSingleTempPath = ""
		If Ubound(PlusArray, 1) > 0 Then
			For i = 0 To UBound(PlusArray, 2)
				If PlusArray(6, i) = PluginMark Then ' 外标识
					If PlusArray(1, i) = Str Then '  内标识
						getSingleTempPath = PlusArray(3, i)
						Exit For
					End If
				End If
			Next
		End If
	End Function
	
	Public Function getPluginPath(ByVal Str)
		Dim i
		getPluginPath = ""
		If Ubound(PlusArray, 1) > 0 Then
			For i = 0 To UBound(PlusArray, 2)
				If PlusArray(6, i) = PluginMark Then ' 外标识
					If PlusArray(1, i) = Str Then '  内标识
						getPluginPath = PlusArray(4, i)
						Exit For
					End If
				End If
			Next
		End If
	End Function
	
	Public Function getSingleName(ByVal Str)
		Dim i
		getSingleName = ""
		If Ubound(PlusArray, 1) > 0 Then
			For i = 0 To UBound(PlusArray, 2)
				If PlusArray(6, i) = PluginMark Then ' 外标识
					If PlusArray(1, i) = Str Then '  内标识
						getSingleName = PlusArray(5, i)
						Exit For
					End If
				End If
			Next
		End If
	End Function
	
	Public Function getSingleTemplate(ByVal Str)
		Dim i
		getSingleTemplate = ""
		If Ubound(PlusArray, 1) > 0 Then
			For i = 0 To UBound(PlusArray, 2)
				If PlusArray(6, i) = PluginMark Then ' 外标识
					If PlusArray(1, i) = Str Then '  内标识
						getSingleTemplate = PlusArray(7, i)
						Exit For
					End If
				End If
			Next
		End If
	End Function
	
	Public Function RemovePlusApplication
		Application.Lock
        Application.Contents.Remove(Sys.CookieName & "_plus_" & blog_DefaultSkin)
        Application.UnLock
	End Function
	
End Class
%>