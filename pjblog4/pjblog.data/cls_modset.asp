<%
Dim Model
Set Model = New ModSet

Class ModSet

	Private this, ModArray, i, j, ModName
	Public GetError
	' @ -18901 : 数据不存在
	' @ -18902 : 缓存不存在
	' @ -18903 : 插件名为空

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
	
	Public Sub Open(ByVal Str)
		ModName = Str
		If Not IsArray(Application(Sys.CookieName & "_mod_" & ModName)) Then
			GetError = -18902
			ReLoad
		Else
			ModArray = Application(Sys.CookieName & "_mod_" & ModName)
		End If
	End Sub
	
	Public Sub ReLoad
		Dim Folder, Counts
		If len(ModName) = 0 Then
			GetError = -18903
			Exit Sub
		End If
		Folder = Sys.doGet("Select plugin_folder From blog_Module Where plugin_Mark='" & ModName & "'")(0)
		Set this = Server.CreateObject("Adodb.RecordSet")
		this.open "Select * From blog_modsetting Where xml_name='" & ModName & "'", Conn, 1, 1
		If this.Bof Or this.Eof Then
			GetError = -18901
			ReDim ModArray(0, 0)
		Else
			i = 0
			Counts = this.RecordCount
			ReDim ModArray(Counts, 1)
			Do While Not this.Eof
				ModArray(i, 0) = this("xml_key").value
				ModArray(i, 1) = this("xml_value").value
				i = i + 1
			this.MoveNext
			Loop
			ModArray(Counts, 0) = "PluginPath"
			ModArray(Counts, 1) = Folder
		End If
		Application.Lock()
		Application(Sys.CookieName & "_mod_" & ModName) = ModArray
		Application.UnLock()
		this.Close
		Set this = Nothing
	End Sub
	
	Public Function getKeyvalue(ByVal Str)
		For j = 0 To UBound(ModArray, 1)
			If ModArray(j, 0) = Str Then
				getKeyvalue = ModArray(j, 1)
				Exit For
			End If
		Next
	End Function
	
	Public Function GetPath
		GetPath = ModArray(UBound(ModArray, 1), 1)
	End Function
	
	Public Function RemoveApplication
        Application.Lock
        Application.Contents.Remove(Sys.CookieName & "_mod_" & ModName)
        Application.UnLock
    End Function

End Class
%>