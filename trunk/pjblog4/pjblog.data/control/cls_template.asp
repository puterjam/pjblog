<%
Dim Temp
Set Temp = New Sys_Template

Class Sys_Template

	Private Arrays
	
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
	
	Public Function doUpdate(ByVal f1, ByVal f2)
		Arrays = Array(Array("blog_DefaultSkin", f1), Array("blog_style", f2))
		doUpdate = Sys.doRecord("blog_Info", Arrays, "update", "blog_ID", 1)
	End Function
	
	Public Function AddPlus(f1, tp_pluginSingleMark, tp_pluginSinglePath, tp_pluginSingleTempPath, folder, tp_pluginSingleName, Plugin_Mark, tp_pluginSingleTempValue, tp_plugintag)
		Arrays = Array(Array("tp_templateName", Trim(f1)), Array("tp_pluginSingleMark", Trim(tp_pluginSingleMark)), Array("tp_pluginSinglePath", Trim(tp_pluginSinglePath)), Array("tp_pluginSingleTempPath", Trim(tp_pluginSingleTempPath)), Array("tp_pluginPath", Trim(folder)), Array("tp_pluginSingleName", Trim(tp_pluginSingleName)), Array("tp_PluginMark", Trim(Plugin_Mark)), Array("tp_pluginSingleTempValue", Trim(tp_pluginSingleTempValue)), Array("tp_plugintag", Trim(tp_plugintag)))
		AddPlus = Sys.doRecord("blog_tempPlugin", Arrays, "insert", "tp_ID", "")
	End Function
	
End Class
%>