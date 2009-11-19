<%
' *****************************************************
' 	分类管理
' *****************************************************
Dim Plugin_Category, Plugin_Category_String, Plugin_loop_i
Plugin_Category = Cache.CategoryCache(1)
If UBound(Plugin_Category, 1) = 0 Then
	Plugin_Category_String = ""
Else
	Plugin_Category_String = ""
	For Plugin_loop_i = 0 To UBound(Plugin_Category, 2)
		If Int(Plugin_Category(9, Plugin_loop_i)) = 0 Or Int(Plugin_Category(9, Plugin_loop_i)) = 2 Then
			Plugin_Category_String = Plugin_Category_String & "<img src=""" & Plugin_Category(6, Plugin_loop_i) & """ border=""0"" style=""margin:3px 4px -4px 0px;"" alt=""" & Plugin_Category(3, Plugin_loop_i) & """/>"
			If Plugin_Category(4, Plugin_loop_i) Then
				Plugin_Category_String = Plugin_Category_String & "<a class=""CategoryA"" href=""" & Plugin_Category(5, Plugin_loop_i) & """ title=""" & Plugin_Category(3, Plugin_loop_i) & """>" & Plugin_Category(1, Plugin_loop_i) & "</a><br/>"
			Else
				Plugin_Category_String = Plugin_Category_String & "<a class=""CategoryA"" href=""cate_" & Plugin_Category(11, Plugin_loop_i) & "_1.html"" title=""" & Plugin_Category(3, Plugin_loop_i) & """>" & Plugin_Category(1, Plugin_loop_i) & "</a><br/>"
			End If
		End If
	Next
End If
%>
	PluginTempValue = ['plugin_Category', '<%=outputSideBar(Plugin_Category_String)%>'];
	PluginOutPutString.push(PluginTempValue);