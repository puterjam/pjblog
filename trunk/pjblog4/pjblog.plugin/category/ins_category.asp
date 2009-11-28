<%
' *****************************************************
' 	分类管理
' *****************************************************
Dim Plugin_Category, Plugin_Category_String, Plugin_loop_i, temps, temps1, temps2, temps3, tempStr
Plugin_Category = Cache.CategoryCache(1)
If UBound(Plugin_Category, 1) = 0 Then
	Plugin_Category_String = ""
Else
	' --------------------------------------------------
	'	打开插件模板缓存
	' --------------------------------------------------
	plus.open("sys.category")
	temps = Split(Asp.UnCheckStr(Plus.getSingleTemplate("Sidecategory")), "|$|")
	temps1 = temps(0)
	temps2 = temps(1)
	temps3 = temps(2)
	
	' --------------------------------------------------
	'	替换模板
	' --------------------------------------------------
	Plugin_Category_String = ""
	For Plugin_loop_i = 0 To UBound(Plugin_Category, 2)
		tempStr = temps2
		If Int(Plugin_Category(9, Plugin_loop_i)) = 0 Or Int(Plugin_Category(9, Plugin_loop_i)) = 2 Then
			tempStr = Replace(tempStr, "<#icon#>", Plugin_Category(6, Plugin_loop_i))
			tempStr = Replace(tempStr, "<#alt#>", Plugin_Category(3, Plugin_loop_i))
			If Plugin_Category(4, Plugin_loop_i) Then
				tempStr = Replace(tempStr, "<#url#>", Plugin_Category(5, Plugin_loop_i))
			Else
				tempStr = Replace(tempStr, "<#url#>", "cate_" & Plugin_Category(11, Plugin_loop_i) & "_1.html")
			End If
		End If
		tempStr = Replace(tempStr, "<#title#>", Plugin_Category(1, Plugin_loop_i))
		Plugin_Category_String = Plugin_Category_String & tempStr
	Next
	Plugin_Category_String = temps1 & Plugin_Category_String & temps3
End If
%>
	PluginTempValue = ['Sidecategory', '<%=outputSideBar(Plugin_Category_String)%>'];
	PluginOutPutString.push(PluginTempValue);