<%
' ======================================================================================
'	全部使用方法:
'	Dim c, page
'	Set c = New template                                             	创建新对象实例
'		c.Path = "tmp"													设置模板文件夹位置
'		c.FileName = "temp.html"										设置模板文件名
'		c.open															打开模板
'		c.Sets("dd") = "123"											执行普通替换
'		c.PageUrl = "?page={$page}"										分页URL样式
'		page = Trim(Asp.CheckStr(Request.QueryString("page")))			获取当前页
'		c.CurrentPage = Asp.CheckPage(page)								设置当前页
'		c.Buffer														执行缓冲
'		c.Flush															输出页面
'	Set c = Nothing														实例终结
'	Sys.Close															关闭数据库
' ======================================================================================
Class template

	Private c_Char, c_Path, c_FileName, c_Content, c_PageUrl, c_CurrentPage, c_PageStr
	Private TagName
	
	' ***************************************
	'	设置编码
	' ***************************************
	Public Property Let Char(ByVal Str)
		c_Char = Str
	End Property
	Public Property Get Char
		Char = c_Char
	End Property
	
	' ***************************************
	'	设置模板文件夹路径
	' ***************************************
	Public Property Let Path(ByVal Str)
		c_Path = Str
	End Property
	Public Property Get Path
		Path = c_Path
	End Property
	
	' ***************************************
	'	设置模板文件名
	' ***************************************
	Public Property Let FileName(ByVal Str)
		c_FileName = Str
	End Property
	Public Property Get FileName
		FileName = c_FileName
	End Property
	
	' ***************************************
	'	获得模板文件具体路径
	' ***************************************
	Public Property Get FilePath
		If Len(Path) > 0 Then Path = Replace(Path, "\", "/")
		If Right(Path, 1) <> "/" Then Path = Path & "/"
		FilePath = Path & FileName
	End Property
	
	' ***************************************
	'	设置分页URL
	' ***************************************
	Public Property Let PageUrl(ByVal Str)
		c_PageUrl = Str
	End Property
	Public Property Get PageUrl
		PageUrl = c_PageUrl
	End Property
	
	' ***************************************
	'	设置分页 当前页
	' ***************************************
	Public Property Let CurrentPage(ByVal Str)
		c_CurrentPage = Str
	End Property
	Public Property Get CurrentPage
		CurrentPage = c_CurrentPage
	End Property
	
	' ***************************************
	'	输出内容
	' ***************************************
	Public Property Get Flush
		Response.Write(c_Content)
	End Property
	
	' ***************************************
	'	类初始化
	' ***************************************
	Private Sub Class_Initialize
		TagName = "pjblog"
		c_Char = "UTF-8"
	End Sub
	
	' ***************************************
	'	过滤冲突字符
	' ***************************************
	Private Function doQuote(ByVal Str)
		doQuote = Replace(Str, Chr(34), "&quot;")
	End Function
	
	' ***************************************
	'	类终结
	' ***************************************
	Private Sub Class_Terminate
	End Sub
	
	' ***************************************
	'	加载文件方法
	' ***************************************
	Private Function LoadFromFile(ByVal cPath)
		Dim obj
		Set obj = Server.CreateObject("ADODB.Stream")
			With obj
			    .Type = 2
				.Mode = 3
				.Open
				.Charset = Char
				.Position = .Size
				.LoadFromFile Server.MapPath(cPath)
				LoadFromFile = .ReadText
				.close
			End With
		Set obj = Nothing
	End Function
	
	' ***********************************************
	'	获取正则匹配对象
	' ***********************************************
	Public Function GetMatch(ByVal Str, ByVal Rex)
		Dim Reg, Mag
		Set Reg = New RegExp
		With Reg
			.IgnoreCase = True
			.Global = True
			.Pattern = Rex
			Set Mag = .Execute(Str)
			If Mag.Count > 0 Then
				Set GetMatch = Mag
			Else
				Set GetMatch = Server.CreateObject("Scripting.Dictionary")
			End If
		End With
		Set Reg = nothing
	End Function
	
	' ***************************************
	'	打开文档
	' ***************************************
	Public Sub open
		c_Content = LoadFromFile(FilePath)
	End Sub
	
	' ***************************************
	'	缓冲执行
	' ***************************************
	Public Sub Buffer
		c_Content = GridView(c_Content)
		Call ExecuteFunction
	End Sub
	
	' ***************************************
	'	GridView
	' ***************************************
	Private Function GridView(ByVal o_Content)
		Dim Matches, SubMatches, SubText
		Dim Attribute, Content
		Set Matches = GetMatch(o_Content, "\<" & TagName & "\:(\d+?)(.+?)\>([\s\S]+?)<\/" & TagName & "\:\1\>")
		If Matches.Count > 0 Then
			For Each SubMatches In Matches
				Attribute = SubMatches.SubMatches(1) ' kocms
				Content = SubMatches.SubMatches(2)   ' <Columns>...</Columns>
				SubText = Process(Attribute, Content)
				o_Content = Replace(o_Content, SubMatches.value, "<" & SubText(2) & SubText(0) & ">" & SubText(1) & "</" & SubText(2) & ">")
			Next
		End If
		Set Matches = Nothing
		GridView = o_Content
	End Function
	
	' ***************************************
	'	确定属性
	' ***************************************
	Private Function Process(ByVal Attribute, ByVal Content)
		Dim Matches, SubMatches, Text
		Dim MatchTag, MatchContent
		Dim datasource, Name, Element, page, id
		datasource = "" : Name = "" : Element = "" : page = 0 : id = ""
		Set Matches = GetMatch(Attribute, "\s(.+?)\=\""(.+?)\""")
		If Matches.Count > 0 Then
			For Each SubMatches In Matches
				MatchTag = SubMatches.SubMatches(0)
				MatchContent = SubMatches.SubMatches(1)
				If Lcase(MatchTag) = "name" Then Name = MatchContent
				If Lcase(MatchTag) = "datasource" Then datasource = MatchContent
				If Lcase(MatchTag) = "element" Then Element = MatchContent
				If Lcase(MatchTag) = "page" Then page = MatchContent
				If Lcase(MatchTag) = "id" Then id = MatchContent
			Next
			If Len(Name) > 0 And Len(MatchContent) > 0 Then
				Text = Analysis(datasource, Name, Content, page, id)
				If Len(datasource) > 0 Then Attribute = Replace(Attribute, "datasource=""" & datasource & """", "")
				If page > 0 Then Attribute = Replace(Attribute, "page=""" & page & """", "")
				Attribute = Replace(Attribute, "name=""" & Name & """", "")
				Attribute = Replace(Attribute, "element=""" & Element & """", "")
				Process = Array(Attribute, Text, Element)
			Else
				Process = Array(Attribute, "", "div")
			End If
		Else
			Process = Array(Attribute, "", "div")
		End If
		Set Matches = Nothing
	End Function
	
	' ***************************************
	'	解析
	' ***************************************
	Private Function Analysis(ByVal id, ByVal Name, ByVal Content, ByVal page, ByVal PageID)
		Dim Data
		Select Case Lcase(Name)
			Case "loop" Data = DataBind(id, Content, page, PageID)
			Case "for"  Data = DataFor(id, Content, page, PageID)
		End Select
		Analysis = Data
	End Function
	
	' ***************************************
	'	绑定数据源
	' ***************************************
	Private Function DataBind(ByVal id, ByVal Content, ByVal page, ByVal PageID)
		Dim Text, Matches, SubMatches, SubText
		Execute "Text = " & id & "(1)"
		Set Matches = GetMatch(Content, "\<Columns\>([\s\S]+)\<\/Columns\>")
		If Matches.Count > 0 Then
			For Each SubMatches In Matches
				SubText = ItemTemplate(SubMatches.SubMatches(0), Text, page, PageID)
				Content = Replace(Content, SubMatches.value, SubText)
			Next
			DataBind = Content
		Else
			DataBind = ""
		End If
		Set Matches = Nothing
	End Function
	
	' ***************************************
	'	匹配模板实例
	' ***************************************
	Private Function ItemTemplate(ByVal TextTag, ByVal Text, ByVal page, ByVal PageID)
		Dim Matches, SubMatches, SubMatchText
		Dim SecMatch, SecSubMatch
		Dim i, TempText
		Dim TextLen, TextLeft, TextRight
		Set Matches = GetMatch(TextTag, "\<ItemTemplate\>([\s\S]+)\<\/ItemTemplate\>")
		If Matches.Count > 0 Then
			For Each SubMatches In Matches
				SubMatchText = SubMatches.SubMatches(0)
				' ---------------------------------------------
				'	循环嵌套开始
				' ---------------------------------------------
				SubMatchText = GridView(SubMatchText)
				' ---------------------------------------------
				'	循环嵌套结束
				' ---------------------------------------------
				If UBound(Text, 1) = 0 Then
					TempText = ""
				Else
					TempText = ""
					' -----------------------------------------------
					'	开始分页
					' -----------------------------------------------
					If Len(page) > 0 And page > 0 Then
						If Len(CurrentPage) = 0 Or CurrentPage = 0 Then CurrentPage = 1
						TextLen = UBound(Text, 2)
						TextLeft = (CurrentPage - 1) * page
						TextRight = CurrentPage * page - 1
						If TextLeft < 0 Then TextLeft = 0
						If TextRight > TextLen Then TextRight = TextLen
						c_PageStr = MultiPage(TextLen + 1, page, CurrentPage, PageUrl, "float:right", "", False)
						If Len(c_PageStr) > 0 Then
							c_Content = Replace(c_Content, "<page:" & PageID & "/>", c_PageStr)
						Else
							c_Content = Replace(c_Content, "<page:" & PageID & "/>", "")
						End If
					Else
						TextLeft = 0
						TextRight = UBound(Text, 2)
					End If
					
					For i = TextLeft To TextRight
						TempText = TempText & ItemReSec(i, SubMatchText, Text)
					Next
				End If
			Next
			ItemTemplate = TempText
		Else
			ItemTemplate = ""
		End If
		Set Matches = Nothing
	End Function
	
	' ***************************************
	'	替换模板字符串
	' ***************************************
	Private Function ItemReSec(ByVal i, ByVal Text, ByVal Arrays)
		Dim Matches, SubMatches
		Set Matches = GetMatch(Text, "\$(\d+?)")
		If Matches.Count > 0 Then
			For Each SubMatches In Matches
				Text = Replace(Text, SubMatches.value, doQuote(Arrays(SubMatches.SubMatches(0), i)))
			Next
			ItemReSec = Text
		Else
			ItemReSec = ""
		End If
		Set Matches = Nothing
	End Function
	
	' ***************************************
	'	全局变量函数
	' ***************************************
	Private Sub ExecuteFunction
		Dim Matches, SubMatches, Text, ExeText
		Set Matches = GetMatch(c_Content, "\<function\:([0-9a-zA-Z_\.]*?)\((.*?)\""(.+?)\""(.*?)\)\/\>")
		If Matches.Count > 0 Then
			For Each SubMatches In Matches
				Text = SubMatches.SubMatches(0) & "(" & SubMatches.SubMatches(1) & """" & SubMatches.SubMatches(2) & """" & SubMatches.SubMatches(3) & ")"
				Execute "ExeText=" & Text
				c_Content = Replace(c_Content, SubMatches.value, ExeText)
			Next
		End If
		Set Matches = Nothing
	End Sub
	
	' ***************************************
	'	普通替换全局标签
	' ***************************************
	Public Property Let Sets(ByVal t, ByVal s)
		Dim SetMatch, Bstr, SetSubMatch
		Set SetMatch = GetMatch(c_Content, "(\<Set\:([0-9a-zA-Z_\.]*?)\(((.*?)" + t + "(.*?))?\)\/\>)")
		If SetMatch.Count > 0 Then
			For Each SetSubMatch In SetMatch
				Execute "Bstr = " & SetSubMatch.SubMatches(1) & "(" & SetSubMatch.SubMatches(3) & """" & s & """" & SetSubMatch.SubMatches(4) & ")"
				c_Content = Replace(c_Content, SetSubMatch.Value, Bstr)
			Next
		End If
		Set SetMatch = Nothing
		Set SetMatch = GetMatch(c_Content, "(\<Set\:" + t + "\/\>)")
		If SetMatch.Count > 0 Then
			For Each SetSubMatch In SetMatch
				c_Content = Replace(c_Content, SetSubMatch.Value, s)
			Next
		End If
		Set SetMatch = Nothing
	End Property
	
	' ***************************************
	'	普通替换全局标签
	' ***************************************
	Private Function DataFor(ByVal DataSource, ByVal Content, ByVal page, ByVal PageID)
		
	End Function
	
End Class
%>