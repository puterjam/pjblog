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
' --------------------------------------------------------------------------------------
' 模板实例:
' 支持无限级循环
'<div>
'模板循环实例一:
'	<pjblog:1 width="1000" height="300" style=" border:1px solid #000; background:#ccc" datasource="Cache.UserRight" name="loop" element="ul">
'    <Columns>
'    	<ItemTemplate>
'        	<li>
'            $1
'                <pjblog:2 width="1000" height="300" style=" border:1px solid #000; background:#ccc" datasource="Cache.UserRight" name="loop" element="ul">
'                <Columns>
'                    <ItemTemplate>
'                        <li>$1
'                        
'                        </li>
'                    </ItemTemplate>
'                </Columns>
'                </pjblog:2>
'            </li>
'        </ItemTemplate>
'    </Columns>
'    </pjblog:1>
'    
'模板循环实例二:
'  <pjblog:3 width="1000" height="300" style=" border:1px solid #000; background:#ccc" datasource="Cache.UserRight" name="loop" element="ul">
'    <Columns>
'    	<ItemTemplate>
'        	<li>$0</li>
'        </ItemTemplate>
'    </Columns>
'    </pjblog:3>
'    
'模板循环实例三:
'    <pjblog:4 width="1000" height="300" style=" border:1px solid #000; background:#ccc" datasource="Cache.UserRight" name="loop" element="ul" id="g" page="2">
'    <Columns>
'    	<ItemTemplate>
'        	<li><function:Replace("$0", "SupAdmin", "3")/></li>
'        </ItemTemplate>
'    </Columns>
'    </pjblog:4>
'    <page:g/> 
'    <function:Replace("<set:dd/>", "2", "7")/>
'</div>
' ======================================================================================
Class template

	Private c_Char, c_Path, c_FileName, c_Content, c_PageUrl, c_CurrentPage, c_PageStr, ReplacePageStr, doForLeft, doForRight
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
		ReplacePageStr = Array("", "")
		doForLeft = ""
		doForRight = ""
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
	'	静态化方法
	' ***************************************
	Public Sub Save(ByVal File)
		Dim Obj
		Set Obj = New Stream
			Obj.SaveToFile c_Content, File
		Set Obj = Nothing
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
		Call include
	End Sub
	
	' ***************************************
	'	缓冲执行
	' ***************************************
	Public Sub Buffer
		c_Content = GridView(c_Content)
		c_Content = ExecuteFunction(c_Content)
		Call IfelseEndif
	End Sub
	
	' ***************************************
	'	GridView
	' ***************************************
	Private Function GridView(ByVal o_Content)
		Dim Matches, SubMatches, SubText, LoopLeft, LoopRight
		Dim Attribute, Content
		LoopLeft = "" : LoopRight = ""
		Set Matches = GetMatch(o_Content, "\<" & TagName & "\:(\d+?)(.+?)\>([\s\S]+?)<\/" & TagName & "\:\1\>")
		If Matches.Count > 0 Then
			For Each SubMatches In Matches
				Attribute = SubMatches.SubMatches(1) 	' kocms
				Content = SubMatches.SubMatches(2)   	' <Columns>...</Columns>
				SubText = Process(Attribute, Content) 	' 返回所有过程执行后的结果
				If len(doForLeft) > 0 Then LoopLeft = doForLeft : doForLeft = ""
				If len(doForRight) > 0 Then LoopRight = doForRight : doForRight = ""
				o_Content = Replace(o_Content, SubMatches.value, "<" & SubText(2) & SubText(0) & ">" & LoopLeft & SubText(1) & LoopRight & "</" & SubText(2) & ">", 1, -1, 1)	' 替换标签变量
				LoopLeft = ""										
				LoopRight = ""
			Next
		End If
		Set Matches = Nothing
		If Len(ReplacePageStr(0)) > 0 Then				' 判断是否标签变量有值,如果有就替换掉.
			o_Content = Replace(o_Content, ReplacePageStr(0), ReplacePageStr(1), 1, -1, 1)
			ReplacePageStr = Array("", "")				' 替换后清空该数组变量
		End If
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
				MatchTag = SubMatches.SubMatches(0)								' 取得属性名
				MatchContent = SubMatches.SubMatches(1)							' 取得属性值
				If Lcase(MatchTag) = "name" Then Name = MatchContent			' 取得name属性值
				If Lcase(MatchTag) = "datasource" Then datasource = MatchContent' 取得datasource属性值
				If Lcase(MatchTag) = "element" Then Element = MatchContent		' 取得element属性值
				If Lcase(MatchTag) = "page" Then page = MatchContent			' 取得page属性值
				If Lcase(MatchTag) = "id" Then id = MatchContent				' 取得id属性值
			Next
			If Len(Name) > 0 And Len(MatchContent) > 0 Then
				Text = Analysis(datasource, Name, Content, page, id)			' 执行解析属性
				If Len(datasource) > 0 Then Attribute = Replace(Attribute, "datasource=""" & datasource & """", "")
				If page > 0 Then Attribute = Replace(Attribute, "page=""" & page & """", "")
				Attribute = Replace(Attribute, "name=""" & Name & """", "", 1, -1, 1)
				Attribute = Replace(Attribute, "element=""" & Element & """", "", 1, -1, 1)
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
		Select Case Lcase(Name)													' 选择数据源
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
		Execute "Text = " & id & "(1)"											' 加载数据源
		Set Matches = GetMatch(Content, "\<Columns\>([\s\S]*?)(\<ItemTemplate\>([\s\S]+)\<\/ItemTemplate\>)([\s\S]*?)\<\/Columns\>")
		If Matches.Count > 0 Then
			For Each SubMatches In Matches
				SubText = ItemTemplate(SubMatches.SubMatches(1), Text, page, PageID)' 执行模块替换
				Content = Replace(Content, SubMatches.value, SubText, 1, -1, 1)
				doForLeft = SubMatches.SubMatches(0)
				doForRight = SubMatches.SubMatches(3)
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

						If Int(Len(c_PageStr)) > 0 Then
							ReplacePageStr = Array("<page:" & Trim(PageID) & "/>", c_PageStr)
						Else
							ReplacePageStr = Array("<page:" & Trim(PageID) & "/>", "")
						End If
					Else
						TextLeft = 0
						TextRight = UBound(Text, 2)
					End If
					
					For i = TextLeft To TextRight
						TempText = TempText & ItemReSec(i, SubMatchText, Text)		' 加载模板内容
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
		Dim Matches, SubMatches, TempValue
		Set Matches = GetMatch(Text, "\$([\d^\s^\<^\>^\{^\}^\""]+?)")
		If Matches.Count > 0 Then
			For Each SubMatches In Matches
				TempValue = Arrays(SubMatches.SubMatches(0), i)
				If TempValue <> Null And TempValue <> "" Then TempValue = doQuote(TempValue)
				If Len(TempValue) > 0 Then
					Text = Replace(Text, SubMatches.value, TempValue, 1, -1, 1) '执行替换
				Else
					Text = Replace(Text, SubMatches.value, "", 1, -1, 1) '执行替换
				End If
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
	Private Function ExecuteFunction(ByVal o_Centent)
		Dim Matches, SubMatches, Text, ExeText, val
		Set Matches = GetMatch(o_Centent, "\{function\:([0-9a-zA-Z_\.]*?)\((.+)\)\}")
		If Matches.Count > 0 Then
			For Each SubMatches In Matches
				val = SubMatches.SubMatches(1)
				val = ExecuteFunction(val)
				Text = SubMatches.SubMatches(0) & "(" & SubMatches.SubMatches(1) & ")"
				Execute "ExeText=" & Text
				o_Centent = Replace(o_Centent, SubMatches.value, ExeText, 1, -1, 1)
			Next
		End If
		Set Matches = Nothing
		ExecuteFunction = o_Centent
	End Function
	
	' ***************************************
	'	普通替换全局标签
	' ***************************************
	Public Property Let Sets(ByVal t, ByVal s)
		Dim SetMatch, Bstr, SetSubMatch
		Set SetMatch = GetMatch(c_Content, "(\<Set\:([0-9a-zA-Z_\.]*?)\(((.*?)" & t & "(.*?))?\)\/\>)")
		If SetMatch.Count > 0 Then
			For Each SetSubMatch In SetMatch
				Execute "Bstr = " & SetSubMatch.SubMatches(1) & "(" & SetSubMatch.SubMatches(3) & """" & s & """" & SetSubMatch.SubMatches(4) & ")"
				c_Content = Replace(c_Content, SetSubMatch.Value, Bstr, 1, -1, 1)
			Next
		End If
		Set SetMatch = Nothing
		Set SetMatch = GetMatch(c_Content, "(\<Set\:" & t & "\/\>)")
		If SetMatch.Count > 0 Then
			For Each SetSubMatch In SetMatch
				c_Content = Replace(c_Content, SetSubMatch.Value, s, 1, -1, 1)
			Next
		End If
		Set SetMatch = Nothing
	End Property
	
	' ***************************************
	'	条件语句判断
	' ***************************************
	Public Sub IfelseEndif
		Dim SetMatch, SetSubMatch
		Dim Condition, TrueResult, ElseCondition, FalseResult, CheckTrue, Str
		Set SetMatch = GetMatch(c_Content, "\<if\:(.+?)\>(.*?)(\<else\>(.+?))?\<end\sif\>")
		If SetMatch.Count > 0 Then
			For Each SetSubMatch In SetMatch
				Condition = SetSubMatch.SubMatches(0)
				TrueResult = SetSubMatch.SubMatches(1)
				ElseCondition = SetSubMatch.SubMatches(2)
				FalseResult = SetSubMatch.SubMatches(3)
				'Response.Write(Condition)
				Execute "CheckTrue = " & Condition
				If Len(ElseCondition) > 0 Then
					If CheckTrue Then
						Str = TrueResult
					Else
						Str = FalseResult
					End If
				Else
					If CheckTrue Then
						Str = TrueResult
					Else
						Str = ""
					End If
				End If
				c_Content = Replace(c_Content, SetSubMatch.Value, Str, 1, -1, 1)
			Next
		End If
		Set SetMatch = Nothing
	End Sub
	
	' ***************************************
	'	匹配包含文件并替换
	' ***************************************
	Private Sub include
		Dim SetMatch, SetSubMatch, mPath
		Set SetMatch = GetMatch(c_Content, "\<include\:(.+?)\/\>")
		If SetMatch.Count > 0 Then
			For Each SetSubMatch In SetMatch
				If Len(Path) > 0 Then Path = Replace(Path, "\", "/")
				If Right(Path, 1) <> "/" Then Path = Path & "/"
				mPath = Path & SetSubMatch.SubMatches(0)
				c_Content = Replace(c_Content, SetSubMatch.value, LoadFromFile(mPath), 1, -1, 1)
			Next
		End If
	End Sub
	
End Class
%>