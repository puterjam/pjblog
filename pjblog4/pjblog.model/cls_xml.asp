<%
Class xml

	Private e_FilePath, e_Str, e_xml, e_xmlDoc

' =============================================
' 	设置属性
' =============================================
	Public property LET FilePath(ByVal values)
		e_FilePath = Server.MapPath(values)
	End property
	Public property Get FilePath()
		FilePath = e_FilePath
	End property
	
	Public property LET Str(ByVal values)
		e_Str = values
	End property
	Public property Get Str()
		Str = e_Str
	End property
	
' =============================================
' 	取得根节点
' =============================================
	Public Property Get Root 
		Set Root = e_xmlDoc 
	End Property 
	
' =============================================
' 	取得根节点下节点数
' ============================================= 
	Property Get Length
		Length = e_xmlDoc.childNodes.length 
	End Property 

' =============================================
' 	类初始化和类的销毁
' ============================================= 		
    Private Sub Class_Initialize()
       	Set e_xml = Server.CreateObject("Microsoft.XMLDOM")
			e_xml.async = False
    End Sub

    Private Sub Class_Terminate()
		Set e_xml = Nothing
    End Sub
	
' =============================================
' 	从外部读入XML文档
' ============================================= 
	Public Function open
		open = False
		If e_xml.load(FilePath) Then
			Set e_xmlDoc = e_xml.documentElement
			open = True
		End If
	End Function

' =============================================
' 	从外部读入XML字符串
' ============================================= 
	Public Function openxml
		openxml = False
		If e_xml.loadXML(Str) Then
			Set e_xmlDoc = e_xml.documentElement
			openxml = True
		End If
	End Function
	
' =============================================
' 	新增一个子节点
' ============================================= 
	Public Function AddChildNode(ByVal NewNodeType, ByRef ParentNode)
		' NewNodeType string 新节点类型
		' ParentNode object 父节点的引用
		Dim TempNode
		Set TempNode = e_xml.createElement(NewNodeType)
		ParentNode.appendChild TempNode
		Set AddChildNode = TempNode
	End Function

' =============================================
' 	给节点增加属性
' ============================================= 	
	Public Function AddAttribute(ByRef Node , ByVal AttributeName, ByVal AttributeValue) 
		' AttributeName STRING 属性名称 
		' AttributeValue STRING 属性值 
		' Node OBJECT 增加属性的对象 
		Node.setAttribute AttributeName, AttributeValue 
	End Function 
	
' =============================================
' 	新增节点内容
' ============================================= 	
	Public Function AddText(ByRef Node, ByVal Tstr, ByVal CDBool) 
		' Node Object 节点的引用
		' Tstr string 内容
		' CDBool string 是否为 CDATA 内容
		AddText = False
		Dim TempText 
		If CDBool Then 
			Set TempText = e_xml.createCDataSection(Tstr) 
		Else 
			Set TempText = e_xml.createTextNode(Tstr) 
		End If 
		Node.appendChild TempText 
		AddText = True
	End Function 
	
' =============================================
' 	取得节点指定属性值
' ============================================= 	
	Public Function GetAttribute(ByRef Node, ByVal AttributeName)
		' AttributeName STRING 属性名称 
		' Node OBJECT 节点引用
		dim TempValue 
		TempValue = Node.getAttribute(AttributeName) 
		GetAttribute = TempValue 
	End Function
	
' =============================================
' 	取得节点名称
' ============================================= 	
	Public Function GetNodeName(ByRef Node)
		' Node OBJECT 节点引用 
		GetNodeName = Node.nodeName 
	End Function
	
' =============================================
' 	取得节点内容
' ============================================= 	
	Public Function GetNodeText(ByRef Node) 
		' Node OBJECT 节点引用 
		GetNodeText = Node.childNodes(0).nodeValue 
	End Function 
	
' =============================================
' 	取得节点类型
' ============================================= 	
	Public Function GetNodeType(ByRef Node) 
		' Node OBJECT 节点引用 
		GetNodeType = Node.nodeType 
	End Function 

' =============================================
' 	查找节点名相同的所有节点
' ============================================= 		
	Public Function FindNodes(ByVal NodeName) 
		Dim TempNodes 
		Set TempNodes = e_xml.getElementsByTagName(NodeName) 
		Set FindNodes = TempNodes 
	End Function 
	
' =============================================
' 	查找一个相同节点
' ============================================= 
	Public Function FindNode(ByVal NodeName) 
		Dim TempNodes 
		Set TempNodes = e_xml.selectSingleNode(NodeName) 
		Set FindNode = TempNodes 
	End Function
	
' =============================================
' 	删除一个节点
' ============================================= 
	Public Function DelNode(ByRef NodeName) 
		Dim TempNodes, ParentNode
		Set TempNodes = NodeName
		Set ParentNode = TempNodes.parentNode 
		ParentNode.removeChild(TempNodes) 
	End Function 
	
' =============================================
' 	替换一个节点
' ============================================= 
	Public Function ReplaceNode(ByVal NodeName, ByVal Text, ByVal CDBool) 
		Dim TempNodes, TempText 
		Set TempNodes = FindNode(NodeName)
		'AddText Text,CDBool,TempNodes 
		If CDBool Then 
			Set TempText = e_xml.createCDataSection(Text) 
		Else 
			Set TempText = e_xml.createTextNode(Text) 
		End If 
		TempNodes.replaceChild TempText, TempNodes.firstChild 
	End Function 
	
' =============================================
' 	创建XML声明
' ============================================= 
	Private Function ProcessingInstruction
		Dim objPi 
		Set objPi = e_xml.createProcessingInstruction("xml", "version="&chr(34)&"1.0"&chr(34)&" encoding="&chr(34)&"utf-8"&chr(34)) 
		'//--把xml生命追加到xml文档 
		e_xml.insertBefore objPi, e_xml.childNodes(0) 
	End Function 
	
' =============================================
' 	保存XML文档
' ============================================= 
	Public Function SaveXML() 
		e_xml.save(e_FilePath) 
	End Function 
	
' =============================================
' 	另存XML文档
' ============================================= 
	Public Function SaveAsXML(ByVal SavePath) 
		ProcessingInstruction() 
		e_xml.save(Server.MapPath(SavePath)) 
	End Function
	
	Public Function SaveToXml(ByVal SavePath)
		e_xml.save(Server.MapPath(SavePath))
	End Function

End Class
%>
