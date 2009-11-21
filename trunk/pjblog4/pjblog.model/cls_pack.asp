<%
Class Sys_Pack

	Private xDom, tStream, fileList, baseM, upl, fso, cxml, cstream, temp
	Public fileTotal, xmlPath, parentPath, NeedOut, xmlfirstvalue, SubMatch, SetMatch, AllFilePathText
	
	Private Sub Class_Initialize()
		NeedOut = True
		xmlfirstvalue = "<?xml version=""1.0""?><files xmlns:dt=""urn:schemas-microsoft-com:datatypes""></files>"
	   	On Error Resume Next
	   	fileTotal=0
	   	Set baseM = New packbase64
		Set fso = New cls_fso
		Set cxml = New xml
		Set cstream = New Stream
	   	Set tStream = Server.CreateObject("adodb.stream")
	   		If Err<>0 Then
	     		Err.clear
		 		Response.Write(Err.Description)
		 		Response.End
	   		End if
	   tStream.Type = 1
	   tStream.Mode = 3
	End Sub
  
	Private Sub Class_Terminate()
		If IsObject(cstream) Then Set cstream = Nothing
	   If IsObject(cxml) Then Set cxml = Nothing
	   If IsObject(fso) Then Set fso = Nothing
	   If IsObject(xDom) Then Set xDom = nothing
	   If IsObject(tStream) Then Set tStream = nothing
	   If IsObject(baseM) Then Set baseM = nothing
	   If IsObject(upl) Then Set upl = nothing
	End sub
      
	Public Sub open()
	    Dim i
		Set xDom = Server.CreateObject(getXMLDOM)
		xDom.load(Server.MapPath(xmlPath))
		fileTotal = xDom.getElementsByTagName("files/f").length
	end Sub
  
	Public Sub UnPack
		On Error Resume Next
	    Dim tempStr, i
		For i = 0 To (fileTotal - 1)
			tStream.Open
			tStream.Position = 0
			tStream.write baseM.decode(xDom.getElementsByTagName("files/f/fb").item(i).text)
			tempStr = xDom.getElementsByTagName("files/f/fn").item(i).text
			tempStr = parentPath & tempStr
			If NeedOut Then
				Response.Write "<div>" & tempStr & "</div>"
				Response.Flush()
			End If
			fso.WholeFolder(tempStr)
			tStream.SaveToFile Server.MapPath(tempStr), 2
			tStream.Close
		Next
	End sub
	
	Public Function GetDir(ByVal Dir) ' 无/
		Dim DirFileList, DirFolderList, FileSplit, FolderSplit, i, Str, j
		Str = ""
		DirFileList = fso.FileItem(Dir)
		DirFolderList = fso.FolderItem(Dir)
		FileSplit = Split(DirFileList, "|")
		FolderSplit = Split(DirFolderList, "|")
		If Int(FileSplit(0)) > 0 Then
			For i = 1 To UBound(FileSplit)
				If Dir = "." Then
					Str = Str & "<" & FileSplit(i) & ">" ' 返回文件路径
				Else
					Str = Str & "<" & Dir & "/" & FileSplit(i) & ">" ' 返回文件路径
				End If
			Next
		End If
		If Int(FolderSplit(0)) > 0 Then
			For j = 1 To UBound(FolderSplit)
				If Dir = "." Then
					Str = Str & GetDir(FolderSplit(j))
				Else
					Str = Str & GetDir(Dir & "/" & FolderSplit(j))
				End If
			Next
		End If
		GetDir = Str
	End Function
	
	Public Function doPack(ByVal ceePath, ByVal ToPath)
		cxml.Str = xmlfirstvalue
		If cxml.openxml Then
			AllFilePathText = GetDir(ceePath)
			Set SetMatch = Asp.GetMatch(AllFilePathText, "\<(.*?)\>")
				If SetMatch.Count > 0 Then
					Dim RootFiles : Set RootFiles = cxml.Root
					Dim A, B, ContentText
					On Error Resume Next
					For Each SubMatch In SetMatch
						ContentText = "error"
						tStream.Open
						tStream.LoadFromFile Server.MapPath(SubMatch.SubMatches(0))
						tStream.Position = 0
						If Err Then
                			tStream.Close
               		 		Err.Clear
            			End If
						ContentText = tStream.Read
						tStream.Close
						ContentText = baseM.encode(ContentText)
						If NeedOut Then
							Response.Write(SubMatch.SubMatches(0) & "<br />")
							Response.Flush()
						End If
						Set A = cxml.AddChildNode("f", RootFiles)
						Set B = cxml.AddChildNode("fn", A)
							cxml.AddText B, SubMatch.SubMatches(0), False
						Set B = cxml.AddChildNode("fb", A)
							cxml.AddText B, ContentText, False
					Next
					cxml.SaveToXml(ToPath)
					doPack = Array(True, "打包成功!")
				Else
					doPack = Array(False, "加载文件数据出错!")
				End If
			Set SetMatch = Nothing
		Else
			doPack = Array(False, "加载文件出错!")
		End If
	End Function
	
End Class

'=====================base64 encode/decode==============
class packbase64
  Private objXmlDom
  Private objXmlNode
  
  Private Sub Class_Initialize()
      Set objXmlDom = Server.CreateObject(getXMLDOM)
  end sub
  
  Private Sub Class_Terminate()
	  Set objXmlDom =nothing
  end sub

  public function encode(AnsiCode)
    encode=""
	Set objXmlNode = objXmlDom.createElement("file")
	objXmlNode.datatype = "bin.base64"
    objXmlNode.nodeTypedvalue = AnsiCode
    encode = objXmlNode.text
	Set objXmlNode =nothing
  end function
  
  public function decode(base64Code)
    decode=""
	Set objXmlNode = objXmlDom.createElement("file")
	objXmlNode.datatype = "bin.base64"
	objXmlNode.text = base64Code
    decode = objXmlNode.nodeTypedvalue
 	Set objXmlNode =nothing
  end function
end class

Function getXMLDOM
	On Error Resume Next
	Dim Temp
	getXMLDOM="Microsoft.XMLDOM"
	Err = 0
	Dim TmpObj
	Set TmpObj = Server.CreateObject(getXMLDOM)
		Temp = Err
		IF Temp = 1 OR Temp = -2147221005 Then
			getXMLDOM="Msxml2.DOMDocument.5.0"
		End IF
		Err.Clear
	Set TmpObj = Nothing
	Err = 0
end function
' -------------------------------------------------------------
' 	使用方法
' -------------------------------------------------------------
' include.asp
' pjblog.model/cls_stream.asp
' pjblog.model/cls_fso.asp
' pjblog.model/cls_xml.asp
' pjblog.model/cls_pack.asp
'Dim cStream, c, cxml, a, b, d
'dim filepath, filecontent, pack, tt
'Set cStream = New Stream
'	c = cStream.SaveToLocal("http://www.evio.name/eBak/123.rar", "ee/1.pbd")
'Set cStream = Nothing
'if c(0) = 0 Then
'	set pack = New Sys_Pack
'		pack.NeedOut = True
'		pack.parentPath = "ee/" '生成在哪个文件夹中
'		pack.xmlPath = "ee/1.pbd" ' 被解压对象
'		pack.open	' 打开对象
'		pack.UnPack	' 安装
'	set pack = nothing
'	Response.Write("OK")
'End If
'
'Dim fso, pack, c
'set pack = New Sys_Pack
'c = pack.doPack("pj", "123.pbd")
'Response.Write("OK")
'Set pack = nothing
%>