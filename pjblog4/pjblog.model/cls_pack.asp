<%
Class Sys_Pack

	Private xDom, tStream, fileList, baseM, upl, fso
	Public fileTotal, xmlPath, parentPath, NeedOut
	
	Private Sub Class_Initialize()
		NeedOut = True
	   	On Error Resume Next
	   	fileTotal=0
	   	Set baseM = New packbase64
		Set fso = New cls_fso
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
'	set pack = New Sys_Pack
'		pack.NeedOut = False
'		pack.parentPath = "ee/" '生成在哪个文件夹中
'		pack.xmlPath = "ee/1.pbd" ' 被解压对象
'		pack.open	' 打开对象
'		pack.UnPack	' 安装
'	set pack = nothing
%>