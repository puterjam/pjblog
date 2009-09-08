<%
class PJSetup
	Private xDom
	Private tStream
	Private fileList
	Private baseM
	Private upl
	Public fileTotal
	Public xmlPath
	
	Private Sub Class_Initialize()
	   On Error Resume Next
	   fileTotal=0
	   set baseM=new base64
	   set tStream = Server.CreateObject("adodb.stream")
	   If Err<>0 Then
	     Err.clear
		 response.write Err.Description
		 response.end
	   end if
	   tStream.Type = 1
	   tStream.Mode = 3
	end sub
  
	Private Sub Class_Terminate()
	   'DeleteFiles(Server.MapPath(xmlPath))
	   if IsObject(xDom) then set xDom=nothing
	   if IsObject(tStream) then set tStream=nothing
	   if IsObject(baseM) then set baseM=nothing
	   if IsObject(upl) then Set upl=nothing
	end sub
      
	Public Sub open()
	    dim i
		Set xDom = Server.CreateObject(getXMLDOM)
		xDom.load(Server.MapPath(xmlPath))
		fileTotal=xDom.getElementsByTagName("files/f").length
	end sub
	
	Public Function element_fn(ByVal i)
		element_fn = Trim(xDom.getElementsByTagName("files/f/fn").item(i).text)
	End Function
	
	Public Function element_fb(ByVal i)
		element_fb = Trim(xDom.getElementsByTagName("files/f/fb").item(i).text)
	End Function
end class

'=====================base64 encode/decode==============
class base64
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
%>