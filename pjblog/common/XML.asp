<% 
'=================================================
' XML Class for PJBlog2
' Author: PuterJam
' UpdateDate: 2006-1-19
'=================================================

Class PXML
  Public XmlPath
  Private errorcode
  Public XMLDocument

  Private Sub Class_Initialize()
   errorcode=-1
  end sub
  
  Private Sub Class_Terminate()

  end sub
  
  '------------------------------------------------ 
  '函数名字：Open() 
  'Open=0，XMLDocument就是一个成功装载XML文档的对象了。 
  '------------------------------------------------ 
  Public function Open()
     dim strSourceFile,strError,xmlDom
     xmlDom = getXMLDOM()
     if xmlDom = false then 
     	errorcode=-18239123
	    exit function
     end if
     
     Set XMLDocument = Server.CreateObject(xmlDom)
     XMLDocument.async = false  
     strSourceFile = Server.MapPath(XmlPath) 
     XMLDocument.load(strSourceFile) 
     errorcode=XMLDocument.parseerror.errorcode
  end function 

  '------------------------------------------------ 
  '函数名字：OpenXML() 
  'Open=0，XMLDocument就是一个成功装载XML文档的对象了。 
  '------------------------------------------------ 
  Public function OpenXML(xmlStr)
     dim strSourceFile,strError,xmlDom
     xmlDom = getXMLDOM()
     if xmlDom = false then 
     	errorcode=-18239123
	    exit function
     end if
     
     Set XMLDocument = Server.CreateObject(getXMLDOM())
     XMLDocument.async = false
     XMLDocument.load(xmlStr)
     errorcode=XMLDocument.parseerror.errorcode
  end function 

  '------------------------------------------------ 
  '函数名字：getError() 
  '------------------------------------------------ 
  Public function getError()
   getError=errorcode
  end function 
  
  '------------------------------------------------ 
  '函数名字：CloseXml() 
  '------------------------------------------------ 
  Public function CloseXml() 
   if IsObject(XMLDocument) then 
   set XMLDocument=nothing 
   end if 
  end function 
  
  
  '------------------------------------------------ 
  'SelectXmlNodeText(elementname) 
  '获得当个 elementname 元素
  '------------------------------------------------ 
  Public function SelectXmlNodeText(elementname)  
      dim xmlItems
      selectXmlNodeText = ""
      set xmlItems = XMLDocument.getElementsByTagName(elementname)
      if xmlItems.length <> 0 then selectXmlNodeText = xmlItems.item(0).text
  end function
  
  '------------------------------------------------ 
  'SelectXmlNode(elementname,itemID) 
  '获得当个 elementname 元素
  '------------------------------------------------ 
  Public function SelectXmlNode(elementname,itemID) 
      set SelectXmlNode = XMLDocument.getElementsByTagName(elementname).item(itemID)
  end function
  
    '------------------------------------------------ 
  'GetXmlNodeLength(elementname) 
  '获得当个 elementname 元素的Length值 
  '------------------------------------------------ 
  Public function GetXmlNodeLength(elementname)  
      GetXmlNodeLength = XMLDocument.getElementsByTagName(elementname).length
  end function
  
  '------------------------------------------------ 
  'GetAttributes(elementname,nodeName,ID) 
  '获得当个 elementname 元素的attributes值 
  '------------------------------------------------ 
  Public function GetAttributes(elementname,nodeName,itemID)  
      dim XmlAttributes,i
      set XmlAttributes=XMLDocument.getElementsByTagName(elementname).item(itemID).attributes
      for i=0 to XmlAttributes.length-1
       if XmlAttributes(i).name=nodeName then 
        GetAttributes=XmlAttributes(i).value
        exit function
       end if
      next
      GetAttributes = 0
  end function  
  
  '------------------------------------------------ 
  'SelectXmlNodeItemText(elementname,ID) 
  '获得当个某 elementname 元素的Length值 
  '------------------------------------------------ 
  Public function SelectXmlNodeItemText(elementname,ID) 
      dim xmlItems
      SelectXmlNodeItemText = ""
      set xmlItems = XMLDocument.getElementsByTagName(elementname)
      if xmlItems.length <> 0 then SelectXmlNodeItemText = xmlItems.item(ID).text
  end function

  '------------------------------------------------ 
  'WriteXmlNodeItemText(elementname,ID) 
  '写入当个某 elementname 元素的text值 
  '------------------------------------------------ 
  Public function WriteXmlNodeItemText(elementname,ID,str) 
      WriteXmlNodeItemText=0
      dim temp,temp1
      set temp=XMLDocument.getElementsByTagName(elementname).item(ID)
      temp.childNodes(0).text=str
	  XMLDocument.save Server.MapPath(XmlPath)
  end function

  '------------------------------------------------ 
  'IsXmlNode(elementname) 
  '检测是否存在 elementname 元素 
  'True代表存在,False代表不存在 
  '------------------------------------------------ 
  Public function IsXmlNode(elementname)
   dim Temp
   IsXmlNode=true
   set Temp=XMLDocument.getElementsByTagName(elementname)
   
   if Temp.length = 0 then
    IsXmlNode=false
   end if
  end function
end Class
%>