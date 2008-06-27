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
        errorcode = -1
    End Sub

    Private Sub Class_Terminate()

    End Sub

    '------------------------------------------------
    '函数名字：Open()
    'Open=0，XMLDocument就是一个成功装载XML文档的对象了。
    '------------------------------------------------

    Public Function Open()
        Dim strSourceFile, strError, xmlDom
        xmlDom = getXMLDOM()
        If xmlDom = False Then
            errorcode = -18239123
            Exit Function
        End If

        Set XMLDocument = Server.CreateObject(xmlDom)
        XMLDocument.async = False
        strSourceFile = Server.MapPath(XmlPath)
        XMLDocument.load(strSourceFile)
        errorcode = XMLDocument.parseerror.errorcode
    End Function

    '------------------------------------------------
    '函数名字：OpenXML()
    'Open=0，XMLDocument就是一个成功装载XML文档的对象了。
    '------------------------------------------------

    Public Function OpenXML(xmlStr)
        Dim strSourceFile, strError, xmlDom
        xmlDom = getXMLDOM()
        If xmlDom = False Then
            errorcode = -18239123
            Exit Function
        End If

        Set XMLDocument = Server.CreateObject(getXMLDOM())
        XMLDocument.async = False
        XMLDocument.load(xmlStr)
        errorcode = XMLDocument.parseerror.errorcode
    End Function

    '------------------------------------------------
    '函数名字：getError()
    '------------------------------------------------

    Public Function getError()
        getError = errorcode
    End Function

    '------------------------------------------------
    '函数名字：CloseXml()
    '------------------------------------------------

    Public Function CloseXml()
        If IsObject(XMLDocument) Then
            Set XMLDocument = Nothing
        End If
    End Function


    '------------------------------------------------
    'SelectXmlNodeText(elementname)
    '获得当个 elementname 元素
    '------------------------------------------------

    Public Function SelectXmlNodeText(elementname)
        Dim xmlItems
        selectXmlNodeText = ""
        Set xmlItems = XMLDocument.getElementsByTagName(elementname)
        If xmlItems.Length <> 0 Then selectXmlNodeText = xmlItems.Item(0).text
    End Function

    '------------------------------------------------
    'SelectXmlNode(elementname,itemID)
    '获得当个 elementname 元素
    '------------------------------------------------

    Public Function SelectXmlNode(elementname, itemID)
        Set SelectXmlNode = XMLDocument.getElementsByTagName(elementname).Item(itemID)
    End Function

    '------------------------------------------------
    'GetXmlNodeLength(elementname)
    '获得当个 elementname 元素的Length值
    '------------------------------------------------

    Public Function GetXmlNodeLength(elementname)
        GetXmlNodeLength = XMLDocument.getElementsByTagName(elementname).Length
    End Function

    '------------------------------------------------
    'GetAttributes(elementname,nodeName,ID)
    '获得当个 elementname 元素的attributes值
    '------------------------------------------------

    Public Function GetAttributes(elementname, nodeName, itemID)
        Dim XmlAttributes, i
        Set XmlAttributes = XMLDocument.getElementsByTagName(elementname).Item(itemID).Attributes
        For i = 0 To XmlAttributes.Length -1
            If XmlAttributes(i).Name = nodeName Then
                GetAttributes = XmlAttributes(i).Value
                Exit Function
            End If
        Next
        GetAttributes = 0
    End Function

    '------------------------------------------------
    'SelectXmlNodeItemText(elementname,ID)
    '获得当个某 elementname 元素的Length值
    '------------------------------------------------

    Public Function SelectXmlNodeItemText(elementname, ID)
        Dim xmlItems
        SelectXmlNodeItemText = ""
        Set xmlItems = XMLDocument.getElementsByTagName(elementname)
        If xmlItems.Length <> 0 Then SelectXmlNodeItemText = xmlItems.Item(ID).text
    End Function

    '------------------------------------------------
    'WriteXmlNodeItemText(elementname,ID)
    '写入当个某 elementname 元素的text值
    '------------------------------------------------

    Public Function WriteXmlNodeItemText(elementname, ID, Str)
        WriteXmlNodeItemText = 0
        Dim temp, temp1
        Set temp = XMLDocument.getElementsByTagName(elementname).Item(ID)
        temp.childNodes(0).text = Str
        XMLDocument.save Server.MapPath(XmlPath)
    End Function

    '------------------------------------------------
    'IsXmlNode(elementname)
    '检测是否存在 elementname 元素
    'True代表存在,False代表不存在
    '------------------------------------------------

    Public Function IsXmlNode(elementname)
        Dim Temp
        IsXmlNode = True
        Set Temp = XMLDocument.getElementsByTagName(elementname)

        If Temp.Length = 0 Then
            IsXmlNode = False
        End If
    End Function

End Class
%>
