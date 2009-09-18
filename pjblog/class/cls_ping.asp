<%
Class ping
	
	Private Pname, Purl, Rs
	
	' ********************************************
	'	设置属性
	' ********************************************
	Public property LET Ping_Name(ByVal values)
		Pname = values
	End property
	
	Public property LET Ping_Url(ByVal values)
		Purl = values
	End property
	
	' ********************************************
	'	类初始化
	' ********************************************
	Private Sub Class_Initialize
        
    End Sub 
    
	' ********************************************
	'	类终结
	' ********************************************
    Private Sub Class_Terminate
        
    End Sub
	
	' ********************************************
	'	获得Ping地址串
	' ********************************************
	Private Function GetPingContent
		Dim PingArray, PingCount, i
		PingArray = Array()
		Set Rs = Server.CreateObject("Adodb.RecordSet")
			Rs.open "Select * From blog_Ping", Conn, 1, 3
			PingCount = Rs.RecordCount
			If Int(PingCount) > 0 Then
				ReDim Preserve PingArray(Int(PingCount) - 1)
				i = 0
				Do While Not Rs.Eof
					PingArray(i) = Trim(Rs("Ping_url").value)
					i = i + 1
				Rs.MoveNext
				Loop
			End If
			Rs.Close
		Set Rs = Nothing
		GetPingContent = PingArray
	End Function
	
	' ********************************************
	'	将Ping地址插入数据库
	' ********************************************
	Public Sub insPingBase
		Set Rs = Server.CreateObject("Adodb.RecordSet")
			Rs.open "blog_Ping", Conn, 1, 2
			Rs.AddNew
				Rs("Ping_Name") = Trim(Pname)
				Rs("Ping_url") = Trim(Purl)
			Rs.Update
			Rs.Close
		Set Rs = Nothing
	End Sub
	
	' ********************************************
	'	将Ping地址更新到数据库
	' ********************************************
	Public Sub updatePingBase(ByVal id)
		If Not IsInteger(id) Then Exit Sub
		Set Rs = Server.CreateObject("Adodb.RecordSet")
			Rs.open "Select * From blog_Ping Where Ping_ID=" & id, Conn, 1, 3
				Rs("Ping_Name") = Trim(Pname)
				Rs("Ping_url") = Trim(Purl)
			Rs.Update
			Rs.Close
		Set Rs = Nothing
	End Sub
	
	' ********************************************
	'	将Ping地址从数据库删除
	' ********************************************
	Public Sub DelPingBase(ByVal id)
		If Not IsInteger(id) Then Exit Sub
		Conn.Execute("Delete * From blog_Ping Where Ping_ID=" & id)
	End Sub
	
	' ********************************************
	'	单个发送Ping主函数
	' ********************************************
	Private Sub SendPingContent(ByVal CpingUrl, ByVal ArticleUrl)
		On Error Resume Next
		Dim xmlStr, xmlSite
		xmlSite = siteURL
		xmlSite = Replace(xmlSite, "\", "/")
		If Right(xmlSite, 1) = "/" Then xmlSite = Mid(xmlSite, 1, (Len(xmlSite) - 1))
		xmlSite = xmlSite & "/" & ArticleUrl
		xmlStr = "<?xml version=""1.0""?><methodCall><methodName>weblogUpdates.ping</methodName><params><param><value>"&SiteName&"</value></param><param><value>"&xmlSite&"</value></param></params></methodCall>"
		Dim objPing
      	Set objPing = Server.CreateObject("MSXML2.ServerXMLHTTP")
			objPing.SetTimeOuts 10000, 10000, 10000, 10000 
			' **************************************
			'@ 第一个数值：解析DNS名字的超时时间10秒 
			'@ 第二个数值：建立Winsock连接的超时时间10秒 
			'@ 第三个数值：发送数据的超时时间10秒 
			'@ 第四个数值：接收Response的超时时间10
			' **************************************
			objPing.open "POST", CpingUrl, False
			objPing.setRequestHeader "Content-Type", "text/xml"
			objPing.send xmlStr
      	Set objPing = Nothing
      	Err.Clear
	End Sub
	
	' ********************************************
	'	批量发送Ping主函数
	' ********************************************
	Public Sub WebPing(ByVal LogUrl)
		Dim WebArray, i, xmlSite
		WebArray = GetPingContent
		xmlSite = siteURL
		xmlSite = Replace(xmlSite, "\", "/")
		If Right(xmlSite, 1) = "/" Then xmlSite = Mid(xmlSite, 1, (Len(xmlSite) - 1))
		LogUrl = xmlSite & "/" & LogUrl
		For i = 0 To UBound(WebArray)
			SendPingContent WebArray(i), LogUrl
		Next
	End Sub
	
End Class
%>