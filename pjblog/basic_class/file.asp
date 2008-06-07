<%
'**文件处理类**********
' by puterjam
'*********************

Class IO
	Private stream
	Private openType
	
	Private Sub Class_Initialize()
		set stream = Server.CreateObject("ADODB.Stream")
		setEncode "utf-8"
		openType = 2
	end sub
	
	Private Sub Class_Terminate()
		set stream = nothing
	end sub
	
	Public Sub setEncode(ByVal encoding)
		stream.Charset = encoding
	end Sub
	   
	Public Function loadFile(ByVal File)
		loadFile = ""

		With stream
			.Type = openType
			.Mode = 3
			.Open
			.Position = stream.Size
        	.LoadFromFile Server.MapPath(File)
			loadFile = .ReadText
			.Close
		End With
	End Function
	
	Public Function saveFile(ByVal strBody,ByVal File)
		saveFile = false
		With stream
			.Type = openType
			.Open
			.Position = stream.Size
			.WriteText = strBody
			.SaveToFile Server.MapPath(File), 2
			saveFile = true
			.Close
		End With
	End Function
end Class
%>
