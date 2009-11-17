<!--#include file = "../include.asp" -->
<%
Dim Code
Set Code = New ChkCode
Set Code = Nothing
Class ChkCode

	Public Property Get Action
		Action = Request.QueryString("action")
	End Property

	' ***********************************************
	'	类初始化
	' ***********************************************
	Private Sub Class_Initialize
		Select Case Lcase(Action)
			Case "default" Call default
			Case "article" Call Article
			Case "category" Call category
		End Select
    End Sub 
     
	' ***********************************************
	'	类终结化
	' ***********************************************
    Private Sub Class_Terminate
		Sys.Close
    End Sub
	
	Private Sub default
		Dim Rs, cookie
		Response.Write("function load(){")
		If Not stat_EditAll Then Response.Write("try{JsCopy.index.edit();}catch(e){}")
		If Not stat_DelAll Then Response.Write("try{JsCopy.index.del();}catch(e){}")
		cookie = Request.Cookies(Sys.CookieName & "_IndexArticleList")
		Set Rs = Conn.Execute("Select log_ID, log_CommNums, log_ViewNums, log_QuoteNums From blog_Content Where log_ID in (" & cookie & ")")
		If Not (Rs.Bof And Rs.Eof) Then
			Do While Not Rs.Eof
				Response.Write("try{$('indexComment_" & Rs(0).value & "').innerHTML = '" & Rs(1).value & "'}catch(e){}")
				Response.Write("try{$('indexQuote_" & Rs(0).value & "').innerHTML = '" & Rs(3).value & "'}catch(e){}")
				Response.Write("try{$('indexView_" & Rs(0).value & "').innerHTML = '" & Rs(2).value & "'}catch(e){}")
			Rs.MoveNext
			Loop
		End If
		Set Rs = Nothing
		Call FillSideBar
		Response.Write("}" & vbcrlf & "window.onload = function(){load();}")
	End Sub
	
	Private Sub Article
		Dim Rs, cookie
		If Not stat_EditAll Then Response.Write("try{JsCopy.index.edit();}catch(e){}")
		If Not stat_DelAll Then Response.Write("try{JsCopy.index.del();}catch(e){}")
		cookie = Request.Cookies(Sys.CookieName & "_IndexArticleList")
		Set Rs = Conn.Execute("Select log_ID, log_CommNums, log_ViewNums, log_QuoteNums From blog_Content Where log_ID in (" & cookie & ")")
		If Not (Rs.Bof And Rs.Eof) Then
			Do While Not Rs.Eof
				Response.Write("try{$('indexComment_" & Rs(0).value & "').innerHTML = '" & Rs(1).value & "'}catch(e){}")
				Response.Write("try{$('indexQuote_" & Rs(0).value & "').innerHTML = '" & Rs(3).value & "'}catch(e){}")
				Response.Write("try{$('indexView_" & Rs(0).value & "').innerHTML = '" & Rs(2).value & "'}catch(e){}")
			Rs.MoveNext
			Loop
		End If
		Set Rs = Nothing
		Call FillSideBar
	End Sub
	
	Private Sub category
		Dim Rs, cookie
		Response.Write("function load(){")
		If Not stat_EditAll Then Response.Write("try{JsCopy.index.edit();}catch(e){}")
		If Not stat_DelAll Then Response.Write("try{JsCopy.index.del();}catch(e){}")
		cookie = Request.Cookies(Sys.CookieName & "_categoryArticleList")
		Set Rs = Conn.Execute("Select log_ID, log_CommNums, log_ViewNums, log_QuoteNums From blog_Content Where log_ID in (" & cookie & ")")
		If Not (Rs.Bof And Rs.Eof) Then
			Do While Not Rs.Eof
				Response.Write("try{$('indexComment_" & Rs(0).value & "').innerHTML = '" & Rs(1).value & "'}catch(e){}")
				Response.Write("try{$('indexQuote_" & Rs(0).value & "').innerHTML = '" & Rs(3).value & "'}catch(e){}")
				Response.Write("try{$('indexView_" & Rs(0).value & "').innerHTML = '" & Rs(2).value & "'}catch(e){}")
			Rs.MoveNext
			Loop
		End If
		Set Rs = Nothing
		Call FillSideBar
		Response.Write("}" & vbcrlf & "window.onload = function(){load();}")
	End Sub
	
	Private Sub FillSideBar
		Response.Write("if (PluginOutPutString.length > 0){for (var a = 0 ; a < PluginOutPutString.length ; a++){try{$(PluginOutPutString[a][0]).innerHTML = outputSideBar(PluginOutPutString[a][1])}catch(e){}}}else{alert(PluginOutPutString.length)}")
	End Sub

End Class
%>