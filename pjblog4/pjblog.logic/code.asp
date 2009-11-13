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
	End Sub

End Class
%>