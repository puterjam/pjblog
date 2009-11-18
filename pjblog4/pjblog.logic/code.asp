<!--#include file = "../include.asp" -->
<!--#include file = "../pjblog.model/cls_Stream.asp" -->
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
		Dim Rs, ID, doo, local, Str, Str2, Str3
		Dim GetRow, i, RowLeft, RowRight, PageSize, PageLen, CurrtPage, PageStr
		ID = Asp.CheckStr(Request.QueryString("id"))
		If Not Asp.IsInteger(ID) Then Exit Sub
		Sys.doGet("Update blog_Content Set log_ViewNums = log_ViewNums + 1 Where log_ID=" & ID)
		If Not stat_EditAll Then Response.Write("try{JsCopy.index.edit();}catch(e){}")
		If Not stat_DelAll Then Response.Write("try{JsCopy.index.del();}catch(e){}")
		Set Rs = Conn.Execute("Select log_ID, log_CommNums, log_ViewNums, log_QuoteNums From blog_Content Where log_ID=" & ID)
		If Not (Rs.Bof And Rs.Eof) Then
			Do While Not Rs.Eof
				Response.Write("try{$('indexComment').innerHTML = '" & Rs(1).value & "'}catch(e){}")
				Response.Write("try{$('indexQuote').innerHTML = '" & Rs(3).value & "'}catch(e){}")
				Response.Write("try{$('indexView').innerHTML = '" & Rs(2).value & "'}catch(e){}")
			Rs.MoveNext
			Loop
		End If
		Set Rs = Nothing
		If Len(memName) > 0 Then Response.Write("try{if ($('passArea')){$('comm_Author').value='" & memName & "';$('comm_Author').readOnly = true;$('passArea').parentNode.removeChild($('passArea'))}}catch(e){};")
		' ---------------------------------------------------
		' 	加载评论
		' ---------------------------------------------------
		local = "../pjblog.template/" & blog_DefaultSkin & "/"
		Set Rs = Conn.Execute("Select comm_ID, comm_Author, comm_Content, comm_DisSM, comm_DisUBB, comm_DisIMG, comm_AutoURL, comm_AutoKEY, comm_Email, comm_WebSite, comm_PostIP, comm_PostTime From blog_Comment Where blog_ID=" & ID)
		'								0			1			2				3			4			5			6				
'	7				8			9			10			11				
		If Not (Rs.Bof And Rs.Eof) Then
			Set doo = New Stream
				Str = doo.LoadFile(local & "l_comment.html")
			Set doo = Nothing
			GetRow = Rs.GetRows	
		Else
			Redim GetRow(0, 0)
		End If
		Set Rs = Nothing
		If UBound(GetRow, 1) > 0 Then
			PageStr = "javascript:CheckForm.comment.page({$page}, false, '../pjblog.logic/code.asp?action=article&id=" & ID & "');"
			CurrtPage = Request.QueryString("page")
			If Not Asp.Isinteger(CurrtPage) Then CurrtPage = 1
			If CurrtPage < 1 Then CurrtPage = 1
			PageSize = blogcommpage
			PageLen = UBound(GetRow, 2)
			RowLeft = (CurrtPage - 1) * PageSize
			RowRight = CurrtPage * PageSize - 1
			If RowLeft < 0 Then RowLeft = 0
			If RowRight > PageLen Then RowRight = PageLen
			Str3 = ""
			For i = RowLeft To RowRight
				Str2 = Str
				Str2 = Replace(Str2, "<#comm_id#>", Asp.BlankString(GetRow(0, i)), 1, -1, 1)
				Str2 = Replace(Str2, "<#comm_author#>", Asp.BlankString(GetRow(1, i)), 1, -1, 1)
				Str2 = Replace(Str2, "<#comm_Content#>", Asp.BlankString(UBBCode(GetRow(2, i), GetRow(3, i), GetRow(4, i), GetRow(5, i), GetRow(6, i), GetRow(7, i), True)), 1, -1, 1)
				Str2 = Replace(Str2, "<#comm_mail#>", Asp.BlankString(GetRow(8, i)), 1, -1, 1)
				Str2 = Replace(Str2, "<#comm_weburl#>", Asp.BlankString(GetRow(9, i)), 1, -1, 1)
				Str2 = Replace(Str2, "<#comm_ip#>", Asp.BlankString(GetRow(10, i)), 1, -1, 1)
				Str2 = Replace(Str2, "<#comm_posttime#>", Asp.BlankString(Asp.DateToStr(GetRow(11, i), "Y-m-d H:I:S")), 1, -1, 1)
				Str2 = "<div id=""comment_" & Asp.BlankString(GetRow(0, i)) & """ class=""CommPart"">" & Str2 & "</div>"
				Str3 = Str3 & Str2
			Next
			Str3 = Str3 & MultiPage(PageLen + 1, blogcommpage, CurrtPage, PageStr, "float:right", "", False)
			Response.Write("try{$('commentBox').innerHTML = '" & outputSideBar(Str3) & "';}catch(e){}")
		End If
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
		Response.Write("function doSide(){" & vbcrlf)
		Response.Write("for (var a = 0 ; a < PluginOutPutString.length ; a++){try{$(PluginOutPutString[a][0]).innerHTML = outputSideBar(PluginOutPutString[a][1])}catch(e){}}" & vbcrlf)
		Response.Write("}" & vbcrlf)
		Response.Write("if (PluginOutPutString.length > 0){doSide();}else{fillSide();doSide();}")
	End Sub

End Class
%>