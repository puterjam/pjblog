<!--#include file = "../include.asp" -->
<!--#include file = "../pjblog.model/cls_Stream.asp" -->
<!--#include file = "../pjblog.common/md5.asp" -->
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
		If Not stat_EditAll Then Response.Write("try{$('.index_edit').remove();}catch(e){}")
		If Not stat_DelAll Then Response.Write("try{$('.index_del').remove();}catch(e){}")
		cookie = Request.Cookies(Sys.CookieName & "_IndexArticleList")
		Set Rs = Conn.Execute("Select log_ID, log_CommNums, log_ViewNums, log_QuoteNums From blog_Content Where log_ID in (" & cookie & ")")
		If Not (Rs.Bof And Rs.Eof) Then
			Do While Not Rs.Eof
				Response.Write("try{$('#indexComment_" & Rs(0).value & "').text('" & Rs(1).value & "')}catch(e){}")
				Response.Write("try{$('#indexQuote_" & Rs(0).value & "').text('" & Rs(3).value & "')}catch(e){}")
				Response.Write("try{$('#indexView_" & Rs(0).value & "').text('" & Rs(2).value & "')}catch(e){}")
			Rs.MoveNext
			Loop
		End If
		Set Rs = Nothing
		Call FillSideBar
		Response.Write("}" & vbcrlf & "$(document).ready(function(){load();})")
	End Sub
	
	Private Sub Article
		Dim Rs, ID, doo, local, Str, Str2, Str3, ReplyStr, TempReply
		Dim GetRow, i, RowLeft, RowRight, PageSize, PageLen, CurrtPage, PageStr
		ID = Asp.CheckStr(Request.QueryString("id"))
		If Not Asp.IsInteger(ID) Then Exit Sub
		If Not Asp.IsInteger(Request.QueryString("page")) Then
			Sys.doGet("Update blog_Content Set log_ViewNums = log_ViewNums + 1 Where log_ID=" & ID)
		End If
		If Not stat_EditAll Then Response.Write("try{$('.index_edit').remove();}catch(e){}")
		If Not stat_DelAll Then Response.Write("try{$('.index_del').remove();}catch(e){}")
		Set Rs = Conn.Execute("Select log_ID, log_CommNums, log_ViewNums, log_QuoteNums From blog_Content Where log_ID=" & ID)
		If Not (Rs.Bof And Rs.Eof) Then
			Do While Not Rs.Eof
				Response.Write("try{$('#indexComment').text('" & Rs(1).value & "')}catch(e){}")
				Response.Write("try{$('#indexQuote').text('" & Rs(3).value & "')}catch(e){}")
				Response.Write("try{$('#indexView').text('" & Rs(2).value & "')}catch(e){}")
			Rs.MoveNext
			Loop
		End If
		Set Rs = Nothing
		If Len(memName) > 0 Then Response.Write("try{if ($('#passArea')){$('#comm_Author').val('" & memName & "');$('#comm_Author').attr('readOnly', true);$('#passArea').remove();$('#removevalidate').remove();}}catch(e){};")
		' ---------------------------------------------------
		' 	加载评论 - 开始
		' ---------------------------------------------------
		local = "../pjblog.template/" & blog_DefaultSkin & "/"
		Set Rs = Conn.Execute("Select comm_ID, comm_Author, comm_Content, comm_DisSM, comm_DisUBB, comm_DisIMG, comm_AutoURL, comm_AutoKEY, comm_Email, comm_WebSite, comm_PostIP, comm_PostTime, comm_IsAudit, replay_content, reply_author, reply_time From blog_Comment Where blog_ID=" & ID & " Order By comm_PostTime Desc")
		'								0			1			2				3			4			5			6				
'	7				8			9			10			11				12				13				14			15
		If Not (Rs.Bof And Rs.Eof) Then
			Set doo = New Stream
				Str = doo.LoadFile(local & "l_comment.html")
				ReplyStr = doo.LoadFile(local & "l_replylayout.html")
			Set doo = Nothing
			GetRow = Rs.GetRows	
		Else
			Redim GetRow(0, 0)
		End If
		Set Rs = Nothing
		If UBound(GetRow, 1) > 0 Then
			PageStr = "javascript:ceeevio.Comment.page({$page}, false, '../pjblog.logic/code.asp?action=article&id=" & ID & "');"
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
				
				If blog_commAduit Then
					If Asp.BlankString(GetRow(1, i)) = memName Or GetRow(12, i) Or stat_Admin Then
						Str2 = Replace(Str2, "<#comm_Content#>", Asp.BlankString(UBBCode(GetRow(2, i), GetRow(3, i), GetRow(4, i), GetRow(5, i), GetRow(6, i), GetRow(7, i), True)), 1, -1, 1)
					Else
						Str2 = Replace(Str2, "<#comm_Content#>", "评论正在审核中...", 1, -1, 1)
					End If
				Else
					Str2 = Replace(Str2, "<#comm_Content#>", Asp.BlankString(UBBCode(GetRow(2, i), GetRow(3, i), GetRow(4, i), GetRow(5, i), GetRow(6, i), GetRow(7, i), True)), 1, -1, 1)
				End If

				Str2 = Replace(Str2, "<#comm_mail#>", Asp.BlankString(GetRow(8, i)), 1, -1, 1)
				Str2 = Replace(Str2, "<#comm_weburl#>", Asp.BlankString(GetRow(9, i)), 1, -1, 1)
				Str2 = Replace(Str2, "<#comm_ip#>", Asp.BlankString(GetRow(10, i)), 1, -1, 1)
				Str2 = Replace(Str2, "<#comm_posttime#>", Asp.BlankString(Asp.DateToStr(GetRow(11, i), "Y-m-d H:I:S")), 1, -1, 1)
				Dim Gra
				Set Gra = New Gravatar
					Gra.Gravatar_r = "r"
					Gra.Gravatar_EmailMd5 = MD5(GetRow(8, i))
					Str2 = Replace(Str2, "<#GRA#>", Gra.outPut, 1, -1, 1)
				Set Gra = Nothing
				
				If stat_CommentDel Then
					Str2 = Replace(Str2, "<#comm_del#>", "<a href=""javascript:ceeevio.Comment.del(" & GetRow(0, i) & ");"" onclick=""return confirm('确定删除?\n删除后无法恢复')"">删除</a>", 1, -1, 1)				
				Else
					Str2 = Replace(Str2, "<#comm_del#>", "", 1, -1, 1)
				End If
				
				If stat_Admin Then
					Str2 = Replace(Str2, "<#comm_reply#>", "<a href=""javascript:void(0);"" onclick=""ceeevio.Comment.replyBox(" & GetRow(0, i) & ")"">回复</a>", 1, -1, 1)
				Else
					Str2 = Replace(Str2, "<#comm_reply#>", "", 1, -1, 1)
				End If
				' -----------------------------------------------------
				'	回复样式 开始
				' -----------------------------------------------------
				If IsNull(GetRow(13, i)) Or IsEmpty(GetRow(13, i)) Or Len(GetRow(13, i)) = 0 Then
					Str2 = Replace(Str2, "<#comm_replycontent#>", "", 1, -1, 1)
				Else
					If blog_commAduit And (Not GetRow(12, i)) Then
						Str2 = Replace(Str2, "<#comm_replycontent#>", "", 1, -1, 1)
					Else
						TempReply = ReplyStr
						TempReply = Replace(TempReply, "<#replyAuthor#>", Asp.BlankString(GetRow(14, i)), 1, -1, 1)
						TempReply = Replace(TempReply, "<#replyTime#>", Asp.BlankString(Asp.DateToStr(GetRow(15, i), "Y-m-d H:I:S")), 1, -1, 1)
						TempReply = Replace(TempReply, "<#replyContent#>", Asp.BlankString(GetRow(13, i)), 1, -1, 1)
						If GetRow(13, i) = null Or Len(GetRow(13, i)) = 0 Or GetRow(13, i) = "" Or GetRow(13, i) = Empty Then
							Str2 = Replace(Str2, "<#comm_replycontent#>", "", 1, -1, 1)
						Else
							Str2 = Replace(Str2, "<#comm_replycontent#>", Asp.BlankString(TempReply), 1, -1, 1)
						End If
					End If
				End If
				' -----------------------------------------------------
				'	回复样式 结束
				' -----------------------------------------------------
				
				If blog_commAduit And stat_Admin Then
					If GetRow(12, i) Then
						Str2 = Replace(Str2, "<#comm_Aduit#>", "[ <a href=""javascript:void(0);"" onclick=""ceeevio.Comment.Aduit(" & Asp.BlankString(GetRow(0, i)) & ", 0, this)"">取消审核</a> ]")
					Else
						Str2 = Replace(Str2, "<#comm_Aduit#>", "[ <a href=""javascript:void(0);"" onclick=""ceeevio.Comment.Aduit(" & Asp.BlankString(GetRow(0, i)) & ", 1, this)"">通过审核</a> ]")
					End If
				Else
					Str2 = Replace(Str2, "<#comm_Aduit#>", "")
				End If
				Str2 = "<div id=""comment_" & Asp.BlankString(GetRow(0, i)) & """ class=""CommPart"">" & Str2 & "</div>"
				Str3 = Str3 & Str2
			Next
			Str3 = Str3 & MultiPage(PageLen + 1, blogcommpage, CurrtPage, PageStr, "float:right", "", False) & "<div class=""clear""></div>"
			Response.Write("try{$('#commentBox').html('" & outputSideBar(Str3) & "');}catch(e){}")
		End If
		' ---------------------------------------------------
		' 	加载评论 - 结束
		'	加载评论框 - 开始
		' ---------------------------------------------------
		'Set doo = New Stream
			'Str = doo.LoadFile(local & "l_commBox.html")
		'Set doo = Nothing
		'Response.Write("try{$('#CommPostBox').html('" & outputSideBar(Str) & "')}catch(e){}")
		' ---------------------------------------------------
		' 	加载评论框 - 结束
		'	加载侧边 - 开始
		' ---------------------------------------------------
		Call FillSideBar
		' ---------------------------------------------------
		' 	加载评论 - 结束
		' ---------------------------------------------------
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
		Response.Write("for (var a = 0 ; a < PluginOutPutString.length ; a++){try{$('#' + PluginOutPutString[a][0]).html(outputSideBar(PluginOutPutString[a][1]));}catch(e){}}" & vbcrlf)
		Response.Write("}" & vbcrlf)
'		Response.Write("alert(jQuery(PluginOutPutString).length);jQuery(PluginOutPutString).each(function(i){$('#' + PluginOutPutString[i][0]).html(outputSideBar(PluginOutPutString[i][1]))})}")
		Response.Write("if (PluginOutPutString.length > 0){doSide();}else{window.location.reload();}")
	End Sub

End Class
%>