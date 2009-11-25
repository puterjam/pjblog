<!--#include file = "../../include.asp" -->
<!--#include file = "../../pjblog.model/cls_template.asp" -->
<!--#include file = "../../pjblog.data/cls_modset.asp" -->
<!--#include file = "../../pjblog.data/cls_plus.asp" -->
<!--#include file = "../../pjblog.data/cls_webconfig.asp" -->
<%
	Public Sub ouHtml
		Dim Mud, page
		Set Mud = New template
		Mud.Path = "template"
		Mud.FileName = "guestbook.html"
		Mud.TemplateContent = ""
		Mud.open
		Call web.General(Mud)
		Mud.Sets("KeyWords") = blog_KeyWords
		Mud.Sets("Description") = blog_Description
		Mud.Sets("GetCode") = Asp.GetCode("../pjblog.common/GetCode.asp")
		Mud.PageUrl = "default.asp?page={$page}"
		page = Trim(Asp.CheckStr(Request.QueryString("page")))
		Mud.CurrentPage = Asp.CheckPage(page)
		Mud.Buffer
		Mud.Flush
	End Sub
	Call ouHtml
	' ---------------------------------------------------
	'	留言本数据源
	' ---------------------------------------------------
	Public Function GuestBook
		Dim Rs, Arr
		Set Rs = Conn.Execute("Select book_ID, book_Messager, book_QQ, book_Url, book_Mail, book_face, book_IP, book_Content, book_PostTime, book_reply, book_HiddenReply, book_replyTime, book_replyAuthor From blog_book Order By book_PostTime")
		'								0			1			2			3		4			5			6		7	
'	8				9			10					11				12		
		If Rs.Bof Or Rs.Eof Then
			ReDim Arr(0, 0)
		Else
			Arr = Rs.GetRows
		End If
		Set Rs = Nothing
		GuestBook = Arr
	End Function
Sys.Close
%>