<!--#include file = "../../include.asp" -->
<!--#include file = "../../pjblog.data/control/cls_category.asp" -->
<!--#include file = "../../pjblog.model/cls_fso.asp" -->
<!--#include file = "../../pjblog.express/log_control.asp" -->
<%
Dim doCategory
Set doCategory = New do_Category
Set doCategory = Nothing

Class do_Category

	Private cate_ID, cate_Order, cate_icon, cate_Name, cate_Intro, cate_Folder, cate_URL, cate_local, cate_Secret, cate_count, cate_OutLink, cate_Lock
	
	Private Str, i, ErrorJoin
	
	Public Property Get Action
		Action = Asp.CheckStr(Request.QueryString("action"))
	End Property
	
	' ***********************************************
	'	类初始化
	' ***********************************************
	Private Sub Class_Initialize
		Select Case Action
			Case "add" Call Add
			Case "edit" Call edit
			Case "del" Call del
			Case "clear" Call Clear
		End Select
    End Sub 
     
	' ***********************************************
	'	类终结化
	' ***********************************************
    Private Sub Class_Terminate
		Sys.Close
    End Sub
	
	Private Sub Add
		On Error Resume Next
		cate_Order = Int(Asp.CheckUrl(Asp.CheckStr(Request.Form("cate_Order"))))
		cate_icon = Trim(Asp.CheckUrl(Asp.CheckStr(Request.Form("cate_icon"))))
		cate_Name = Trim(Asp.CheckUrl(Asp.CheckStr(Request.Form("cate_Name"))))
		cate_Intro = Trim(Asp.CheckUrl(Asp.CheckStr(Request.Form("cate_Intro"))))
		cate_Folder = Trim(Asp.CheckUrl(Asp.CheckStr(Request.Form("cate_Folder"))))
		cate_URL = Trim(Asp.CheckUrl(Asp.CheckStr(Request.Form("cate_URL"))))
		cate_local = Int(Asp.CheckUrl(Asp.CheckStr(Request.Form("cate_local"))))
		cate_Secret = CBool(Asp.CheckUrl(Asp.CheckStr(Request.Form("cate_Secret"))))
		cate_Lock = False
		
		If Len(cate_URL) > 0 Then cate_OutLink = True Else cate_OutLink = False
		If Len(cate_Order) < 1 Then cate_Order = Int(Sys.doGet("Select Count(*) From blog_Category")(0)) + 1
		
		Category.cate_Order = cate_Order
		Category.cate_icon = cate_icon
		Category.cate_Name = cate_Name
		Category.cate_Intro = cate_Intro
		Category.cate_Folder = cate_Folder
		Category.cate_URL = cate_URL
		Category.cate_local = cate_local
		Category.cate_Secret = cate_Secret
		Category.cate_count = 0
		Category.cate_OutLink = cate_OutLink
		Category.cate_Lock = cate_Lock
		Str = Category.Add
		If Err.Number > 0 Then
			Response.Write("{Suc : false, Info : '" & Err.Description & "'}")
		Else
			Response.Write("{Suc : " & Lcase(Str(0)) & ", Info : '" & Str(1) & "'}")
		End If
	End Sub
	
	Public Sub edit
		cate_ID = Split(Trim(Asp.CheckStr(Request.Form("cate_ID"))), ",")
		cate_Order = Split(Trim(Asp.CheckStr(Request.Form("cate_Order"))), ",")
		cate_icon = Split(Trim(Asp.CheckStr(Request.Form("Cate_icons"))), ",")
		cate_Name = Split(Trim(Asp.CheckStr(Request.Form("cate_Name"))), ",")
		cate_Intro = Split(Trim(Asp.CheckStr(Request.Form("cate_Intro"))), ",")
		cate_Folder = Split(Trim(Asp.CheckStr(Request.Form("Cate_Part"))), ",")
		cate_URL = Split(Trim(Asp.CheckStr(Request.Form("cate_URL"))), ",")
		cate_local = Split(Trim(Asp.CheckStr(Request.Form("cate_local"))), ",")
		cate_Secret = Split(Trim(Asp.CheckStr(Request.Form("cate_Secret"))), ",")
		cate_count = Split(Trim(Asp.CheckStr(Request.Form("cate_count"))), ",")
		
		If UBound(cate_ID) >= 0 Then
			ErrorJoin = ""
			On Error Resume Next
			For i = 0 To UBound(cate_ID)
				Category.cate_ID = Int(cate_ID(i))
				If Len(cate_Order(i)) > 0 Then
					Category.cate_Order = Int(cate_Order(i))
				Else
					Category.cate_Order = Int(Sys.doGet("Select Count(*) From blog_Category")(0)) + 1
				End If
				Category.cate_icon = Trim(cate_icon(i))
				Category.cate_Name = Trim(cate_Name(i))
				Category.cate_Intro = Trim(cate_Intro(i))
				Category.cate_Folder = Trim(cate_Folder(i))
				Category.cate_URL = Trim(cate_URL(i))
				Category.cate_local = Int(cate_local(i))
				Category.cate_Secret = CBool(cate_Secret(i))
				Category.cate_count = Int(cate_count(i))
				If Len(Trim(cate_URL(i))) > 0 Then Category.cate_OutLink = True Else Category.cate_OutLink = False
				Str = Category.edit
				If Not Str(0) Then ErrorJoin = ErrorJoin & Str(1) & " - "
			Next
			If len(ErrorJoin) > 0 Then ErrorJoin = Mid(ErrorJoin, 1, (Len(ErrorJoin) - 3))
			If Len(ErrorJoin) > 0 Then Str = ErrorJoin Else Str = "更新分类信息成功!"
			If Err.Number > 0 Then
				Response.Write("{Suc : false, Info : '" & Err.Description & "'}")
			Else
				Response.Write("{Suc : true, Info : '" & Str & "'}")
			End If
		Else
			Response.Write("{Suc : false, Info : '不需要更新数据'}")
		End If
	End Sub
	
	Private Sub del
		ErrorJoin = ""
		cate_ID = Split(Trim(Asp.CheckStr(Request.Form("SelectID"))), ",")
		If UBound(cate_ID) >= 0 Then
			On Error Resume Next
			For i = 0 To UBound(cate_ID)
				Category.cate_ID = Int(cate_ID(i))
				Str = Category.del
				If Not Str(0) Then ErrorJoin = ErrorJoin & Str(1) & " - "
			Next
			If len(ErrorJoin) > 0 Then ErrorJoin = Mid(ErrorJoin, 1, (Len(ErrorJoin) - 3))
			If Len(ErrorJoin) > 0 Then Str = ErrorJoin Else Str = "删除分类成功!"
			If Err.Number > 0 Then
				Response.Write("{Suc : false, Info : '" & Err.Description & "'}")
			Else
				Response.Write("{Suc : true, Info : '" & Str & "'}")
			End If
		Else
			Response.Write("{Suc : false, Info : '不需要删除数据'}")
		End If
	End Sub
	
	Private Sub Clear
		Application.Lock
		Control.FreeApplicationMemory
		Application.UnLock
		Response.Write "<br/><span><b style='color:#040'>缓存清理完毕...	</b></span>"
	End Sub
	
End Class
%>