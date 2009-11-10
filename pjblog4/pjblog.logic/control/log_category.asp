<!--#include file = "../../include.asp" -->
<!--#include file = "../../pjblog.data/control/cls_category.asp" -->
<%
Dim doCategory
Set doCategory = New do_Category
Set doCategory = Nothing

Class do_Category

	Private cate_ID, cate_Order, cate_icon, cate_Name, cate_Intro, cate_Folder, cate_URL, cate_local, cate_Secret, cate_count, cate_OutLink, cate_Lock
	
	Private Str
	
	Public Property Get Action
		Action = Asp.CheckStr(Request.QueryString("action"))
	End Property
	
	' ***********************************************
	'	类初始化
	' ***********************************************
	Private Sub Class_Initialize
		Select Case Action
			Case "add" Call Add
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
	
End Class
%>