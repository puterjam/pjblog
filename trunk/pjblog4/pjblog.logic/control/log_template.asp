<!--#include file = "../../include.asp" -->
<!--#include file = "../../pjblog.data/control/cls_template.asp" -->
<!--#include file = "../../pjblog.model/cls_fso.asp" -->
<!--#include file = "../../pjblog.model/cls_xml.asp" -->
<%
Dim doTemp
Set doTemp = New do_Template
Set doTemp = Nothing

Class do_Template

	Private fso, cxml, OK

	Public Property Get Action
		Action = Request.QueryString("action")
	End Property
	
	' ***********************************************
	'	类初始化
	' ***********************************************
	Private Sub Class_Initialize
		Select Case Action
			Case "Choose" Call Choose
			Case "update" Call doUpdate
		End Select
    End Sub 
     
	' ***********************************************
	'	类终结化
	' ***********************************************
    Private Sub Class_Terminate
		Sys.Close
    End Sub
	
	Private Sub Choose
		Dim f1, ReturnStr, Arrays, m, Str, cname
		f1 = Trim(Asp.CheckStr(Request.Form("f1")))
		If Len(f1) > 0 Then
			Set fso = New cls_fso
			Set cxml = New xml
				ReturnStr = fso.FolderItem("../../pjblog.template/" & f1 & "/style")
				Arrays = Split(ReturnStr, "|")
				If Int(Arrays(0)) > 0 Then
					Str = ""
					On Error Resume Next
					For m = 1 To UBound(Arrays)
						If fso.FileExists("../../pjblog.template/" & f1 & "/style/" & Arrays(m) & "/style.xml") Then
							cxml.FilePath = "../../pjblog.template/" & f1 & "/style/" & Arrays(m) & "/style.xml"
							If cxml.open Then
								cname = cxml.GetNodeText(cxml.FindNode("//SkinSet/SkinName"))
								If Err Then Err.Clear : cname = ""
								If Arrays(m) = blog_style And blog_DefaultSkin = f1 Then
									Str = Str & "<div><input type=\'radio\' value=\'" & Arrays(m) & "\' name=\'f2\' checked>" & cname & "</div><br />"
								Else
									Str = Str & "<div><input type=\'radio\' value=\'" & Arrays(m) & "\' name=\'f2\'>" & cname & "</div><br />"
								End If
							End If
						End If
					Next
					If Len(Str) = 0 Then Response.Write("{Suc:false, Info:'没有样式'}") : Exit Sub
				Else
					Response.Write("{Suc:false, Info:'没有样式'}")
				End If
			Set cxml = Nothing
			Set fso = Nothing
			Response.Write("{Suc:true,Info:'" & Str & "'}")
		Else
			Response.Write("{Suc:false, Info:'路径错误'}")
		End If
	End Sub
	
	Private Sub doUpdate
		Dim f1, f2
		f1 = Trim(Asp.CheckStr(Request.Form("f1")))
		f2 = Trim(Asp.CheckStr(Request.Form("f2")))
		If Asp.IsBlank(f1) Or Asp.IsBlank(f2) Then
			Session(Sys.CookieName & "_ShowMsg") = True
			Session(Sys.CookieName & "_MsgText") = "应用主题失败, 请选择样式或主题!"
			RedirectUrl(RedoUrl)
		End If
		OK = Temp.doUpdate(f1, f2)
		Session(Sys.CookieName & "_ShowMsg") = True
		If OK(0) Then
			Cache.GlobalCache(2)
			Session(Sys.CookieName & "_MsgText") = "应用主题成功!"
		Else
			Session(Sys.CookieName & "_MsgText") = OK(1)
		End If
		RedirectUrl(RedoUrl)
	End Sub
	
End Class
%>