<!--#include file = "../include.asp" -->
<%
'*************** PJblog4 基本设置 *******************
' PJblog4 Copyright 2009
' Update: 2009-11-03
' Author: evio
' Owner : PuterJam
' Mode  : Class
' MoName: Ajax Form
'***************************************************
Dim Ajax
Set Ajax = New Sys_Ajax
Set Ajax = Nothing

Class Sys_Ajax

	Private Action

	' ***********************************************
	'	整站初始化类初始化
	' ***********************************************
	Private Sub Class_Initialize
		Action = Asp.CheckStr(Request.QueryString("action"))
		Select Case Action
			Case "CheckCode" Call CheckCode
		End Select
    End Sub 
     
	' ***********************************************
	'	整站初始化类终结化
	' ***********************************************
    Private Sub Class_Terminate
		Sys.Close
    End Sub
	
	Private Sub CheckCode
		Response.Write(Trim(Session("GetCode")))
	End Sub

End Class
%>