<%
'*************** PJblog4 基本设置 *******************
' PJblog4 Copyright 2009
' Update: 2009-11-03
' Author: evio
' Owner : PuterJam
' Mode  : Class
' MoName: User Model
'***************************************************
Dim User : Set User = New Sys_User

Class Sys_User

	' ***********************************************
	'	用户操作方法类初始化
	' ***********************************************
	Private Sub Class_Initialize
    End Sub 
     
	' ***********************************************
	'	用户操作方法类终结化
	' ***********************************************
    Private Sub Class_Terminate
    End Sub
	
	' ***********************************************
	'	用户登入
	' ***********************************************
	Public Function UserLogin(ByVal UserName, ByVal Password, ByVal validate, ByVal KeepLogin)
		Dim ReInfo, HashKey, memLogin, SQL
		UserName = Asp.CheckStr(UserName)
    	Password = Asp.CheckStr(Password)
		validate = Trim(validate)
		
		ReInfo = Array(lang.Set.Asp(6), "", "MessageIcon", False)
		
		' -------------------------------------------------
		'	判断用户名是否为空
		' -------------------------------------------------
		If Trim(UserName) = "" Or Trim(Password) = "" Then
			ReInfo(0) = lang.Set.Asp(6)
			ReInfo(1) = "<b>" & lang.Set.Asp(7) & "</b><br/><a href=""javascript:history.go(-1);"">" & lang.Set.Asp(8) & "</a><br/>"
			ReInfo(2) = "WarningIcon"
			UserLogin = ReInfo
			Init.logout(False)
			Exit Function
		End If
		
		' -------------------------------------------------
		'	判断验证码是否为空
		' -------------------------------------------------
		If validate = "" Then
			ReInfo(0) = lang.Set.Asp(6)
			ReInfo(1) = "<b>" & lang.Set.Asp(9) & "</b><br/><a href=""javascript:history.go(-1);"">" & lang.Set.Asp(8) & "</a><br/>"
			ReInfo(2) = "WarningIcon"
			UserLogin = ReInfo
			Init.logout(False)
			Exit Function
		End If
		
		' -------------------------------------------------
		'	判断是否为有效的用户名
		' -------------------------------------------------
		If Asp.IsValidUserName(UserName) = False Then
			ReInfo(0) = lang.Set.Asp(6)
			ReInfo(1) = "<b>" & lang.Set.Asp(10) & "</b><br/><a href=""javascript:history.go(-1);"">" & lang.Set.Asp(8) & "</a><br/>"
			ReInfo(2) = "ErrorIcon"
			UserLogin = ReInfo
			Init.logout(False)
			Exit Function
		End If
		
		' -------------------------------------------------
		'	判断验证码是否正确
		' -------------------------------------------------
		If CStr(LCase(Session("GetCode"))) <> CStr(LCase(validate)) Then
			ReInfo(0) = lang.Set.Asp(6)
			ReInfo(1) = "<b>" & lang.Set.Asp(11) & "</b><br/><a href=""javascript:history.go(-1);"">" & lang.Set.Asp(8) & "</a><br/>"
			ReInfo(2) = "ErrorIcon"
			UserLogin = ReInfo
			Init.logout(False)
			Exit Function
		End If
		
		' -------------------------------------------------
		'	定义hashkey随机数
		' -------------------------------------------------
		HashKey = SHA1(randomStr(6) & Now())
		
		' -------------------------------------------------
		'	开始查询匹配数据库信息
		' -------------------------------------------------
		Set memLogin = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT Top 1 mem_Name, mem_Password, mem_salt, mem_Status, mem_LastIP, mem_lastVisit, mem_hashKey FROM blog_member WHERE mem_Name='" & UserName & "' AND mem_salt<>''"
		memLogin.Open SQL, conn, 1, 3
		If memLogin.EOF And memLogin.BOF Then
			memLogin.Close
			SQL = "SELECT Top 1 mem_Name,mem_Password,mem_salt,mem_Status,mem_LastIP,mem_lastVisit,mem_hashKey FROM blog_member WHERE mem_Name='" & UserName & "' AND mem_Password='" & MD5(Password) & "'"
			memLogin.Open SQL, conn, 1, 3
			If memLogin.EOF And memLogin.BOF Then
				ReInfo(0) = lang.Set.Asp(6)
				ReInfo(1) = "<b>" & lang.Set.Asp(12) & "</b><br/><a href=""javascript:history.go(-1);"">" & lang.Set.Asp(8) & "</a><br/>"
				ReInfo(2) = "ErrorIcon"
				Init.logout(False)
			Else
				'进行MD5密码验证，转换旧帐户密码验证方式
				Dim strSalt
				strSalt = Asp.randomStr(6)
				memLogin("mem_salt") = strSalt
				memLogin("mem_LastIP") = Asp.getIP
				memLogin("mem_lastVisit") = Now()
				memLogin("mem_hashKey") = HashKey
				memLogin("mem_Password") = SHA1(Password & strSalt)
				Response.Cookies(Sys.CookieName)("memName") = memLogin("mem_Name")
				Response.Cookies(Sys.CookieName)("memHashKey") = HashKey
				If Int(KeepLogin) = 1 Then
					Response.Cookies(Sys.CookieName).Expires = Date + 365
					Response.Cookies(Sys.CookieName)("exp") = DateAdd("d", 365, date())
					Response.Cookies(Sys.CookieName)("DisValidate") = blog_validate
				End If
				memLogin.Update
				ReInfo(0) = lang.Set.Asp(13)
				ReInfo(1) = lang.Set.Asp(14)(memLogin("mem_Name")) & "<br/><a href=""default.asp"">" & lang.Set.Asp(15) & "</a><meta http-equiv=""refresh"" content=""3;url=default.asp""/>"
				ReInfo(2) = "MessageIcon"
				ReInfo(3) = True
			End If
		Else
			If memLogin("mem_Password") <> SHA1(Password & memLogin("mem_salt")) Then
				ReInfo(0) = lang.Set.Asp(6)
				ReInfo(1) = "<b>" & lang.Set.Asp(12) & "</b><br/><a href=""javascript:history.go(-1);"">" & lang.Set.Asp(8) & "</a><br/>"
				ReInfo(2) = "ErrorIcon"
				Init.logout(False)
			Else
				memLogin("mem_LastIP") = Asp.getIP
				memLogin("mem_lastVisit") = Now()
				memLogin("mem_hashKey") = HashKey
				Response.Cookies(CookieName)("memName") = memLogin("mem_Name")
				Response.Cookies(CookieName)("memHashKey") = HashKey
				If Int(KeepLogin) = 1 Then
					Response.Cookies(Sys.CookieName).Expires = Date + 365
					Response.Cookies(Sys.CookieName)("exp") = DateAdd("d", 365, date())
					Response.Cookies(Sys.CookieName)("DisValidate") = blog_validate
				End If
				memLogin.Update
				ReInfo(0) = lang.Set.Asp(13)
				ReInfo(1) = lang.Set.Asp(14)(memLogin("mem_Name")) & "<br/><a href=""default.asp"">" & lang.Set.Asp(15) & "</a><meta http-equiv=""refresh"" content=""3;url=default.asp""/>"
				ReInfo(2) = "MessageIcon"
				ReInfo(3) = True
			End If
		End If
		memLogin.Close
		Set memLogin = Nothing
		UserLogin = ReInfo
	End Function
	
End Class
%>