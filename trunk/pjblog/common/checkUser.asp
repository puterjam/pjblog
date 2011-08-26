<!--#include file="md5.asp" -->
<!--#include file="sha1.asp" -->
<%
'===============================================================
'  Check User For PJblog2
'    更新时间: 2006-5-29
'===============================================================
Dim UserID, memName, memStatus
memStatus = "Guest"

Function login(UserName, Password)
    Dim validate, ReInfo, HashKey
    UserName = CheckStr(UserName)
    Password = CheckStr(Password)

    validate = Trim(request.Form("validate"))
    ReInfo = Array("错误信息", "", "MessageIcon", False)
    If Trim(UserName) = "" Or Trim(Password) = "" Then
        ReInfo(0) = "错误信息"
        ReInfo(1) = "<b>请将信息输入完整</b><br/><a href=""javascript:history.go(-1);"">请返回重新输入</a><br/>"
        ReInfo(2) = "WarningIcon"
        login = ReInfo
        logout(False)
        Exit Function
    End If

    If validate = "" Then
        ReInfo(0) = "错误信息"
        ReInfo(1) = "<b>请输入登录验证码</b><br/><a href=""javascript:history.go(-1);"">请返回重新输入</a><br/>"
        ReInfo(2) = "WarningIcon"
        login = ReInfo
        logout(False)
        Exit Function
    End If

    If IsValidUserName(UserName) = False Then
        ReInfo(0) = "错误信息"
        ReInfo(1) = "<b>非法用户名！<br/>请尝试使用其他用户名！</b><br/><a href=""javascript:history.go(-1);"">单击返回</a><br/>"
        ReInfo(2) = "ErrorIcon"
        login = ReInfo
        logout(False)
        Exit Function
    End If

    If CStr(LCase(Session("GetCode")))<>CStr(LCase(validate)) Then
        ReInfo(0) = "错误信息"
        ReInfo(1) = "<b>验证码有误，请返回重新输入</b><br/><a href=""javascript:history.go(-1);"">请返回重新输入</a><br/>"
        ReInfo(2) = "ErrorIcon"
        login = ReInfo
        logout(False)
        Exit Function
    End If
    HashKey = SHA1(randomStr(6)&Now())
    Dim memLogin
    Set memLogin = Server.CreateObject("ADODB.Recordset")
    SQL = "SELECT Top 1 mem_Name,mem_Password,mem_salt,mem_Status,mem_LastIP,mem_lastVisit,mem_hashKey FROM blog_member WHERE mem_Name='"&UserName&"' AND mem_salt<>''"
    memLogin.Open SQL, conn, 1, 3
    SQLQueryNums = SQLQueryNums + 1
    If memLogin.EOF And memLogin.BOF Then
        memLogin.Close
        SQL = "SELECT Top 1 mem_Name,mem_Password,mem_salt,mem_Status,mem_LastIP,mem_lastVisit,mem_hashKey FROM blog_member WHERE mem_Name='"&UserName&"' AND mem_Password='"&md5(Password)&"'"
        memLogin.Open SQL, conn, 1, 3
        SQLQueryNums = SQLQueryNums + 1
        If memLogin.EOF And memLogin.BOF Then
            ReInfo(0) = "错误信息"
            ReInfo(1) = "<b>用户名与密码错误</b><br/><a href=""javascript:history.go(-1);"">请返回重新输入</a><br/>"
            ReInfo(2) = "ErrorIcon"
            logout(False)
        Else
            '进行MD5密码验证，转换旧帐户密码验证方式
            Dim strSalt
            strSalt = randomStr(6)
            memLogin("mem_salt") = strSalt
            memLogin("mem_LastIP") = getIP()
            memLogin("mem_lastVisit") = Now()
            memLogin("mem_hashKey") = HashKey
            memLogin("mem_Password") = SHA1(Password&strSalt)
            Response.Cookies(CookieName)("memName") = memLogin("mem_Name")
            Response.Cookies(CookieName)("memHashKey") = HashKey
            If Request.Form("KeepLogin") = "1" Then
                Response.Cookies(CookieName).Expires = Date+365
                Response.Cookies(CookieName)("exp") = DateAdd("d", 365, date())
				Response.Cookies(CookieName)("DisValidate") = blog_validate
            End If
            memLogin.Update
            ReInfo(0) = "登录成功"
            ReInfo(1) = "<b>"&memLogin("mem_Name")&"</b>，欢迎你的再次光临。<br/><a href=""default.asp"">点击返回主页</a><meta http-equiv=""refresh"" content=""3;url=default.asp""/>"
            ReInfo(2) = "MessageIcon"
            ReInfo(3) = True
        End If
    Else
        If memLogin("mem_Password")<>SHA1(Password&memLogin("mem_salt")) Then
            ReInfo(0) = "错误信息"
            ReInfo(1) = "<b>用户名与密码错误</b><br/><a href=""javascript:history.go(-1);"">请返回重新输入</a><br/>"
            ReInfo(2) = "ErrorIcon"
            logout(False)
        Else
            memLogin("mem_LastIP") = getIP()
            memLogin("mem_lastVisit") = Now()
            memLogin("mem_hashKey") = HashKey
            Response.Cookies(CookieName)("memName") = memLogin("mem_Name")
            Response.Cookies(CookieName)("memHashKey") = HashKey
            If Request.Form("KeepLogin") = "1" Then
                Response.Cookies(CookieName).Expires = Date+365
                Response.Cookies(CookieName)("exp") = DateAdd("d", 365, date())
				Response.Cookies(CookieName)("DisValidate") = blog_validate
            End If
            memLogin.Update
            ReInfo(0) = "登录成功"
            ReInfo(1) = "<b>"&memLogin("mem_Name")&"</b>，欢迎你的再次光临。<br/><a href=""default.asp"">点击返回主页</a><meta http-equiv=""refresh"" content=""3;url=default.asp""/>"
            ReInfo(2) = "MessageIcon"
            ReInfo(3) = True
        End If
    End If
    memLogin.Close
    Set memLogin = Nothing
    login = ReInfo
End Function

Function login2(UserName, Password)
    Dim validate, ReInfo, HashKey
    UserName = CheckStr(UserName)
    Password = CheckStr(Password)

    ReInfo = Array("错误信息", "", "MessageIcon", False)
    If Trim(UserName) = "" Or Trim(Password) = "" Then
        ReInfo(0) = "错误信息"
        ReInfo(1) = "<b>请将信息输入完整</b><br/><a href=""javascript:history.go(-1);"">请返回重新输入</a><br/>"
        ReInfo(2) = "WarningIcon"
        login2 = ReInfo
        logout(False)
        UserRight(1)
        Exit Function
    End If

    If IsValidUserName(UserName) = False Then
        ReInfo(0) = "错误信息"
        ReInfo(1) = "<b>非法用户名！<br/>请尝试使用其他用户名！</b><br/><a href=""javascript:history.go(-1);"">单击返回</a><br/>"
        ReInfo(2) = "ErrorIcon"
        login2 = ReInfo
        logout(False)
        UserRight(1)
        Exit Function
    End If

    HashKey = SHA1(randomStr(6)&Now())
    Dim memLogin
    Set memLogin = Server.CreateObject("ADODB.Recordset")
    SQL = "SELECT Top 1 mem_Name,mem_Password,mem_salt,mem_Status,mem_LastIP,mem_lastVisit,mem_hashKey FROM blog_member WHERE mem_Name='"&UserName&"' AND mem_salt<>''"
    memLogin.Open SQL, conn, 1, 3
    SQLQueryNums = SQLQueryNums + 1
    If memLogin.EOF And memLogin.BOF Then
        memLogin.Close
        SQL = "SELECT Top 1 mem_Name,mem_Password,mem_salt,mem_Status,mem_LastIP,mem_lastVisit,mem_hashKey FROM blog_member WHERE mem_Name='"&UserName&"' AND mem_Password='"&md5(Password)&"'"
        memLogin.Open SQL, conn, 1, 3
        SQLQueryNums = SQLQueryNums + 1
        If memLogin.EOF And memLogin.BOF Then
            ReInfo(0) = "错误信息"
            ReInfo(1) = "<b>用户名与密码错误</b><br/><a href=""javascript:history.go(-1);"">请返回重新输入</a><br/>"
            ReInfo(2) = "ErrorIcon"
            logout(False)
        Else
            '进行MD5密码验证，转换旧帐户密码验证方式
            Dim strSalt
            strSalt = randomStr(6)
            memLogin("mem_salt") = strSalt
            memLogin("mem_LastIP") = getIP()
            memLogin("mem_lastVisit") = Now()
            memLogin("mem_hashKey") = HashKey
            memLogin("mem_Password") = SHA1(Password&strSalt)
            memLogin.Update
            memName = memLogin("mem_Name")
            memStatus = memLogin("mem_Status")
            ReInfo(0) = "登录成功"
            ReInfo(1) = "<b>"&memLogin("mem_Name")&"</b>，欢迎你的再次光临。<br/><a href=""default.asp"">点击返回主页</a><meta http-equiv=""refresh"" content=""3;url=default.asp""/>"
            ReInfo(2) = "MessageIcon"
            ReInfo(3) = True
        End If
    Else
        If memLogin("mem_Password")<>SHA1(Password&memLogin("mem_salt")) Then
            ReInfo(0) = "错误信息"
            ReInfo(1) = "<b>用户名与密码错误</b><br/><a href=""javascript:history.go(-1);"">请返回重新输入</a><br/>"
            ReInfo(2) = "ErrorIcon"
            logout(False)
        Else
            memName = memLogin("mem_Name")
            memStatus = memLogin("mem_Status")
            ReInfo(0) = "登录成功"
            ReInfo(1) = "<b>"&memLogin("mem_Name")&"</b>，欢迎你的再次光临。<br/><a href=""default.asp"">点击返回主页</a><meta http-equiv=""refresh"" content=""3;url=default.asp""/>"
            ReInfo(2) = "MessageIcon"
            ReInfo(3) = True
        End If
    End If
    UserRight(1)
    memLogin.Close
    Set memLogin = Nothing
    login2 = ReInfo
End Function

Sub checkCookies()
    Dim Guest_IP, Guest_Browser, Guest_Refer
    Guest_IP = getIP()
    Guest_Browser = getBrowser(Request.ServerVariables("HTTP_USER_AGENT"))

    If Session("GuestIP")<>Guest_IP Then
        Conn.Execute("UPDATE blog_Info SET blog_VisitNums=blog_VisitNums+1")
        SQLQueryNums = SQLQueryNums + 1
        getInfo(2)
        Session("GuestIP") = Guest_IP
        If blog_CountNum>0 And Guest_Browser(1)<>"Unkown" Then
            Dim tmpC
            tmpC = conn.Execute("select count(coun_ID) as cnt from [blog_Counter]")(0)
            SQLQueryNums = SQLQueryNums + 1
            Guest_Refer = Trim(Request.ServerVariables("HTTP_REFERER"))
            If tmpC>= blog_CountNum Then
                Dim tmpLC
                tmpLC = conn.Execute("select top 1 coun_ID from [blog_Counter] order by coun_Time ASC")(0)
                Conn.Execute("update [blog_Counter] set coun_Time=#"&DateToStr(Now(), "Y-m-d H:I:S")&"#,coun_IP='"&Guest_IP&"',coun_OS='"&Guest_Browser(1)&"',coun_Browser='"&Guest_Browser(0)&"',coun_Referer='"&HTMLEncode(CheckStr(Guest_Refer))&"' where coun_ID="&tmpLC)
                SQLQueryNums = SQLQueryNums + 2
            Else
                Conn.Execute("INSERT INTO blog_Counter(coun_IP,coun_OS,coun_Browser,coun_Referer) VALUES ('"&Guest_IP&"','"&Guest_Browser(1)&"','"&Guest_Browser(0)&"','"&HTMLEncode(CheckStr(Guest_Refer))&"')")
                SQLQueryNums = SQLQueryNums + 1
            End If
        End If
    End If

    Dim tempName, tempHashKey
    tempName = CheckStr(Request.Cookies(CookieName)("memName"))
    tempHashKey = CheckStr(Request.Cookies(CookieName)("memHashKey"))
    If tempHashKey = "" Then
        logout(False)
    Else
        Dim CheckCookie
        Set CheckCookie = Server.CreateObject("ADODB.RecordSet")
        SQL = "SELECT Top 1 mem_ID,mem_Name,mem_Password,mem_salt,mem_Status,mem_LastIP,mem_lastVisit,mem_hashKey FROM blog_member WHERE mem_Name='"&tempName&"' AND mem_hashKey='"&tempHashKey&"' AND mem_hashKey<>''"
        CheckCookie.Open SQL, Conn, 1, 1
        SQLQueryNums = SQLQueryNums + 1
        If CheckCookie.EOF And CheckCookie.BOF Then
            logout(False)
        Else
            UserID = CheckCookie("mem_ID")
'            If CheckCookie("mem_LastIP")<>Guest_IP Or IsNull(CheckCookie("mem_LastIP")) Then
'                logout(True)
'            Else
                memName = CheckStr(Request.Cookies(CookieName)("memName"))
                memStatus = CheckCookie("mem_Status")
'            End If
        End If
        CheckCookie.Close
        Set CheckCookie = Nothing
    End If

End Sub

Sub logout(clearHashKey)
    On Error Resume Next
	Response.Cookies(CookieName)("DisValidate") = "False"
    If clearHashKey Then conn.Execute("UPDATE blog_member set mem_hashKey='' where mem_ID="&UserID)
    If Err Then Err.Clear
    Response.Cookies(CookieName)("memName") = ""
    Response.Cookies(CookieName)("memHashKey") = ""
    memStatus = "Guest"
End Sub
%>
