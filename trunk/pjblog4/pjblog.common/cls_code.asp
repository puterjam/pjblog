<%
'*************** PJblog4 基本设置 *******************
' PJblog4 Copyright 2009
' Update: 2009-11-03
' Author: evio
' Owner : PuterJam
' Mode  : Class
' MoName: Function
'***************************************************
Dim Asp : Set Asp = New Sys_Asp

Class Sys_Asp
	
	' ***********************************************
	'	类初始化
	' ***********************************************
	Private Sub Class_Initialize
    End Sub 
     
	' ***********************************************
	'	类终结化
	' ***********************************************
    Private Sub Class_Terminate
    End Sub
	
	Public Function CheckPage(ByVal Pages)
		If Not IsInteger(Pages) Then
			Pages = 1
		Else
			If Pages = 0 Then Pages = 1
		End If
		CheckPage = Pages
	End Function
	
	Public Function MessageInfo(ByVal Arrays)
		If Not IsArray(Arrays) Then Exit Function
		MessageInfo = ""
		MessageInfo = MessageInfo & "<br/><br/>"
   		MessageInfo = MessageInfo & "<div style=""text-align:center;"">"
    	MessageInfo = MessageInfo & "<div id=""MsgContent"" style=""width:300px"">"
      	MessageInfo = MessageInfo & "<div id=""MsgHead"">" & Arrays(0) & "</div>"
      	MessageInfo = MessageInfo & "<div id=""MsgBody"">"
	   	MessageInfo = MessageInfo & "<div class=""" & Arrays(2) & """></div>"
       	MessageInfo = MessageInfo & "<div class=""MessageText"">" & Arrays(1) & "</div>"
	  	MessageInfo = MessageInfo & "</div>"
		MessageInfo = MessageInfo & "</div>"
  		MessageInfo = MessageInfo & "</div><br/><br/>"
	End Function
	
	' ***********************************************
	'	过滤HTML标签
	' ***********************************************
	Public Function FilterHtmlTags(ByVal s_Description)
		If IsBlank(s_Description) Then Exit Function
		Dim FaStr, re
		Set re = New RegExp
			re.IgnoreCase = True
			re.Global = True
			re.Pattern = "<[^>]*?>"
				
			'去掉 尖括号和换行
			FaStr = re.replace(s_Description, "")
			FaStr = replace(FaStr,Chr(13), "")
			FaStr = replace(FaStr,Chr(10), "")
		Set re = nothing
		FilterHtmlTags = FaStr
	End Function
	
	' ***********************************************
	'	判断是否为空类型
	' ***********************************************
	Public Function IsBlank(ByRef TempVar) 
		IsBlank = False 
		TempVar = Trim(TempVar)
		Select Case VarType(TempVar)
			Case 0, 1 							'Empty & Null 
				IsBlank = True 
			Case 8 								'String 
				If Len(TempVar) = 0 Then 
					IsBlank = True 
				End If 
			Case 9 								'Object 
				Dim tmpType 
				tmpType = TypeName(TempVar) 
				If (tmpType = "Nothing") or (tmpType = "Empty") Then IsBlank = True 
			Case 8192, 8204, 8209 				'Array 
				If UBound(TempVar) = -1 Then IsBlank = True
		End select 
	End Function  
	
	' ***********************************************
	'	客户端获取IP函数
	' ***********************************************
	Public Function GetIP
		Dim strIP, IP_Ary, strIP_list
		strIP_list = Replace(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), "'", "")
		If InStr(strIP_list, ",")<>0 Then
			IP_Ary = Split(strIP_list, ",")
			strIP = IP_Ary(0)
		Else
			strIP = strIP_list
		End If
		If strIP = Empty Then strIP = Replace(Request.ServerVariables("REMOTE_ADDR"), "'", "")
		getIP = strIP
	End Function
	
	' ***********************************************
	'	获取随机数函数
	' ***********************************************
	Public Function RandomStr(ByVal intLength)
		Dim strSeed, seedLength, pos, Str, i
		strSeed = "abcdefghijklmnopqrstuvwxyz1234567890"
		seedLength = Len(strSeed)
		Str = ""
		Randomize
		For i = 1 To intLength
			Str = Str + Mid(strSeed, Int(seedLength * Rnd) + 1, 1)
		Next
		RandomStr = Str
	End Function
	
	'*****************************************************
	'	截取字符串长度
	'*****************************************************
	Public Function CutStr(byVal Str, byVal StrLen)
		Dim l, t, c, i
		If IsNull(Str) Then
			CutStr = ""
			Exit Function
		End If
		l = Len(Str)
		StrLen = Int(StrLen)
		t = 0
		For i = 1 To l
			c = Asc(Mid(Str, i, 1))
			If c<0 Or c>255 Then t = t + 2 Else t = t + 1
			If t>= StrLen Then
				CutStr = Left(Str, i)&"..."
				Exit For
			Else
				CutStr = Str
			End If
		Next
	End Function
	
	' ***********************************************
	'	字符串过滤函数
	' ***********************************************
	Public Function CheckStr(byVal ChkStr)
		Dim Str
		Str = ChkStr
		If IsNull(Str) Then
			CheckStr = ""
			Exit Function
		End If
		Str = Replace(Str, "&", "&amp;")
		Str = Replace(Str, "'", "&#39;")
		Str = Replace(Str, """", "&#34;")
		Str = Replace(Str, "<", "&lt;")
		Str = Replace(Str, ">", "&gt;")
		Dim re
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = "(w)(here)"
		Str = re.Replace(Str, "$1h&#101;re")
		re.Pattern = "(s)(elect)"
		Str = re.Replace(Str, "$1el&#101;ct")
		re.Pattern = "(i)(nsert)"
		Str = re.Replace(Str, "$1ns&#101;rt")
		re.Pattern = "(c)(reate)"
		Str = re.Replace(Str, "$1r&#101;ate")
		re.Pattern = "(d)(rop)"
		Str = re.Replace(Str, "$1ro&#112;")
		re.Pattern = "(a)(lter)"
		Str = re.Replace(Str, "$1lt&#101;r")
		re.Pattern = "(d)(elete)"
		Str = re.Replace(Str, "$1el&#101;te")
		re.Pattern = "(u)(pdate)"
		Str = re.Replace(Str, "$1p&#100;ate")
		re.Pattern = "(\s)(or)"
		Str = re.Replace(Str, "$1o&#114;")
		'----------------------------------
		re.Pattern = "(java)(script)"
		Str = re.Replace(Str, "$1scri&#112;t")
		re.Pattern = "(j)(script)"
		Str = re.Replace(Str, "$1scri&#112;t")
		re.Pattern = "(vb)(script)"
		Str = re.Replace(Str, "$1scri&#112;t")
		'----------------------------------
		If Instr(Str, "expression") <> 0 Then
			Str = Replace(Str, "expression", "e&#173;xpression", 1, -1, 0) '防止xss注入
		End If
		Set re = Nothing
		CheckStr= Str
	End Function
	
	' ***********************************************
	'	字符串反过滤函数
	' ***********************************************
	Public Function UnCheckStr(ByVal Str)
		If IsNull(Str) Then
			UnCheckStr = ""
			Exit Function
		End If
		Str = Replace(Str, "&#39;", "'")
		Str = Replace(Str, "&#34;", """")
		Str = Replace(Str, "&lt;", "<")
		Str = Replace(Str, "&gt;", ">")
		Str = Replace(Str, "&amp;", "&")
		Dim re
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = "(w)(h&#101;re)"
		Str = re.Replace(Str, "$1here")
		re.Pattern = "(s)(el&#101;ct)"
		Str = re.Replace(Str, "$1elect")
		re.Pattern = "(i)(ns&#101;rt)"
		Str = re.Replace(Str, "$1nsert")
		re.Pattern = "(c)(r&#101;ate)"
		Str = re.Replace(Str, "$1reate")
		re.Pattern = "(d)(ro&#112;)"
		Str = re.Replace(Str, "$1rop")
		re.Pattern = "(a)(lt&#101;r)"
		Str = re.Replace(Str, "$1lter")
		re.Pattern = "(d)(el&#101;te)"
		Str = re.Replace(Str, "$1elete")
		re.Pattern = "(u)(p&#100;ate)"
		Str = re.Replace(Str, "$1pdate")
		re.Pattern = "(\s)(o&#114;)"
		Str = re.Replace(Str, "$1or")
		'----------------------------------
		re.Pattern = "(java)(scri&#112;t)"
		Str = re.Replace(Str, "$1script")
		re.Pattern = "(j)(scri&#112;t)"
		Str = re.Replace(Str, "$1script")
		re.Pattern = "(vb)(scri&#112;t)"
		Str = re.Replace(Str, "$1script")
		'----------------------------------
		If Instr(Str, "e&#173;xpression") > 0 Then
			Str = Replace(Str, "e&#173;xpression", "expression", 1, -1, 0) '防止xss注入
		End If
		Set re = Nothing
		UnCheckStr = Str
	End Function
	
	'*************************************
	'	过滤html标签
	'*************************************
	Public Function FilterUBBTags(ByVal s_Description)
		If isblank(s_Description) Then Exit Function
		Dim re, eHtml, eobj, en
		Set re = New RegExp
			re.IgnoreCase = True
			re.Global = True
			re.Pattern = "\[(.+)(\=(.+))?\](.*)?\[\/\1\]"
			Set eobj = re.Execute(s_Description)
				For Each en In eobj
					eHtml = eobj(0).SubMatches(3)
					s_Description = Replace(s_Description, en.value, eHtml)
					'去掉 尖括号和换行
					s_Description = Replace(s_Description, Chr(13), "")
					s_Description = Replace(s_Description, Chr(10), "")
				Next
			Set eobj = Nothing
		Set re = nothing
			FilterUBBTags = s_Description
	End Function
	
	' ***********************************************
	'	防止外部提交
	' ***********************************************
	Public Function ChkPost()
		Dim server_v1, server_v2
		chkpost = False
		server_v1 = CStr(Request.ServerVariables("HTTP_REFERER"))
		server_v2 = CStr(Request.ServerVariables("SERVER_NAME"))
		if instr(server_v1, replace(replace(server_v2, "http://", ""), "www.", ""))=0 then
			chkpost = False
		Else
			chkpost = True
		End If
	End Function
	
	'*************************************
	'IP过滤
	'*************************************
	Public Function MatchIP(IP)
		MatchIP = False
		Dim SIp, SplitIP
		For Each SIp in FilterIP
			SIp = Replace(SIp, "*", "\d*")
			SplitIP = Split(SIp, ".")
			Dim re, strMatchs, strIP
			Set re = New RegExp
			re.IgnoreCase = True
			re.Global = True
			re.Pattern = "("&SplitIP(0)&"|)."&"("&SplitIP(1)&"|)."&"("&SplitIP(2)&"|)."&"("&SplitIP(3)&"|)"
			Set strMatchs = re.Execute(IP)
			strIP = strMatchs(0).SubMatches(0) & "." & strMatchs(0).SubMatches(1)& "." & strMatchs(0).SubMatches(2)& "." & strMatchs(0).SubMatches(3)
			If strIP = IP Then
				MatchIP = True
				Exit Function
			End If
			Set strMatchs = Nothing
			Set re = Nothing
		Next
	End Function
	
	'*************************************
	'获得注册码
	'*************************************
	Public Function getcode()
		getcode = "<img id=""vcodeImg"" src=""about:blank"" onerror=""this.onerror=null;this.src='common/getcode.asp?s='+Math.random();"" alt=""" & lang.Set.Asp(4) & """ title=""" & lang.Set.Asp(5) & """ style=""margin-right:40px;cursor:pointer;width:40px;height:18px;margin-bottom:-4px;margin-top:3px;"" onclick=""src='common/getcode.asp?s='+Math.random()""/>"
	End Function
	
	'*************************************
	'限制上传文件类型
	'*************************************
	
	Public Function IsvalidFile(File_Type)
		IsvalidFile = False
		Dim GName
		For Each GName in UP_FileType
			If File_Type = GName Then
				IsvalidFile = True
				Exit For
			End If
		Next
	End Function
	
	'*************************************
	'检测是否只包含英文和数字
	'*************************************
	
	Public Function IsValidChars(Str)
		Dim re, chkstr
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = "[^_\.a-zA-Z\d]"
		IsValidChars = True
		chkstr = re.Replace(Str, "")
		If chkstr<>Str Then IsValidChars = False
		Set re = Nothing
	End Function
	
	'*************************************
	'检测是否只包含英文和数字
	'*************************************
	
	Public Function IsvalidValue(ArrayN, Str)
		IsvalidValue = False
		Dim GName
		For Each GName in ArrayN
			If Str = GName Then
				IsvalidValue = True
				Exit For
			End If
		Next
	End Function
	
	'*************************************
	'检测是否有效的数字
	'*************************************
	
	Public Function IsInteger(Para)
		IsInteger = False
		If Not (IsNull(Para) Or Trim(Para) = "" Or Not IsNumeric(Para)) Then
			IsInteger = True
		End If
	End Function
	
	'*************************************
	'用户名检测
	'*************************************
	
	Public Function IsValidUserName(byVal UserName)
		Dim i, c
		Dim VUserName
		IsValidUserName = True
		For i = 1 To Len(UserName)
			c = LCase(Mid(UserName, i, 1))
			If InStr("$!<>?#^%@~`&*();:+='"" 	", c) > 0 Then
				IsValidUserName = False
				Exit Function
			End If
		Next
		For Each VUserName in Register_UserName
			If UserName = VUserName Then
				IsValidUserName = False
				Exit For
			End If
		Next
	End Function
	
	'*************************************
	'检测是否有效的E-mail地址
	'*************************************
	
	Public Function IsValidEmail(Email)
		Dim names, Name, i, c
		IsValidEmail = True
		Names = Split(email, "@")
		If UBound(names) <> 1 Then
			IsValidEmail = False
			Exit Function
		End If
		For Each Name IN names
			If Len(Name) <= 0 Then
				IsValidEmail = False
				Exit Function
			End If
			For i = 1 To Len(Name)
				c = LCase(Mid(Name, i, 1))
				If InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 And Not IsNumeric(c) Then
					IsValidEmail = False
					Exit Function
				End If
			Next
			If Left(Name, 1) = "." Or Right(Name, 1) = "." Then
				IsValidEmail = False
				Exit Function
			End If
		Next
		If InStr(names(1), ".") <= 0 Then
			IsValidEmail = False
			Exit Function
		End If
		i = Len(names(1)) - InStrRev(names(1), ".")
		If i <> 2 And i <> 3 Then
			IsValidEmail = False
			Exit Function
		End If
		If InStr(email, "..") > 0 Then
			IsValidEmail = False
		End If
	End Function
	
	'*************************************
	'加亮关键字
	'*************************************
	
	Public Function highlight(byVal strContent, byRef arrayWords)
		Dim intCounter, strTemp, intPos, intTagLength, intKeyWordLength, bUpdate
		If Len(arrayWords)<1 Then
			highlight = strContent
			Exit Function
		End If
		For intPos = 1 To Len(strContent)
			bUpdate = False
			If Mid(strContent, intPos, 1) = "<" Then
				On Error Resume Next
				intTagLength = (InStr(intPos, strContent, ">", 1) - intPos)
				If Err Then
					highlight = strContent
					Err.Clear
				End If
				strTemp = strTemp & Mid(strContent, intPos, intTagLength)
				intPos = intPos + intTagLength
			End If
			If arrayWords <> "" Then
				intKeyWordLength = Len(arrayWords)
				If LCase(Mid(strContent, intPos, intKeyWordLength)) = LCase(arrayWords) Then
					strTemp = strTemp & "<span class=""high1"">" & Mid(strContent, intPos, intKeyWordLength) & "</span>"
					intPos = intPos + intKeyWordLength - 1
					bUpdate = True
				End If
			End If
			If bUpdate = False Then
				strTemp = strTemp & Mid(strContent, intPos, 1)
			End If
		Next
		highlight = strTemp
	End Function
	
	'*************************************
	'过滤超链接
	'*************************************
	
	Public Function checkURL(ByVal ChkStr)
		Dim Str
		Str = ChkStr
		Str = Trim(Str)
		If IsNull(Str) Then
			checkURL = ""
			Exit Function
		End If
		Dim re
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = "(d)(ocument\.cookie)"
		Str = re.Replace(Str, "$1ocument cookie")
		re.Pattern = "(d)(ocument\.write)"
		Str = re.Replace(Str, "$1ocument write")
		re.Pattern = "(s)(cript:)"
		Str = re.Replace(Str, "$1cri&#112;t ")
		re.Pattern = "(s)(cript)"
		Str = re.Replace(Str, "$1cri&#112;t")
		re.Pattern = "(o)(bject)"
		Str = re.Replace(Str, "$1bj&#101;ct")
		re.Pattern = "(a)(pplet)"
		Str = re.Replace(Str, "$1ppl&#101;t")
		re.Pattern = "(e)(mbed)"
		Str = re.Replace(Str, "$1mb&#101;d")
		Set re = Nothing
		Str = Replace(Str, ">", "&gt;")
		Str = Replace(Str, "<", "&lt;")
		checkURL = Str
	End Function
	
	'*************************************
	'过滤文件名字
	'*************************************
	
	Public Function FixName(UpFileExt)
		If IsEmpty(UpFileExt) Then Exit Function
		FixName = UCase(UpFileExt)
		FixName = Replace(FixName, Chr(0), "")
		FixName = Replace(FixName, ".", "")
		FixName = Replace(FixName, "ASP", "")
		FixName = Replace(FixName, "ASA", "")
		FixName = Replace(FixName, "ASPX", "")
		FixName = Replace(FixName, "CER", "")
		FixName = Replace(FixName, "CDX", "")
		FixName = Replace(FixName, "HTR", "")
	End Function
	
	'*************************************
	'转换HTML代码
	'*************************************
	
	Public Function HTMLEncode(ByVal reString)
		Dim Str
		Str = reString
		If Not IsNull(Str) Then
			Str = Replace(Str, ">", "&gt;")
			Str = Replace(Str, "<", "&lt;")
			Str = Replace(Str, Chr(9), "&#160;&#160;&#160;&#160;")
			Str = Replace(Str, Chr(32)&Chr(32), "&nbsp;&nbsp;")
			Str = Replace(Str, Chr(39), "&#39;")
			Str = Replace(Str, Chr(34), "&quot;")
			Str = Replace(Str, Chr(13), "")
			Str = Replace(Str, Chr(10), "<br/>")
			HTMLEncode = Str
		End If
	End Function
	
	'*************************************
	'转换最新评论和日志HTML代码
	'*************************************
	
	Public Function CCEncode(ByVal reString)
		Dim Str
		Str = reString
		If Not IsNull(Str) Then
			Str = Replace(Str, ">", "&gt;")
			Str = Replace(Str, "<", "&lt;")
			Str = Replace(Str, Chr(9), "&#160;&#160;&#160;&#160;")
			Str = Replace(Str, Chr(32)&Chr(32), "&nbsp;&nbsp;")
			Str = Replace(Str, Chr(39), "&#39;")
			Str = Replace(Str, Chr(34), "&quot;")
			Str = Replace(Str, Chr(13), "")
			Str = Replace(Str, Chr(10), " ")
			CCEncode = Str
		End If
	End Function
	
	'*************************************
	'反转换HTML代码
	'*************************************
	
	Public Function HTMLDecode(ByVal reString)
		Dim Str
		Str = reString
		If Not IsNull(Str) Then
			Str = Replace(Str, "&gt;", ">")
			Str = Replace(Str, "&lt;", "<")
			Str = Replace(Str, "&#160;&#160;&#160;&#160;", Chr(9))
			Str = Replace(Str, "&nbsp;&nbsp;", Chr(32)&Chr(32))
			Str = Replace(Str, "&#39;", Chr(39))
			Str = Replace(Str, "&quot;", Chr(34))
			Str = Replace(Str, "", Chr(13))
			Str = Replace(Str, "<br/>", Chr(10))
			HTMLDecode = Str
		End If
	End Function
	
	'*************************************
	'恢复&字符
	'*************************************
	
	Public Function ClearHTML(ByVal reString)
		Dim Str
		Str = reString
		If Not IsNull(Str) Then
			Str = Replace(Str, "&amp;", "&")
			ClearHTML = Str
		End If
	End Function
	
	'*************************************
	'过滤textarea
	'*************************************
	
	Public Function UBBFilter(ByVal reString)
		Dim Str
		Str = reString
		If Not IsNull(Str) Then
			Str = Replace(Str, "</textarea>", "<&#47textarea>")
			UBBFilter = Str
		End If
	End Function
	
	'*************************************
	'过滤HTML代码
	'*************************************
	
	Public Function EditDeHTML(byVal Content)
		EditDeHTML = Content
		If Not IsNull(EditDeHTML) Then
			EditDeHTML = UnCheckStr(EditDeHTML)
			EditDeHTML = Replace(EditDeHTML, "&", "&amp;")
			EditDeHTML = Replace(EditDeHTML, "<", "&lt;")
			EditDeHTML = Replace(EditDeHTML, ">", "&gt;")
			EditDeHTML = Replace(EditDeHTML, Chr(34), "&quot;")
			EditDeHTML = Replace(EditDeHTML, Chr(39), "&#39;")
		End If
	End Function
	
	'*************************************
	'日期转换函数
	'*************************************
	
	Public Function DateToStr(DateTime, ShowType)
		Dim DateMonth, DateDay, DateHour, DateMinute, DateWeek, DateSecond
		Dim FullWeekday, shortWeekday, Fullmonth, Shortmonth, TimeZone1, TimeZone2
		TimeZone1 = "+0800"
		TimeZone2 = "+08:00"
		FullWeekday = Array("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
		shortWeekday = Array("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")
		Fullmonth = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
		Shortmonth = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
	
		DateMonth = Month(DateTime)
		DateDay = Day(DateTime)
		DateHour = Hour(DateTime)
		DateMinute = Minute(DateTime)
		DateWeek = Weekday(DateTime)
		DateSecond = Second(DateTime)
		If Len(DateMonth)<2 Then DateMonth = "0"&DateMonth
		If Len(DateDay)<2 Then DateDay = "0"&DateDay
		If Len(DateMinute)<2 Then DateMinute = "0"&DateMinute
		Select Case ShowType
			Case "Y-m-d"
				DateToStr = Year(DateTime)&"-"&DateMonth&"-"&DateDay
			Case "Y-m-d H:I A"
				Dim DateAMPM
				If DateHour>12 Then
					DateHour = DateHour -12
					DateAMPM = "PM"
				Else
					DateHour = DateHour
					DateAMPM = "AM"
				End If
				If Len(DateHour)<2 Then DateHour = "0"&DateHour
				DateToStr = Year(DateTime)&"-"&DateMonth&"-"&DateDay&" "&DateHour&":"&DateMinute&" "&DateAMPM
			Case "Y-m-d H:I:S"
				If Len(DateHour)<2 Then DateHour = "0"&DateHour
				If Len(DateSecond)<2 Then DateSecond = "0"&DateSecond
				DateToStr = Year(DateTime)&"-"&DateMonth&"-"&DateDay&" "&DateHour&":"&DateMinute&":"&DateSecond
			Case "YmdHIS"
				DateSecond = Second(DateTime)
				If Len(DateHour)<2 Then DateHour = "0"&DateHour
				If Len(DateSecond)<2 Then DateSecond = "0"&DateSecond
				DateToStr = Year(DateTime)&DateMonth&DateDay&DateHour&DateMinute&DateSecond
			Case "ym"
				DateToStr = Right(Year(DateTime), 2)&DateMonth
			Case "d"
				DateToStr = DateDay
			Case "ymd"
				DateToStr = Right(Year(DateTime), 4)&DateMonth&DateDay
			Case "mdy"
				Dim DayEnd
				Select Case DateDay
				Case 1
					DayEnd = "st"
				Case 2
					DayEnd = "nd"
				Case 3
					DayEnd = "rd"
				Case Else
					DayEnd = "th"
			End Select
			DateToStr = Fullmonth(DateMonth -1)&" "&DateDay&DayEnd&" "&Right(Year(DateTime), 4)
			Case "w,d m y H:I:S"
				DateSecond = Second(DateTime)
				If Len(DateHour)<2 Then DateHour = "0"&DateHour
				If Len(DateSecond)<2 Then DateSecond = "0"&DateSecond
				DateToStr = shortWeekday(DateWeek -1)&","&DateDay&" "& Left(Fullmonth(DateMonth -1), 3) &" "&Right(Year(DateTime), 4)&" "&DateHour&":"&DateMinute&":"&DateSecond&" "&TimeZone1
			Case "y-m-dTH:I:S"
				If Len(DateHour)<2 Then DateHour = "0"&DateHour
				If Len(DateSecond)<2 Then DateSecond = "0"&DateSecond
				DateToStr = Year(DateTime)&"-"&DateMonth&"-"&DateDay&"T"&DateHour&":"&DateMinute&":"&DateSecond&TimeZone2
			Case Else
				If Len(DateHour)<2 Then DateHour = "0"&DateHour
				DateToStr = Year(DateTime)&"-"&DateMonth&"-"&DateDay&" "&DateHour&":"&DateMinute
		End Select
	End Function
	
	'*************************************
	'切割内容 - 按行分割
	'*************************************
	
	Public Function SplitLines(byVal Content, byVal ContentNums)
		Dim ts, i, l
		ContentNums = Int(ContentNums)
		If IsNull(Content) Then Exit Function
		i = 1
		ts = 0
		For i = 1 To Len(Content)
			l = LCase(Mid(Content, i, 5))
			If l = "<br/>" Then
				ts = ts + 1
			End If
			l = LCase(Mid(Content, i, 4))
			If l = "<br>" Then
				ts = ts + 1
			End If
			l = LCase(Mid(Content, i, 3))
			If l = "<p>" Then
				ts = ts + 1
			End If
			If ts>ContentNums Then Exit For
		Next
		If ts>ContentNums Then
			Content = Left(Content, i -1)
		End If
		SplitLines = Content
	End Function
	
	'*************************************
	'Trackback Function
	'*************************************
	
	Public Function Trackback(trackback_url, url, title, excerpt, blog_name)
		Dim query_string, objXMLHTTP
	
		query_string = "title="&cutStr(Server.URLEncode(title), 100)&"&url="&Server.URLEncode(url)&"&blog_name="&Server.URLEncode(blog_name)&"&excerpt="&cutStr(Server.URLEncode(excerpt), 252)
		Set objXMLHTTP = Server.CreateObject(getXMLHTTP())
	
		objXMLHTTP.Open "POST", trackback_url, False
		objXMLHTTP.setRequestHeader "Content-Type", "application/x-www-Form-urlencoded"
	
		'HAndling timeout
		On Error Resume Next
		objXMLHTTP.Send query_string
		Err.Clear
	
		Set objXMLHTTP = Nothing
	End Function
	
	'*************************************
	'获取客户端浏览器信息
	'*************************************
	
	Public Function getBrowser(strUA)
		Dim arrInfo, strType, temp1, temp2
		strType = ""
		strUA = LCase(strUA)
		arrInfo = Array("Unkown", "Unkown")
		'浏览器判断
		If InStr(strUA, "mozilla")>0 Then arrInfo(0) = "Mozilla"
		If InStr(strUA, "icab")>0 Then arrInfo(0) = "iCab"
		If InStr(strUA, "lynx")>0 Then arrInfo(0) = "Lynx"
		If InStr(strUA, "links")>0 Then arrInfo(0) = "Links"
		If InStr(strUA, "elinks")>0 Then arrInfo(0) = "ELinks"
		If InStr(strUA, "jbrowser")>0 Then arrInfo(0) = "JBrowser"
		If InStr(strUA, "konqueror")>0 Then arrInfo(0) = "konqueror"
		If InStr(strUA, "wget")>0 Then arrInfo(0) = "wget"
		If InStr(strUA, "ask jeeves")>0 Or InStr(strUA, "teoma")>0 Then arrInfo(0) = "Ask Jeeves/Teoma"
		If InStr(strUA, "wget")>0 Then arrInfo(0) = "wget"
		If InStr(strUA, "opera")>0 Then arrInfo(0) = "opera"
	
		If InStr(strUA, "gecko")>0 Then
			strType = "[Gecko]"
			arrInfo(0) = "Mozilla"
			If InStr(strUA, "aol")>0 Then arrInfo(0) = "AOL"
			If InStr(strUA, "netscape")>0 Then arrInfo(0) = "Netscape"
			If InStr(strUA, "firefox")>0 Then arrInfo(0) = "FireFox"
			If InStr(strUA, "chimera")>0 Then arrInfo(0) = "Chimera"
			If InStr(strUA, "camino")>0 Then arrInfo(0) = "Camino"
			If InStr(strUA, "galeon")>0 Then arrInfo(0) = "Galeon"
			If InStr(strUA, "k-meleon")>0 Then arrInfo(0) = "K-Meleon"
			arrInfo(0) = arrInfo(0) + strType
		End If
	
		If InStr(strUA, "bot")>0 Or InStr(strUA, "crawl")>0 Then
			strType = "[Bot/Crawler]"
			arrInfo(0) = ""
			If InStr(strUA, "grub")>0 Then arrInfo(0) = "Grub"
			If InStr(strUA, "googlebot")>0 Then arrInfo(0) = "GoogleBot"
			If InStr(strUA, "msnbot")>0 Then arrInfo(0) = "MSN Bot"
			If InStr(strUA, "slurp")>0 Then arrInfo(0) = "Yahoo! Slurp"
			arrInfo(0) = arrInfo(0) + strType
		End If
	
		If InStr(strUA, "applewebkit")>0 Then
			strType = "[AppleWebKit]"
			arrInfo(0) = ""
			If InStr(strUA, "omniweb")>0 Then arrInfo(0) = "OmniWeb"
			If InStr(strUA, "safari")>0 Then arrInfo(0) = "Safari"
			arrInfo(0) = arrInfo(0) + strType
		End If
	
		If InStr(strUA, "msie")>0 Then
			strType = "[MSIE"
			temp1 = Mid(strUA, (InStr(strUA, "msie") + 4), 6)
			temp2 = InStr(temp1, ";")
			temp1 = Left(temp1, temp2 -1)
			strType = strType & temp1 &"]"
			arrInfo(0) = "Internet Explorer"
			If InStr(strUA, "msn")>0 Then arrInfo(0) = "MSN"
			If InStr(strUA, "aol")>0 Then arrInfo(0) = "AOL"
			If InStr(strUA, "webtv")>0 Then arrInfo(0) = "WebTV"
			If InStr(strUA, "myie2")>0 Then arrInfo(0) = "MyIE2"
			If InStr(strUA, "maxthon")>0 Then arrInfo(0) = "Maxthon"
			If InStr(strUA, "gosurf")>0 Then arrInfo(0) = "GoSurf"
			If InStr(strUA, "netcaptor")>0 Then arrInfo(0) = "NetCaptor"
			If InStr(strUA, "sleipnir")>0 Then arrInfo(0) = "Sleipnir"
			If InStr(strUA, "avant browser")>0 Then arrInfo(0) = "AvantBrowser"
			If InStr(strUA, "greenbrowser")>0 Then arrInfo(0) = "GreenBrowser"
			If InStr(strUA, "slimbrowser")>0 Then arrInfo(0) = "SlimBrowser"
			arrInfo(0) = arrInfo(0) + strType
		End If
	
		'操作系统判断
		If InStr(strUA, "windows")>0 Then arrInfo(1) = "Windows"
		If InStr(strUA, "windows ce")>0 Then arrInfo(1) = "Windows CE"
		If InStr(strUA, "windows 95")>0 Then arrInfo(1) = "Windows 95"
		If InStr(strUA, "win98")>0 Then arrInfo(1) = "Windows 98"
		If InStr(strUA, "windows 98")>0 Then arrInfo(1) = "Windows 98"
		If InStr(strUA, "windows 2000")>0 Then arrInfo(1) = "Windows 2000"
		If InStr(strUA, "windows xp")>0 Then arrInfo(1) = "Windows XP"
	
		If InStr(strUA, "windows nt")>0 Then
			arrInfo(1) = "Windows NT"
			If InStr(strUA, "windows nt 5.0")>0 Then arrInfo(1) = "Windows 2000"
			If InStr(strUA, "windows nt 5.1")>0 Then arrInfo(1) = "Windows XP"
			If InStr(strUA, "windows nt 5.2")>0 Then arrInfo(1) = "Windows 2003"
		End If
		If InStr(strUA, "x11")>0 Or InStr(strUA, "unix")>0 Then arrInfo(1) = "Unix"
		If InStr(strUA, "sunos")>0 Or InStr(strUA, "sun os")>0 Then arrInfo(1) = "SUN OS"
		If InStr(strUA, "powerpc")>0 Or InStr(strUA, "ppc")>0 Then arrInfo(1) = "PowerPC"
		If InStr(strUA, "macintosh")>0 Then arrInfo(1) = "Mac"
		If InStr(strUA, "mac osx")>0 Then arrInfo(1) = "MacOSX"
		If InStr(strUA, "freebsd")>0 Then arrInfo(1) = "FreeBSD"
		If InStr(strUA, "linux")>0 Then arrInfo(1) = "Linux"
		If InStr(strUA, "palmsource")>0 Or InStr(strUA, "palmos")>0 Then arrInfo(1) = "PalmOS"
		If InStr(strUA, "wap ")>0 Then arrInfo(1) = "WAP"
	
		'arrInfo(0)=strUA
		getBrowser = arrInfo
	End Function
	
	'*************************************
	'自动闭合UBB
	'*************************************
	
	Public Function closeUBB(strContent)
		Dim arrTags, i, OpenPos, ClosePos, re, strMatchs, j, Match
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		arrTags = Array("code", "quote", "list", "color", "align", "font", "size", "b", "i", "u", "s", "html")
		For i = 0 To UBound(arrTags)
			OpenPos = 0
			ClosePos = 0
	
			re.Pattern = "\[" + arrTags(i) + "(=[^\[\]]+|)\]"
			Set strMatchs = re.Execute(strContent)
			For Each Match in strMatchs
				OpenPos = OpenPos + 1
			Next
			re.Pattern = "\[/" + arrTags(i) + "\]"
			Set strMatchs = re.Execute(strContent)
			For Each Match in strMatchs
				ClosePos = ClosePos + 1
			Next
			For j = 1 To OpenPos - ClosePos
				strContent = strContent + "[/" + arrTags(i) + "]"
			Next
		Next
		closeUBB = strContent
	End Function
	
	'*************************************
	'自动闭合HTML
	'*************************************
	
	Public Function closeHTML(strContent)
		Dim arrTags, i, OpenPos, ClosePos, re, strMatchs, j, Match
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		arrTags = Array("p", "div", "span", "table", "ul", "font", "b", "u", "i", "h1", "h2", "h3", "h4", "h5", "h6")
		For i = 0 To UBound(arrTags)
			OpenPos = 0
			ClosePos = 0
	
			re.Pattern = "\<" + arrTags(i) + "( [^\<\>]+|)\>"
			Set strMatchs = re.Execute(strContent)
			For Each Match in strMatchs
				OpenPos = OpenPos + 1
			Next
			re.Pattern = "\</" + arrTags(i) + "\>"
			Set strMatchs = re.Execute(strContent)
			For Each Match in strMatchs
				ClosePos = ClosePos + 1
			Next
			For j = 1 To OpenPos - ClosePos
				strContent = strContent + "</" + arrTags(i) + ">"
			Next
		Next
		closeHTML = strContent
	End Function
			
End Class
%>
<script language="jscript" type="text/jscript" runat="server">
	function MultiPage(Numbers, Perpage, Curpage, Url, Style, event, FirstShortCut){
		var _curPage = parseInt(Curpage);
		var _numbers = parseInt(Numbers);
		event = event || ""
		//var Url = (baseUrl || Request.ServerVariables("URL"))+Url_Add;
		
		var Page, Offset, PageI;
		//	If Int(_numbers)>Int(PerPage) Then
		
		Page = 9;
		Offset = 4;
		
		var Pages, FromPage, ToPage;
		
		if (_numbers%Perpage == 0) {
			Pages = parseInt(_numbers / Perpage);
		}else{
			Pages = parseInt(_numbers / Perpage) + 1;
		}
		
		if (Pages < 2){
			return "";
		}
		
		FromPage = _curPage - Offset;
		ToPage = _curPage + Page - Offset -1;
		
		if (Page>Pages) {
			FromPage = 1;
			ToPage = Pages;
		}else{
			if (FromPage<1) {
				ToPage = _curPage+1 - FromPage;
				FromPage = 1;
				if ((ToPage - FromPage)<Page && (ToPage - FromPage)<Pages) {ToPage = Page}
			}else if (ToPage>Pages) {
				FromPage = _curPage - Pages + ToPage;
				ToPage = Pages;
				if ((ToPage - FromPage)<Page && (ToPage - FromPage)<Pages) {FromPage = Pages - Page+1}
			}
		}
	
		//html start
		var pageCode = ['<div class="page" style="'+Style+'"><ul><li class="pageNumber">']; // & _curPage&"/"&Pages & " | "
		
		//第一页
		if (_curPage!=1 && FromPage>1) {pageCode.push('<a href="'+Url.replace("{$page}", 1)+' title="第一页"  style="text-decoration:none" ' + event + '>«</a> | ')}
			
		if (!FirstShortCut) {ShortCut = ' accesskey=","'}else{ ShortCut = ''}
		
		//上一页
		if (_curPage!=1) {pageCode.push('<a href="'+Url.replace("{$page}", (_curPage - 1))+'" title="上一页" style="text-decoration:none;"'+ShortCut+' ' + event + '>‹</a> | ')}
		
		//列表部分
		for (PageI = FromPage;PageI<=ToPage;PageI++){
			if (PageI!=_curPage) {
				pageCode.push('<a href="'+Url.replace("{$page}", PageI) + '">'+PageI+'</a> | ');
			}else{
				pageCode.push('<strong>'+PageI+'</strong>');
				if (PageI!=Pages) {pageCode.push(' | ')}
			}
		}
		
		if (!FirstShortCut) {ShortCut = ' accesskey="."'} else {ShortCut = ''}
		
		//下一页
		if (_curPage!=Pages) {pageCode.push('<a href="'+Url.replace("{$page}", (_curPage+1))+'" title="下一页" style="text-decoration:none"'+ShortCut+' ' + event + '>›</a>')}
		
		//最后一页
		if (_curPage!=Pages && ToPage<Pages) {pageCode.push(' | <a href="'+Url.replace("{$page}", (Pages+aname))+'" title="最后一页" style="text-decoration:none" ' + event + '>»</a>')}
		
		//html end
		pageCode.push('</li></ul></div>');
		pageCode = pageCode.join("");
		FirstShortCut = true;
		
		return pageCode;
	}
</script>