<%
'===============================================================
'  Function For PJblog2
'    更新时间: 2006-6-2
'===============================================================
'*************************************
'产生不重复的随机数 by evio
'*************************************
Function gen_key(digits) 
dim char_array(50) 
char_array(0) = "0" 
char_array(1) = "1" 
char_array(2) = "2" 
char_array(3) = "3" 
char_array(4) = "4" 
char_array(5) = "5" 
char_array(6) = "6" 
char_array(7) = "7" 
char_array(8) = "8" 
char_array(9) = "9" 
char_array(10) = "a" 
char_array(11) = "b" 
char_array(12) = "c" 
char_array(13) = "d" 
char_array(14) = "e" 
char_array(15) = "f" 
char_array(16) = "g" 
char_array(17) = "h" 
char_array(18) = "i" 
char_array(19) = "j" 
char_array(20) = "k" 
char_array(21) = "l" 
char_array(22) = "m" 
char_array(23) = "n" 
char_array(24) = "o" 
char_array(25) = "p" 
char_array(26) = "q" 
char_array(27) = "r" 
char_array(28) = "s" 
char_array(29) = "t" 
char_array(30) = "u" 
char_array(31) = "v" 
char_array(32) = "w" 
char_array(33) = "x" 
char_array(34) = "y" 
char_array(35) = "z" 
randomize 
dim output,num
do while len(output) < digits 
num = char_array(Int((35 - 0 + 1) * Rnd + 0)) 
output = output + num 
loop 
gen_key = output & year(now) & month(now) & day(now) & hour(now) & minute(now) & second(now) 
End Function 
'*************************************
'判断是否存在文件 by evio
'*************************************
Function FileExist(FilePath) 
    FileExist = False
    Dim FSO
    Set FSO = Server.CreateObject("Scripting.FileSystemObject")
    FilePath = Server.MapPath(FilePath)
    If FSO.FileExists(FilePath) Then FileExist = True
End Function
'*************************************
'创建文件夹 by evio
'*************************************
sub createfolder(catename)
    dim catefso,blogcatepath,blogcatetestpath
	set catefso = server.CreateObject("scripting.filesystemobject")
	blogcatepath = catename
	blogcatetestpath=server.MapPath(".\"&blogcatepath&"")
	if catefso.FolderExists(blogcatetestpath) Then
    else
	catefso.createfolder(blogcatetestpath) 
    end if
	set catefso=nothing
end sub
'*************************************
'自定义路径 by evio
'*************************************
Function Alias(id)
    dim cname,ccate,chtml,ccateID,cnames,ctype
	ccateID=conn.execute("select log_CateID from blog_Content where log_ID="&id)(0)
	ccate=conn.execute("select Cate_Part from blog_Category where cate_ID="&ccateID)(0)
	if ccate="" or ccate=empty or ccate=null or len(ccate)=0 then
	ccate="article/"
	else
	ccate="article/"&ccate&"/"
	end if
	cname=conn.execute("select log_cname from blog_Content where log_ID="&id)(0)
	if cname="" or cname=empty or cname=null or len(cname)=0 then
	cnames=trim(year(now())&"-"&month(now())&"-"&day(now())&"-"&id)
	else
	cnames=cname
	end if
	ctype=conn.execute("select log_ctype from blog_Content where log_ID="&id&"")(0)
	if ctype="0" then
	chtml="htm"
	elseif ctype="1" then
	chtml="html"
	else
	chtml="asp"
	end if
	chtml="."&chtml
	Alias=ccate&cnames&chtml
End Function

'*************************************
'防止外部提交
'*************************************

Function ChkPost()
    Dim server_v1, server_v2
    chkpost = False
    server_v1 = CStr(Request.ServerVariables("HTTP_REFERER"))
    server_v2 = CStr(Request.ServerVariables("SERVER_NAME"))
	if instr(server_v1, replace(replace(server_v2, "http://", ""), "www.", ""))=0 then
'	If Mid(server_v1,8,Len(server_v2))<>server_v2 then
        chkpost = False
    Else
        chkpost = True
    End If
End Function


'*************************************
'IP过滤
'*************************************

Function MatchIP(IP)
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

Function getcode()
    getcode = "<img id=""vcodeImg"" src=""about:blank"" onerror=""this.onerror=null;this.src='common/getcode.asp?s='+Math.random();"" alt=""验证码"" title=""看不清楚?换一张"" style=""margin-right:40px;cursor:pointer;width:40px;height:18px;margin-bottom:-4px;margin-top:3px;"" onclick=""src='common/getcode.asp?s='+Math.random()""/>"
End Function

'*************************************
'限制上传文件类型
'*************************************

Function IsvalidFile(File_Type)
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
'限制插件名称
'*************************************

Function IsvalidPlugins(Plugins_Name)
    Dim NoAllowNames, NoAllowName
    NoAllowNames = "user,bloginfo,calendar,comment,search,links,archive,category,contentlist"
    NoAllowName = Split(NoAllowNames, ",")
    IsvalidPlugins = True
    Dim GName
    Plugins_Name = Trim(LCase(Plugins_Name))
    For Each GName in NoAllowName
        If Plugins_Name = GName Then
            IsvalidPlugins = False
            Exit For
        End If
    Next
End Function


'*************************************
'检测是否只包含英文和数字
'*************************************

Function IsValidChars(Str)
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

Function IsvalidValue(ArrayN, Str)
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

Function IsInteger(Para)
    IsInteger = False
    If Not (IsNull(Para) Or Trim(Para) = "" Or Not IsNumeric(Para)) Then
        IsInteger = True
    End If
End Function

'*************************************
'用户名检测
'*************************************

Function IsValidUserName(byVal UserName)
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

Function IsValidEmail(Email)
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

Function highlight(byVal strContent, byRef arrayWords)
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

Function checkURL(ByVal ChkStr)
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

Function FixName(UpFileExt)
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
'过滤特殊字符
'*************************************

Function CheckStr(byVal ChkStr)
    Dim Str
    Str = ChkStr
    If IsNull(Str) Then
        CheckStr = ""
        Exit Function
    End If
    Str = Replace(Str, "&", "&amp;")
    Str = Replace(Str, "'", "&#39;")
    Str = Replace(Str, """", "&#34;")
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
    Set re = Nothing
    CheckStr = Str
End Function

'*************************************
'恢复特殊字符
'*************************************

Function UnCheckStr(ByVal Str)
    If IsNull(Str) Then
        UnCheckStr = ""
        Exit Function
    End If
    Str = Replace(Str, "&#39;", "'")
    Str = Replace(Str, "&#34;", """")
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
    Set re = Nothing
    Str = Replace(Str, "&amp;", "&")
    UnCheckStr = Str
End Function

'*************************************
'转换HTML代码
'*************************************

Function HTMLEncode(ByVal reString)
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

Function CCEncode(ByVal reString)
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

Function HTMLDecode(ByVal reString)
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

Function ClearHTML(ByVal reString)
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

Function UBBFilter(ByVal reString)
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

Function EditDeHTML(byVal Content)
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

Function DateToStr(DateTime, ShowType)
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
'分页函数
'*************************************
Dim FirstShortCut, ShortCut
FirstShortCut = False

'*************************************
'切割内容 - 按行分割
'*************************************

Function SplitLines(byVal Content, byVal ContentNums)
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
'切割内容 - 按字符分割
'*************************************

Function CutStr(byVal Str, byVal StrLen)
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

'*************************************
'Trackback Function
'*************************************

Function Trackback(trackback_url, url, title, excerpt, blog_name)
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
'删除引用标签
'*************************************

Function DelQuote(strContent)
    If IsNull(strContent) Then Exit Function
    Dim re
    Set re = New RegExp
    re.IgnoreCase = True
    re.Global = True
    re.Pattern = "\[quote\](.[^\]]*?)\[\/quote\]"
    strContent = re.Replace(strContent, "")
    re.Pattern = "\[quote=(.[^\]]*)\](.[^\]]*?)\[\/quote\]"
    strContent = re.Replace(strContent, "")
    re.Pattern = "\[reply=(.[^\]]*)\](.[^\]]*?)\[\/reply\]"
    strContent = re.Replace(strContent, "")
    Set re = Nothing
    DelQuote = strContent
End Function

'*************************************
'获取客户端IP
'*************************************

Function getIP()
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


'*************************************
'获取客户端浏览器信息
'*************************************

Function getBrowser(strUA)
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
'计算随机数
'*************************************

Function randomStr(intLength)
    Dim strSeed, seedLength, pos, Str, i
    strSeed = "abcdefghijklmnopqrstuvwxyz1234567890"
    seedLength = Len(strSeed)
    Str = ""
    Randomize
    For i = 1 To intLength
        Str = Str + Mid(strSeed, Int(seedLength * Rnd) + 1, 1)
    Next
    randomStr = Str
End Function

'*************************************
'自动闭合UBB
'*************************************

Function closeUBB(strContent)
    Dim arrTags, i, OpenPos, ClosePos, re, strMatchs, j, Match
    Set re = New RegExp
    re.IgnoreCase = True
    re.Global = True
    arrTags = Array("code", "quote", "list", "color", "align", "font", "size", "b", "i", "u", "html")
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

Function closeHTML(strContent)
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

'*************************************
'读取文件
'*************************************

Function LoadFromFile(ByVal File)
    Dim objStream
    Dim RText
    RText = Array(0, "")
    Set objStream = Server.CreateObject("ADODB.Stream")
    With objStream
        .Type = 2
        .Mode = 3
        .Open
        .Charset = "utf-8"
        .Position = objStream.Size
        On Error Resume Next
        .LoadFromFile Server.MapPath(File)
        If Err Then
            RText = Array(Err.Number, Err.Description)
            LoadFromFile = RText
            Err.Clear
            Exit Function
        End If
        RText = Array(0, .ReadText)
        .Close
    End With
    LoadFromFile = RText
    Set objStream = Nothing
End Function

'*************************************
'保存文件
'*************************************

Function SaveToFile(ByVal strBody, ByVal File)
    Dim objStream
    Dim RText
    RText = Array(0, "")
    Set objStream = Server.CreateObject("ADODB.Stream")
    With objStream
        .Type = 2
        .Open
        .Charset = "utf-8"
        .Position = objStream.Size
        .WriteText = strBody
        On Error Resume Next
        .SaveToFile Server.MapPath(File), 2
        If Err Then
            RText = Array(Err.Number, Err.Description)
            SaveToFile = RText
            Err.Clear
            Exit Function
        End If
        .Close
    End With
    RText = Array(0, "保存文件成功!")
    SaveToFile = RText
    Set objStream = Nothing
End Function

'*************************************
'数据库添加修改操作
'*************************************

Function DBQuest(table, DBArray, Action)
    Dim AddCount, TempDB, i, v
    If Action<>"insert" Or Action<>"update" Then Action = "insert"
    If Action = "insert" Then v = 2 Else v = 3
    If Not IsArray(DBArray) Then
        DBQuest = -1
        Exit Function
    Else
        Set TempDB = Server.CreateObject("ADODB.RecordSet")
        On Error Resume Next
        TempDB.Open table, Conn, 1, v
        If Err Then
            DBQuest = -2
            Exit Function
        End If
        If Action = "insert" Then TempDB.addNew
        AddCount = UBound(DBArray, 1)
        For i = 0 To AddCount
            TempDB(DBArray(i)(0)) = DBArray(i)(1)
        Next
        TempDB.update
        TempDB.Close
        Set TempDB = Nothing
        DBQuest = 0
    End If
End Function

'*************************************
'显示帮助信息
'*************************************

Sub showmsg(title, des, icon, showType)
    session(CookieName&"_ShowMsg") = True
    session(CookieName&"_title") = title
    session(CookieName&"_des") = des
    session(CookieName&"_icon") = icon
    'icon 类型
    'MessageIcon
    'ErrorIcon
    'WarningIcon
    'QuestionIcon
    If showType = "plugins" Then
        RedirectUrl("../../showmsg.asp")
    Else
        RedirectUrl("showmsg.asp")
    End If
End Sub


'*************************************
'垃圾关键字过滤
'*************************************

Function filterSpam(Str, Path)
    filterSpam = False
    Dim spamXml, spamItem
    Set spamXml = Server.CreateObject(getXMLDOM())
    spamXml.async = False
    spamXml.load(Server.MapPath(Path))
    If spamXml.parseerror.errorcode = 0 Then
        For Each spamItem in spamXml.selectNodes("//key")
            If InStr(LCase(Str), LCase(spamItem.text))<>0 Then
                filterSpam = True
                Exit Function
            End If
        Next
    End If
    Set spamXml = Nothing
End Function

Function regFilterSpam(Str, Path)
    regFilterSpam = False
    Dim spamXml, spamItem, r
    Set spamXml = Server.CreateObject(getXMLDOM())
    spamXml.async = False
    spamXml.load(Server.MapPath(Path))
    If spamXml.parseerror.errorcode = 0 Then
        For Each spamItem in spamXml.selectNodes("//key")
            'r = rgExec(Str, spamItem.getAttribute("re"), spamItem.getAttribute("times"))
            r = rgExec(str,replace(spamItem.getAttribute("re"),"\\","\"),spamItem.getAttribute("times"))
            If r>0 Then
                regFilterSpam = True
                Exit Function
            End If
        Next
    End If
    Set spamXml = Nothing
End Function

Function getServerKey
    Dim serverTime, diffDay
    If Len(Application(CookieName&"_server_Time"))>0 Then '判断是否要更新serverKey
        serverTime = Application(CookieName&"_server_Time")
        diffDay = DateDiff("h", Now, serverTime)
        If diffDay > 0 Or diffDay<0 Then updateServerKey '每个1个小时更新一次 serverKey
    Else
        updateServerKey
    End If

    Dim exc
    exc = Split(Application(CookieName&"_server_excursion"), "|")

    Dim sKey
    sKey = exc(0) & Request.ServerVariables("INSTANCE_META_PATH") & Request.ServerVariables("APPL_PHYSICAL_PATH") & Request.ServerVariables("SERVER_SOFTWARE")

    getServerKey = Mid(sha1(sKey), exc(1) + 1, 10)
End Function

Function updateServerKey
    Randomize
    Application.Lock
    Application(CookieName&"_server_Time") = Now
    Application(CookieName&"_server_excursion") = Int(Rnd * 10000000) & "|" & Int(Rnd * 26)
    Application.UnLock
End Function

Function getTempKey
    getTempKey = randomStr(20)
    session(CookieName&"tempKey") = getTempKey
End Function

%>

<script src="reg.js" Language="JScript" runat="server"></script>
<script Language="JScript" runat="server">
//*************************************
//检测系统组件是否安装
//*************************************
	function CheckObjInstalled(strClassString){
		try{
			var TmpObj = Server.CreateObject(strClassString);
			return true
		}catch(e){
			return false
		}
	}

//*************************************
//判断服务器Microsoft.XMLDOM
//*************************************
	function getXMLDOM(){
		var xmldomversions = ['Microsoft.XMLDOM','MSXML2.DOMDocument','MSXML2.DOMDocument.3.0','MSXML2.DOMDocument.4.0','MSXML2.DOMDocument.5.0'];
		for (var i=0;i<xmldomversions.length;i++){
			try{
				var sc = Server.CreateObject(xmldomversions[i]);
				sc = null;
				return xmldomversions[i];
			}catch(e){}
		}
		return false
	}

//*************************************
//判断服务器MSXML2.ServerXMLHTTP
//*************************************
	function getXMLHTTP(){
		var xmlhttpversions = ['Microsoft.XMLHTTP', 'MSXML2.XMLHTTP', 'MSXML2.XMLHTTP.3.0','MSXML2.XMLHTTP.4.0','MSXML2.XMLHTTP.5.0'];
		for (var i=0;i<xmlhttpversions.length;i++){
			try{
				var st = Server.CreateObject(xmlhttpversions[i]);
				st = null;
				return xmlhttpversions[i];
			}catch(e){}
		}
		return false
	}

//*************************************
//关闭数据库
//*************************************
	function CloseDB(){
		try{
		  	Conn.close();
			Conn = null;
		}catch(e){}
	}
	
//*************************************
//重定向函数
//*************************************
	function RedirectUrl(url){
		CloseDB();
		Response.Redirect(url);
	}
	
//*************************************
//翻页函数，改成js
//*************************************
function MultiPage(Numbers, Perpage, Curpage, Url_Add, aname, Style,baseUrl,event){
    var _curPage = parseInt(Curpage);
    var _numbers = parseInt(Numbers);
	event = event || ""
	aname = aname || ""
    var Url = (baseUrl || Request.ServerVariables("URL"))+Url_Add;
    
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
    if (_curPage!=1 && FromPage>1) {pageCode.push('<a href="'+Url+'page=1'+aname+'" page="1" title="第一页"  style="text-decoration:none" ' + event + '>«</a> | ')}
        
    if (!FirstShortCut) {ShortCut = ' accesskey=","'}else{ ShortCut = ''}
    
    //上一页
    if (_curPage!=1) {pageCode.push('<a href="'+Url+'page='+ (_curPage -1)+aname+'" page="'+(_curPage-1)+'" title="上一页" style="text-decoration:none;"'+ShortCut+' ' + event + '>‹</a> | ')}
    
    //列表部分
    for (PageI = FromPage;PageI<=ToPage;PageI++){
        if (PageI!=_curPage) {
            pageCode.push('<a href="'+Url+'page='+(PageI+aname)+'" ' + event + ' page="'+PageI+'">'+PageI+'</a> | ');
        }else{
            pageCode.push('<strong>'+PageI+'</strong>');
            if (PageI!=Pages) {pageCode.push(' | ')}
        }
    }
    
    if (!FirstShortCut) {ShortCut = ' accesskey="."'} else {ShortCut = ''}
    
    //下一页
    if (_curPage!=Pages) {pageCode.push('<a href="'+Url+'page='+(_curPage+1)+aname+'" title="下一页" page="'+(_curPage+1)+'" style="text-decoration:none"'+ShortCut+' ' + event + '>›</a>')}
    
    //最后一页
    if (_curPage!=Pages && ToPage<Pages) {pageCode.push(' | <a href="'+Url+'page='+(Pages+aname)+'" page="'+Pages+'"  title="最后一页" style="text-decoration:none" ' + event + '>»</a>')}
    
    //html end
    pageCode.push('</li></ul></div>');
	pageCode = pageCode.join("");
    FirstShortCut = true;
    
    return pageCode;
}
</script>


