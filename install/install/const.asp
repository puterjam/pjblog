<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'***************PJblog2 基本设置*******************
' PJblog2 Copyright 2005
' Update:2005-8-16
'**************************************************

Option Explicit
Response.Buffer = True
Server.ScriptTimeOut = 90
Session.CodePage = 65001
Session.LCID = 2057

'定义 Cookie,Application 域，必须修改，否则可能运行不正常
'把"PJBlog3"和"PJBlog3Setting"引号里面的东西替换称任意英文数值值即可
Const CookieName = "<@CookieName#>"
Const CookieNameSetting = "PJBlog3Setting"
Const IPViewURL = "http://www.dheart.net/ip/index.php?ip=" 'IP查询网站地址
Const PluginURL = "http://rp.pjhome.net/ep.asp"
Response.Cookies(CookieNameSetting).Expires = Date+365

'站点开关操作
If Not IsNumeric(Application(CookieName & "_SiteEnable")) Or IsEmpty(Application(CookieName & "_SiteEnable")) Then
    Application.Lock
    Application(CookieName & "_SiteEnable") = 1
    Application(CookieName & "_SiteDisbleWhy") = ""
    Application.UnLock
End If
If Application(CookieName & "_SiteEnable") = 0 And Application(CookieName & "_SiteDisbleWhy")<>"" And InStr(Replace(LCase(Request.ServerVariables("URL")), "\", "/"), "/control.asp") = 0 And InStr(Replace(LCase(Request.ServerVariables("URL")), "\", "/"), "/login.asp") = 0 And InStr(Replace(LCase(Request.ServerVariables("URL")), "\", "/"), "/conmenu.asp") = 0 And InStr(Replace(LCase(Request.ServerVariables("URL")), "\", "/"), "/conhead.asp") = 0 And InStr(Replace(LCase(Request.ServerVariables("URL")), "\", "/"), "/concontent.asp") = 0 Then
    Response.Write("<meta http-equiv=""Content-Type"" content=""text/html; charset=UTF-8"" /><div style='font-size:12px;font-weight:bold;border:1px solid #006;padding:6px;background:#eef'>"&Application(CookieName & "_SiteDisbleWhy")&"</div>")
    Response.End
End If

Dim StartTime, SQLQueryNums
StartTime = Timer()
SQLQueryNums = 0

'定义数据库链接文件，建议修改
'blogDB/PBLog3.asp为数据库路径，PBLog3.asp为数据库，修改这里就要修改相关的文件夹及数据库名字
Const AccessFile = "<@AccessFile#>"

'定义数据库连接
Dim Conn
Dim SQL, TempVar, siteTitle, Skins
Dim log_Year, log_Month, log_Day, SQLFiltrate, cateID
Dim viewType, Url_Add, CurPage
SQLFiltrate = "WHERE"

log_Year = CheckStr(Trim(Request.QueryString("log_Year")))
log_Month = CheckStr(Trim(Request.QueryString("log_Month")))
log_Day = CheckStr(Trim(Request.QueryString("log_Day")))

cateID = CheckStr(Trim(Request.QueryString("cateID")))
viewType = CheckStr(Trim(Request.QueryString("viewType")))
SQLFiltrate = "WHERE"
Url_Add = "?"
If IsInteger(cateID) = True Then
    SQLFiltrate = SQLFiltrate&" log_CateID="&CateID&" AND"
    Url_Add = Url_Add&"CateID="&CateID&"&amp;"
End If
If IsInteger(log_Year) = True Then
    SQLFiltrate = SQLFiltrate&" year(log_PostTime)="&log_Year&" AND"
    Url_Add = Url_Add&"log_Year="&log_Year&"&amp;"
End If
If IsInteger(log_Month) = True Then
    SQLFiltrate = SQLFiltrate&" month(log_PostTime)="&log_Month&" AND"
    Url_Add = Url_Add&"log_Month="&log_Month&"&amp;"
End If
If IsInteger(log_Day) = True Then
    SQLFiltrate = SQLFiltrate&" day(log_PostTime)="&log_Day&" AND"
    Url_Add = Url_Add&"log_Day="&log_Day&"&amp;"
End If
If CheckStr(Request.QueryString("Page"))<>Empty Then
    Curpage = CheckStr(Request.QueryString("Page"))
    If IsInteger(Curpage) = False Then 
    	Curpage = 1
    elseif Curpage<0 then
    	Curpage = 1
    end if
Else
    Curpage = 1
End If
%>
