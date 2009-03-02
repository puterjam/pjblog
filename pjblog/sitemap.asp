<!--#include file="const.asp" -->
<!--#include file="conn.asp" -->
<!--#include file="common/cache.asp" -->
<!--#include file="common/function.asp" -->
<%
'==================================
'  Google SiteMap for PJBlog
'  更新时间: 2005-12-1
'  FlowJZH@gmail.com
'  www.Aryl.net
'==================================

Sub Escape(ByRef s)
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "'", "&apos;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, "<", "&gt;")
    s = Replace(s, ">", "&lt;")
End Sub

'读取Blog设置信息
getInfo(1)

Response.Charset = "UTF-8"
Response.ContentType = "text/xml"
Response.Expires = 60

Dim cate_ID, FeedRows
cate_ID = CheckStr(Request.QueryString("cateID"))
If IsInteger(cate_ID) = False Then
    SQL = "SELECT top 10 L.log_ID,L.log_PostTime FROM blog_Content AS L,blog_Category AS C WHERE C.cate_ID=L.log_cateID AND L.log_IsDraft=false and C.cate_Secret=false ORDER BY log_PostTime DESC"
Else
    SQL = "SELECT top 10 L.log_ID,L.log_PostTime FROM blog_Content AS L,blog_Category AS C WHERE log_cateID="&cate_ID&" AND C.cate_ID=L.log_cateID AND L.log_IsDraft=false and C.cate_Secret=false ORDER BY log_PostTime DESC"
End If

Dim RS, i
Set RS = Conn.Execute(SQL)
If RS.EOF Then
    ReDim FeedRows(0, 0)
Else
    FeedRows = RS.getrows()
End If
RS.Close
Set RS = Nothing
Call CloseDB
%><?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.google.com/schemas/sitemap/0.84">
  <url>
    <loc><%=SiteURL%></loc>
    <lastmod><%=ISO8601(DateAdd("h",-1,Now))%></lastmod>
    <changefreq>always</changefreq>
    <priority>0.9</priority>
  </url>
<%
If UBound(FeedRows, 1)>0 Then
    Dim iPrior, dtNow

    dtNow = Now

    With Response
        For i = 0 To UBound(FeedRows, 2)
            iPrior = 0.5
            .Write "  <url>"
            .Write vbCrLf
            .Write "    <loc>"
            
             If blog_postFile = 2 Then 
             	.Write SiteURL&caload(FeedRows(0, i))
            else
          	 	.Write SiteURL&"article.asp?id="&FeedRows(0, i)
            end if
            
            .Write "</loc>"
            .Write vbCrLf
            .Write "    <lastmod>"
            .Write ISO8601(FeedRows(1, i))
            .Write "</lastmod>"
            .Write vbCrLf
            .Write "    <changefreq>"
            If DateDiff("h", FeedRows(1, i), dtNow) < 24 Then
                .Write "hourly"
                iPrior = 0.8
            ElseIf DateDiff("d", FeedRows(1, i), dtNow) < 7 Then
                .Write "daily"
                iPrior = 0.7
            ElseIf DateDiff("ww", FeedRows(1, i), dtNow) < 4 Then
                .Write "weekly"
            ElseIf DateDiff("m", FeedRows(1, i), dtNow) < 12 Then
                .Write "monthly"
            Else
                .Write "yearly"
            End If
            .Write "</changefreq>"
            .Write vbCrLf
            If iPrior <> 0.5 Then
                .Write "    <priority>"
                .Write iPrior
                .Write "</priority>"
                .Write vbCrLf
            End If
            .Write "  </url>"
            .Write vbCrLf
        Next
    End With
End If

Function ISO8601(DateTime)
    Dim DateMonth, DateDay, DateHour, DateMinute, DateWeek, DateSecond

    DateTime = DateAdd("h", -8, DateTime)
    DateMonth = Month(DateTime)
    DateDay = Day(DateTime)
    DateHour = Hour(DateTime)
    DateMinute = Minute(DateTime)
    DateWeek = Weekday(DateTime)
    DateSecond = Second(DateTime)
    If Len(DateMonth)<2 Then DateMonth = "0"&DateMonth
    If Len(DateDay)<2 Then DateDay = "0"&DateDay
    If Len(DateMinute)<2 Then DateMinute = "0"&DateMinute
    If Len(DateHour)<2 Then DateHour = "0"&DateHour
    If Len(DateSecond)<2 Then DateSecond = "0"&DateSecond
    ISO8601 = Year(DateTime)&"-"&DateMonth&"-"&DateDay&"T"&DateHour&":"&DateMinute&":"&DateSecond&"Z"
End Function
%>
</urlset>
