<!--#include file="const.asp" -->
<!--#include file="conn.asp" -->
<!--#include file="common/cache.asp" -->
<!--#include file="common/function.asp" -->
<!--#include file="common/ubbcode.asp" -->
<%
'测试svn

Function AddSiteURL(ByVal Str)
    If IsNull(Str) Then
        AddSiteURL = ""
        Exit Function
    End If

    Dim re
    Set re = New RegExp
    With re
        .IgnoreCase = True
        .Global = True

        .Pattern = "<img (.*?)src=""(?!(http|https)://)(.*?)"""
        Str = .Replace(Str, "<img $1src=""" & SiteURL & "$3""")

        .Pattern = "<a (.*?)href=""(?!(http|https|ftp|mms|rstp)://)(.*?)"""
        Str = .Replace(Str, "<a $1href=""" & SiteURL & "$3""")
    End With
    Set re = Nothing

    AddSiteURL = Str
End Function

'==================================
'  Feed日志订阅
'    更新时间: 2005-11-19
'==================================
'读取Blog设置信息
getInfo(1)

'写入关键字列表
Keywords(1)

'写入表情符号
Smilies(1)

Response.Charset = "UTF-8"
Response.ContentType = "text/xml"
Response.Expires = 60
Response.Write("<?xml version=""1.0"" encoding=""UTF-8""?>")
Dim cate_ID, FeedCate, FeedTitle, memName, FeedRows
cate_ID = CheckStr(Request.QueryString("cateID"))
FeedCate = False
FeedTitle = UnCheckStr(SiteName)
If IsInteger(cate_ID) = False Then
    SQL = "SELECT TOP 10 L.log_ID,L.log_Title,l.log_Author,L.log_PostTime,L.log_Content,L.log_edittype,C.cate_Name FROM blog_Content AS L,blog_Category AS C WHERE C.cate_ID=L.log_cateID AND L.log_IsShow=true AND L.log_IsDraft=false and C.cate_Secret=false ORDER BY log_PostTime DESC"
Else
    SQL = "SELECT TOP 10 L.log_ID,L.log_Title,l.log_Author,L.log_PostTime,L.log_Content,L.log_edittype,C.cate_Name FROM blog_Content AS L,blog_Category AS C WHERE log_cateID="&cate_ID&" AND C.cate_ID=L.log_cateID AND L.log_IsShow=true AND L.log_IsDraft=false and C.cate_Secret=false ORDER BY log_PostTime DESC"
    FeedCate = True
End If
Dim RS, DisIMG, i
Set RS = Conn.Execute(SQL)
If RS.EOF Or RS.BOF Then
    ReDim FeedRows(0, 0)
Else
    If FeedCate Then FeedTitle = UnCheckStr(SiteName & " - " & RS("cate_Name"))
    FeedRows = RS.getrows()
End If
RS.Close
Set RS = Nothing
Conn.Close
Set Conn = Nothing
%>
<rss version="2.0">
<channel>
<title><![CDATA[<%=FeedTitle%>]]></title>
<link><%=SiteURL%></link>
<description><![CDATA[<%=blog_Title%>]]></description>
<language>zh-cn</language>
<copyright><![CDATA[Copyright 2005 PBlog3 v2.8]]></copyright>
<webMaster><![CDATA[<%=blog_email%>(<%=blog_master%>)]]></webMaster>
<generator>PBlog2 v2.4</generator> 
<image>
	<title><%=SiteName%></title>
	<url><%=SiteURL%>images/logos.gif</url>
	<link><%=SiteURL%></link>
	<description><%=SiteName%></description>
</image>
<%
If UBound(FeedRows, 1) = 0 Then
    Response.Write("<item></item>")
Else
	dim url
    For i = 0 To UBound(FeedRows, 2)
    
    If blog_postFile = 2 Then
        url = SiteURL&caload(FeedRows(0, i))
      else
      	url = SiteURL&"article.asp?id="&FeedRows(0, i)
    end if
%>
			<item>
			<link><%=url%></link>
			<title><![CDATA[<%=FeedRows(1,i)%>]]></title>
			<author><%=blog_email%>(<%=FeedRows(2,i)%>)</author>
			<category><![CDATA[<%=FeedRows(6,i)%>]]></category>
			<pubDate><%=DateToStr(FeedRows(3,i),"w,d m y H:I:S")%></pubDate>
			<guid><%=SiteURL&"default.asp?id="&FeedRows(0,i)%></guid>
		<%
'IF RS("log_IsShow")=False Then
'	Response.Write("<description><![CDATA[这是篇隐藏日志，请到 "&SiteName&" 的首页查看!]]></description>")
'Else
If FeedRows(5, i) = 0 Then
    Response.Write("<description><![CDATA["&AddSiteURL(UnCheckStr(FeedRows(4, i)))&"]]></description>")
Else
    Response.Write("<description><![CDATA["&AddSiteURL(UBBCode(HTMLEncode(FeedRows(4, i)), 0, 0, 0, 1, 1))&"]]></description>")
End If
'End IF

%>
		</item>
		<%
Next
End If
%>
</channel>
</rss>
