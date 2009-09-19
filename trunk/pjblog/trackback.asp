<!--#include file="const.asp" -->
<!--#include file="conn.asp" -->
<!--#include file="common/function.asp" -->
<!--#include file="common/library.asp" -->
<!--#include file="common/cache.asp" -->
<!--#include file="common/checkUser.asp" -->
<!--#include file="class/cls_logAction.asp" -->
<!--#include file="common/ubbcode.asp" -->
<!--#include file="common/sha1.asp" -->
<%
'==================================
'  引用通告处理页面
'    更新时间: 2006-6-1
'==================================
getInfo(1)

Dim tbID, tbURL, tbTitle, tbExcerpt, tbBlog
'Trackback entry mode
tbID = CheckStr(Request.QueryString("tbID"))
If Not IsInteger(tbID) Then
    Response.contentType = "text/xml"
    Response.Write "<?xml version=""1.0"" encoding=""UTF-8""?><?xml-stylesheet type=""text/xsl"" href=""tb.xsl""?>"

%><response><error>1</error><message>错误的操作</message></response><%
Call CloseDB
response.End
End If

If Request.QueryString("action") = "view" Then '查看tb
    'show Trackback List
    Set tbBlog = Conn.Execute("SELECT tb_Title,tb_URL,tb_Intro,tb_PostTime,log_title,tb_ID FROM blog_Trackback,blog_Content WHERE log_ID = blog_ID and blog_ID="&CheckStr(Request.QueryString("tbID"))&" ORDER BY tb_PostTime desc")
    SQLQueryNums = SQLQueryNums + 1

    Response.contentType = "text/xml"
    Response.Write "<?xml version=""1.0"" encoding=""UTF-8""?><?xml-stylesheet type=""text/xsl"" href=""tb.xsl""?>"
    If Not (tbBlog.BOF And tbBlog.EOF) Then

%>
			<response><error>0</error>
				<rss version="0.91">
					<channel>
						<title><![CDATA[<%=tbBlog(4)%>]]></title>
						<link><![CDATA[<%=SiteURL%>default.asp?id=<%=tbID%>]]></link>
						<description><![CDATA[<%=SiteName%>]]></description>
						<language>utf-8</language>
					 	<% do while not tbBlog.eof %>
									<item>
										<title><![CDATA[<%=tbBlog(0)%>]]></title>
										<link><![CDATA[<%=tbBlog(1)%>]]></link>
										<description><![CDATA[<%=tbBlog(2)%>]]></description>
								     	<pubDate><%=DateToStr(tbBlog("tb_PostTime"),"w,d m y H:I:S")%></pubDate>
									</item>
						<%
tbBlog.MoveNext
Loop

%>
					</channel>
				</rss>
			</response>
		   <%
tbBlog.Close
Set tbBlog = Nothing
Else

%><response><error>1</error><message>日志没有被引用.</message></response><%
End If
ElseIf Request.QueryString("action") = "addtb" Then '处理tb
    Dim tbKey, mi
    mi = Int(Minute(Now()) / 10)
    tbKey = sha1(tbID & getServerKey & mi)

    If Request("tbKey") <> tbKey Then
        tbResponseXML 1, "trackback参数不完整"
    End If

    If Request("url")<>Empty Then
        tbURL = checkURL(CheckStr(Request("url")))
        tbTitle = HtmlToText(CheckStr(Request("title")))
        tbExcerpt = HtmlToText(CheckStr(Request("excerpt")))
        tbBlog = HtmlToText(CheckStr(Request("blog_name")))
        doTrackback tbID, tbURL, tbTitle, tbExcerpt, tbBlog
    End If
ElseIf Request.QueryString("action") = "deltb" Then '删除tb
    Call delTrackback
Else
    Response.contentType = "text/xml"
    Response.Write "<?xml version=""1.0"" encoding=""UTF-8""?><?xml-stylesheet type=""text/xsl"" href=""tb.xsl""?>"

%><response><error>1</error><message>错误的操作</message></response><%
End If
Call CloseDB

'*******************************************
'Trackback response function
'*******************************************

Function tbResponseXML(intFlag, strMessage)
    Response.contentType = "text/xml"
    Response.Write "<?xml version=""1.0"" encoding=""UTF-8""?><response><error>"&intFlag&"</error>"
    If intFlag = 1 Then Response.Write "<message>"&strMessage&"</message>"
    Response.Write "</response>"
    Conn.Close
    Set Conn = Nothing
    Response.End
End Function


Function doTrackback(tbID, tbURL, tbTitle, tbExcerpt, tbBlog)
    If Len(tbURL)<1 Or Len(tbExcerpt)<1 Or Len(tbBlog)<1 Then
        tbResponseXML 1, "trackback参数不完整"
    End If
    tbURL = CutStr(tbURL, 252)
    tbExcerpt = CutStr(tbExcerpt, 252)
    tbBlog = CutStr(tbBlog, 100)

    If tbTitle = "" Then
        tbTitle = tbURL
    Else
        tbTitle = CutStr(tbTitle, 100)
    End If
    'Check if allow trackback
    If Conn.Execute("SELECT count(log_ID) FROM blog_Content WHERE (log_IsShow=false or log_IsDraft=true) AND log_ID="&tbID)(0)>0 Then
        tbResponseXML 1, "隐藏日志无法发送引用通告！"
    End If

    If Conn.Execute("SELECT count(tb_ID) FROM blog_Trackback WHERE blog_ID="&tbID&" AND tb_URL='"&tbURL&"' AND tb_Title='"&tbTitle&"' AND tb_Intro='"&tbExcerpt&"' AND tb_Site='"&tbBlog&"'")(0)>0 Then
        tbResponseXML 1, "此引用通告已发送！"
    End If

    If filterSpam(tbExcerpt & tbURL, "spam.xml") Then
        tbResponseXML 1, "此引用通告中包含被屏蔽的字符"
    End If

    Conn.Execute("INSERT INTO blog_TrackBack (blog_ID, tb_URL, tb_Title, tb_Intro, tb_Site, tb_PostTime) VALUES ("&tbID&",'"&tbURL&"','"&tbTitle&"','"&tbExcerpt&"','"&tbBlog&"',Now())")
    Conn.Execute("UPDATE blog_Content SET log_QuoteNums=log_QuoteNums+1 WHERE log_ID="&tbID)
    Conn.Execute("UPDATE blog_Info Set blog_tbNums=blog_tbNums+1")
    Smilies(1)
    Keywords(1)
    getInfo(2)
    PostArticle tbID, False
    SQLQueryNums = SQLQueryNums + 3
    tbResponseXML 0, "Trackback 成功!"
End Function

Function delTrackback
    If Not IsInteger(Request.QueryString("tbID")) And IsInteger(Request.QueryString("logID")) Then
        showmsg "Trackback 错误信息", "<b>无效参数</b><br/><a href=""javascript:history.go(-1);"">" & lang.Tip.SysTem(2) & "</a>", "ErrorIcon", ""
        Exit Function
    End If

    checkCookies
    UserRight(1)
    Dim dele_tb, logID
    logID = CheckStr(Request.QueryString("logID"))
    Set dele_tb = Conn.Execute("SELECT * FROM blog_TrackBack WHERE blog_ID="&logID&" AND tb_ID="&CheckStr(Request.QueryString("tbID")))
    SQLQueryNums = SQLQueryNums + 1
    If dele_tb.EOF And dele_tb.BOF Then
        showmsg "Trackback 错误信息", "<b>引用通告不存在</b><br/><a href=""javascript:history.go(-1);"">" & lang.Tip.SysTem(2) & "</a>", "WarningIcon", ""
    Else
        If stat_Admin Then
            Conn.Execute("UPDATE blog_Content SET log_QuoteNums=log_QuoteNums-1 WHERE log_ID="&dele_tb("blog_ID"))
            Conn.Execute("DELETE * FROM blog_TrackBack WHERE blog_ID="&logID&" AND tb_ID="&CheckStr(Request.QueryString("tbID")))
            Conn.Execute("UPDATE blog_Info Set blog_tbNums=blog_tbNums-1")
            SQLQueryNums = SQLQueryNums + 2
            Smilies(1)
            Keywords(1)
            PostArticle logID, False
            getInfo(2)
            showmsg "提示信息", "<b>引用通告删除成功！</b><br/><a href=""default.asp?id="&logID&""">" & lang.Tip.SysTem(2) & "</a>", "MessageIcon", ""
        Else
            showmsg "Trackback 错误信息", "<b>你没有权限删除引用通告~</b><br/><a href=""default.asp?id="&logID&""">" & lang.Tip.SysTem(2) & "</a>", "ErrorIcon", ""
        End If
    End If
End Function
%>

<script runat="server" Language="javascript">
	function HtmlToText(str) {
	    //filter HTMLToUBBAndText
		str = str.replace(/\r/g,"");
		str = str.replace(/on(load|click|dbclick|mouseover|mousedown|mouseup)="[^"]+"/ig,"");
		//str = str.replace(/<a[^>]+href="([^"]+)"[^>]*>(.*?)<\/a>/ig,"[url=$1]$2[/url]");
		str = str.replace(/<br\/>/ig,"\n");
		str = str.replace(/<[^>]*?>/g,"");
		str = str.replace(/&/g,"&amp;");
		str = str.replace(/&amp;(.*?[^\s]);/g,"&$1;");
		str = str.replace(/\n+/g,"<br/>");
		//str=str.replace(/(\[url=(.[^\[]*)\])(.*?)(\[\/url\])/ig,"<i><u title=\"$2\">$3</u></i>");
		return str;
	}
</script>
