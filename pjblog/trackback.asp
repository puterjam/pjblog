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

dim tbID,tbURL,tbTitle,tbExcerpt,tbBlog
'Trackback entry mode
tbID = CheckStr(Request.QueryString("tbID"))
if not IsInteger(tbID) then 
	Response.contentType="text/xml"
	Response.write "<?xml version=""1.0"" encoding=""UTF-8""?><?xml-stylesheet type=""text/xsl"" href=""tb.xsl""?>"
	%><response><error>1</error><message>错误的操作</message></response><%
	call CloseDB
	response.end
end if

if Request.QueryString("action") = "view" then '查看tb
		'show Trackback List
		Set tbBlog=Conn.Execute("SELECT tb_Title,tb_URL,tb_Intro,tb_PostTime,log_title,tb_ID FROM blog_Trackback,blog_Content WHERE log_ID = blog_ID and blog_ID="&CheckStr(Request.QueryString("tbID"))&" ORDER BY tb_PostTime desc")
		SQLQueryNums=SQLQueryNums+1

		Response.contentType="text/xml"
		Response.write "<?xml version=""1.0"" encoding=""UTF-8""?><?xml-stylesheet type=""text/xsl"" href=""tb.xsl""?>"
		If Not (tbBlog.BOF and tbBlog.EOF) Then
			%>
			<response><error>0</error>
				<rss version="0.91">
					<channel>
						<title><![CDATA[<%=tbBlog(4)%>]]></title>
						<link><![CDATA[<%=SiteURL%>article.asp?id=<%=tbID%>]]></link>
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
							loop
						%>
					</channel>
				</rss>
			</response>
		   <%
		   tbBlog.close
		   set tbBlog = nothing
		else
		   %><response><error>1</error><message>日志没有被引用.</message></response><%
	    end if
elseif Request.QueryString("action") = "addtb" then '处理tb
	dim tbKey,mi
	mi = int(Minute(now()) / 10)
	tbKey = sha1(tbID & getServerKey & mi)
	
	if Request("tbKey") <> tbKey then
		tbResponseXML 1,"trackback参数不完整"
	end if 
	
	If Request("url")<>Empty Then
		tbURL = checkURL(CheckStr(Request("url")))
		tbTitle = HtmlToText(CheckStr(Request("title")))
		tbExcerpt = HtmlToText(CheckStr(Request("excerpt")))
		tbBlog = HtmlToText(CheckStr(Request("blog_name")))
		doTrackback tbID,tbURL,tbTitle,tbExcerpt,tbBlog
	End If
elseif Request.QueryString("action") = "deltb" then '删除tb
	call delTrackback
else
	Response.contentType="text/xml"
	Response.write "<?xml version=""1.0"" encoding=""UTF-8""?><?xml-stylesheet type=""text/xsl"" href=""tb.xsl""?>"
	%><response><error>1</error><message>错误的操作</message></response><%
end if
call CloseDB

'*******************************************
'Trackback response function
'*******************************************
Function tbResponseXML(intFlag, strMessage)
	Response.contentType="text/xml"
	Response.write "<?xml version=""1.0"" encoding=""UTF-8""?><response><error>"&intFlag&"</error>"
	If intFlag=1 Then Response.write "<message>"&strMessage&"</message>"
	Response.write "</response>"
	Conn.close
	set Conn=nothing
	Response.End
End Function


function doTrackback(tbID,tbURL,tbTitle,tbExcerpt,tbBlog)
    if len(tbURL)<1 or len(tbExcerpt)<1 or len(tbBlog)<1 then
	    tbResponseXML 1,"trackback参数不完整"
    end if
	tbURL=CutStr(tbURL,252)
	tbExcerpt=CutStr(tbExcerpt,252)
	tbBlog=CutStr(tbBlog,100)
	
    if tbTitle="" Then
			tbTitle=tbURL
	Else
			tbTitle=CutStr(tbTitle,100)
	End If
		'Check if allow trackback
	if Conn.Execute("SELECT count(log_ID) FROM blog_Content WHERE (log_IsShow=false or log_IsDraft=true) AND log_ID="&tbID)(0)>0 then 
			tbResponseXML 1,"隐藏日志无法发送引用通告！"
	end if
	
	If Conn.Execute("SELECT count(tb_ID) FROM blog_Trackback WHERE blog_ID="&tbID&" AND tb_URL='"&tbURL&"' AND tb_Title='"&tbTitle&"' AND tb_Intro='"&tbExcerpt&"' AND tb_Site='"&tbBlog&"'")(0)>0 Then
			tbResponseXML 1,"此引用通告已发送！"
    end if
    
   if filterSpam(tbExcerpt & tbURL,"spam.xml") then
			tbResponseXML 1,"此引用通告中包含被屏蔽的字符"
   end if
    
	Conn.Execute("INSERT INTO blog_TrackBack (blog_ID, tb_URL, tb_Title, tb_Intro, tb_Site, tb_PostTime) VALUES ("&tbID&",'"&tbURL&"','"&tbTitle&"','"&tbExcerpt&"','"&tbBlog&"',Now())")
	Conn.Execute("UPDATE blog_Content SET log_QuoteNums=log_QuoteNums+1 WHERE log_ID="&tbID)
	Conn.Execute("UPDATE blog_Info Set blog_tbNums=blog_tbNums+1")
	Smilies(1)
	Keywords(1)
	getInfo(2)
	PostArticle tbID
	SQLQueryNums=SQLQueryNums+3
	tbResponseXML 0,"Trackback 成功!"	
end function

function delTrackback
	If Not IsInteger(Request.QueryString("tbID")) AND IsInteger(Request.QueryString("logID")) Then
	  showmsg "Trackback 错误信息","<b>无效参数</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>","ErrorIcon",""
	  exit function
	end if
	
	checkCookies
	UserRight(1)
	Dim dele_tb,logID
	logID=CheckStr(Request.QueryString("logID"))
	Set dele_tb=Conn.ExeCute("SELECT * FROM blog_TrackBack WHERE blog_ID="&logID&" AND tb_ID="&CheckStr(Request.QueryString("tbID")))
	SQLQueryNums=SQLQueryNums+1
	IF dele_tb.EOF AND dele_tb.BOF Then
		   showmsg "Trackback 错误信息","<b>引用通告不存在</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>","WarningIcon",""
       Else
		   If stat_Admin Then
				Conn.ExeCute("UPDATE blog_Content SET log_QuoteNums=log_QuoteNums-1 WHERE log_ID="&dele_tb("blog_ID"))
				Conn.Execute("DELETE * FROM blog_TrackBack WHERE blog_ID="&logID&" AND tb_ID="&CheckStr(Request.QueryString("tbID")))
				Conn.Execute("UPDATE blog_Info Set blog_tbNums=blog_tbNums-1")
				SQLQueryNums=SQLQueryNums+2
				Smilies(1)
				Keywords(1)
                PostArticle logID
                getInfo(2)
				showmsg "提示信息","<b>引用通告删除成功！</b><br/><a href=""default.asp?id="&logID&""">单击返回</a>","MessageIcon",""
			Else
				showmsg "Trackback 错误信息","<b>你没有权限删除引用通告~</b><br/><a href=""default.asp?id="&logID&""">单击返回</a>","ErrorIcon",""
			End If
	End IF
end function
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