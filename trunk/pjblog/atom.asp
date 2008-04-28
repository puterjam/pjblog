<!--#include file="const.asp" -->
<!--#include file="conn.asp" -->
<!--#include file="common/cache.asp" -->
<!--#include file="common/function.asp" -->
<!--#include file="common/ubbcode.asp" -->
<%

function AddSiteURL(ByVal Str)
  If IsNull(Str) Then
  	AddSiteURL = ""
  	Exit Function 
  End If

  Dim re
  Set re=new RegExp
  With re
    .IgnoreCase =True
    .Global=True
    
    .Pattern="<img (.*?)src=""(?!(http|https)://)(.*?)"""
    str = .replace(str,"<img $1src=""" & SiteURL & "$3""")
    
    .Pattern="<a (.*?)href=""(?!(http|https|ftp|mms|rstp)://)(.*?)"""
    str = .replace(str,"<a $1href=""" & SiteURL & "$3""")
  End With
  Set re=Nothing

  AddSiteURL=Str
End Function

'==================================
'  atom日志订阅
'    更新时间: 2005-11-19
'==================================
'读取Blog设置信息
  getInfo(1)

'写入关键字列表
  Keywords(1)

'写入表情符号
  Smilies(1)
  
Response.Charset = "UTF-8"
Response.ContentType="text/xml"
Response.Expires=60
Response.Write("<?xml version=""1.0"" encoding=""UTF-8""?>")
Dim cate_ID,FeedCate,FeedTitle,memName,FeedRows
cate_ID=CheckStr(Request.QueryString("cateID"))
FeedCate=False
FeedTitle=UnCheckStr(SiteName)
IF IsInteger(cate_ID) = False Then
	SQL="SELECT TOP 10 L.log_ID,L.log_Title,l.log_Author,L.log_PostTime,L.log_Content,L.log_edittype,C.cate_Name,C.cate_ID FROM blog_Content AS L,blog_Category AS C WHERE C.cate_ID=L.log_cateID AND L.log_IsShow=true AND L.log_IsDraft=false and C.cate_Secret=false ORDER BY log_PostTime DESC"
Else
	SQL="SELECT TOP 10 L.log_ID,L.log_Title,l.log_Author,L.log_PostTime,L.log_Content,L.log_edittype,C.cate_Name,C.cate_ID FROM blog_Content AS L,blog_Category AS C WHERE log_cateID="&cate_ID&" AND C.cate_ID=L.log_cateID AND L.log_IsShow=true AND L.log_IsDraft=false and C.cate_Secret=false ORDER BY log_PostTime DESC"
    FeedCate=True
End IF
Dim RS,DisIMG,i
Set RS=Conn.ExeCute(SQL)
if RS.EOF or RS.BOF then
	ReDim FeedRows(0,0)
 else
   if FeedCate then FeedTitle=UnCheckStr(SiteName & " - " & RS("cate_Name"))
	FeedRows=RS.getrows()
end if
RS.close
set RS=nothing
Conn.Close
Set Conn=Nothing
%>
  <feed xmlns="http://www.w3.org/2005/Atom">
  <title type="html"><![CDATA[<%=FeedTitle%>]]></title>
  <subtitle type="html"><![CDATA[<%=blog_Title%>]]></subtitle>
  <id><%=SiteURL%></id> 
  <link rel="alternate" type="text/html" href="<%=SiteURL%>" /> 
  <link rel="self" type="application/atom+xml" href="<%=SiteURL%>atom.asp" /> 
  <generator uri="http://www.pjhome.net/" version="2.4.1022">PJBlog2</generator> 
  <updated><%
  if ubound(FeedRows,1)<>0 then
    response.write DateToStr(FeedRows(3,0),"y-m-dTH:I:S")
  else
    response.write DateToStr("2005-10-23 12:00:00","y-m-dTH:I:S")
  end if
  %></updated> 
<%
if ubound(FeedRows,1)<>0 then
	for i=0 to ubound(FeedRows,2)
		%>
  <entry>
	  <title type="html"><![CDATA[<%=FeedRows(1,i)%>]]></title>
	  <author>
		 <name><%=FeedRows(2,i)%></name>
		 <uri><%=SiteURL%></uri>
		 <email><%=blog_email%></email>
	  </author>
	  <category term="" scheme="<%=SiteURL%>default.asp?cateID=<%=FeedRows(7,i)%>" label="<%=FeedRows(6,i)%>" /> 
	  <updated><%=DateToStr(FeedRows(3,i),"y-m-dTH:I:S")%></updated>
	  <published><%=DateToStr(FeedRows(3,i),"y-m-dTH:I:S")%></published>
		  <%  
		  	     IF FeedRows(5,i)=0 then
	               Response.Write("<summary type=""html""><![CDATA["&AddSiteURL(UnCheckStr(FeedRows(4,i)))&"]]></summary>")
	              else
	               Response.Write("<summary type=""html""><![CDATA["&AddSiteURL(UBBCode(HTMLEncode(FeedRows(4,i)),0,0,0,1,1))&"]]></summary>")
			   end if
		  %>
	  <link rel="alternate" type="text/html" href="<%=SiteURL%>default.asp?id=<%=FeedRows(0,i)%>" /> 
	  <id><%=SiteURL%>default.asp?id=<%=FeedRows(0,i)%></id> 
  </entry>	
		<%
	next
end if
%>
</feed>