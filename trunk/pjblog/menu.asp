<!--#include file="commond.asp" -->
<!--#include file="common/function.asp" -->
<!--#include file="common/library.asp" -->
<!--#include file="common/cache.asp" -->
<!--#include file="common/checkUser.asp" -->
<!--#include file="common/ubbcode.asp" -->
<%
'==================================
'  XML输出菜单内容
'    更新时间: 2006-1-22
'==================================
Response.ContentType = "text/xml"
Response.Expires = 60
Response.Write("<?xml version=""1.0"" encoding=""UTF-8""?>")
%>
<CateMenu>
<SiteName><%=SiteName%></SiteName>
<SiteTitle><%=blog_Title%></SiteTitle>
<SiteURL><%=SiteURL%></SiteURL>
<%
Dim MenuXML
SQL = "select * from blog_Category order by cate_Order"
Set MenuXML = conn.Execute(SQL)

Do Until MenuXML.EOF
%>
 <Menu>
  <MenuName><![CDATA[<%=MenuXML("cate_Name")%>]]></MenuName>
  <MenuIntro><![CDATA[<%=MenuXML("cate_Intro")%>]]></MenuIntro>
  <MenuType><%=MenuXML("cate_local")%></MenuType>
  <MenuUrl><%
If MenuXML("cate_OutLink") = False Then
    response.Write "default.asp?cateID="&MenuXML("cate_ID")
Else
    response.Write MenuXML("cate_URL")
End If

%></MenuUrl>
  <logNum><%=MenuXML("cate_count")%></logNum>
 </Menu>
 <%
MenuXML.movenext
Loop
Set MenuXML = Nothing
Call CloseDB

%>
</CateMenu>
