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
Response.ContentType="text/xml"
Response.Expires=60
Response.Write("<?xml version=""1.0"" encoding=""UTF-8""?>")
%>
<CateMenu>
<SiteName><%=SiteName%></SiteName>
<SiteTitle><%=blog_Title%></SiteTitle>
<SiteURL><%=SiteURL%></SiteURL>
<%
dim MenuXML
SQL="select * from blog_Category order by cate_Order"
set MenuXML=conn.execute(SQL)

do until MenuXML.eof
%>
 <Menu>
  <MenuName><![CDATA[<%=MenuXML("cate_Name")%>]]></MenuName>
  <MenuIntro><![CDATA[<%=MenuXML("cate_Intro")%>]]></MenuIntro>
  <MenuType><%=MenuXML("cate_local")%></MenuType>
  <MenuUrl><%
   if MenuXML("cate_OutLink")=false then
    response.write "default.asp?cateID="&MenuXML("cate_ID")
   else
    response.write MenuXML("cate_URL")
   end if
  %></MenuUrl>
  <logNum><%=MenuXML("cate_count")%></logNum>
 </Menu>
 <%
 MenuXML.movenext
 loop
 set MenuXML=nothing
 call CloseDB
 %>
</CateMenu>
