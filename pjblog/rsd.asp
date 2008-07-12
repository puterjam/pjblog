<!--#include file="const.asp" -->
<!--#include file="conn.asp" -->
<!--#include file="common/cache.asp" -->
<!--#include file="common/function.asp" -->
<!--#include file="common/ubbcode.asp" -->
<%
Response.ContentType = "application/xml; charset=UTF-8" 

'读取Blog设置信息
getInfo(1)
%><?xml version="1.0" encoding="UTF-8"?>
<rsd version="1.0" xmlns="http://archipelago.phrasewise.com/rsd">
  <service>
	  <engineName>PJBlog3 v<%=blog_version%></engineName> 
	  <engineLink>http://www.pjhome.net</engineLink> 
	  <homePageLink><%=siteURL%></homePageLink> 
	  <apis>
		  <api name="MovableType" preferred="true" apiLink="<%=siteURL%>xmlrpc.asp" blogID="<%=CookieName%>" /> 
	  </apis>
  </service>
</rsd>
