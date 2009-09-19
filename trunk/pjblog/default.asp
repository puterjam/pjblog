<!--#include file="BlogCommon.asp" -->
<!--#include file="header.asp" -->
<!--#include file="common/ModSet.asp" -->
<!--#include file="plugins.asp" -->
<!--#include file="class/cls_default.asp" -->
<!--#include file="class/cls_logAction.asp" -->
<%
'==================================
'  系统首页
'    更新时间: 2006-1-15
'==================================
%>
 <!--内容-->
  <div id="Tbody">
  <div id="mainContent">
   <div id="innermainContent">
       <div id="mainContent-topimg"></div>
	   <%=content_html_Top_default%>
	   <div id="Content_ContentList" class="content-width">
	     <%ContentList%>
	   </div>
	   <%=content_html_Bottom_default%>
       <div id="mainContent-bottomimg"></div>
   </div>
   </div>
   <%Side_Module_Replace '处理系统侧栏模块信息%>
   <%=lang.Action.Rss("a","b")%>
   <div id="sidebar">
	   <div id="innersidebar">
		     <div id="sidebar-topimg"><!--工具条顶部图象--></div>
			  <%=side_html_default%>
		     <div id="sidebar-bottomimg"></div>
	   </div>
   </div>
   <div style="clear: both;height:1px;overflow:hidden;margin-top:-1px;"></div>
 </div>
  <!--#include file="footer.asp" -->
