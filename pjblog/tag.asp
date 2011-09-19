<!--#include file="BlogCommon.asp" -->
<!--#include file="header.asp" -->
<!--#include file="common/ModSet.asp" -->
<!--#include file="plugins.asp" -->
<!--内容-->
<%
'==================================
'  Tags Cloud
'    更新时间: 2005-10-28
'==================================
%>
 <!--内容-->
  <div id="Tbody">
   <div id="mainContent">
     <div id="innermainContent">
       <div id="mainContent-topimg"></div>
       	 <div id="Content_ContentList" class="content-width">
               <div class="Content">
                 <div class="Content-top"><div class="ContentLeft"></div><div class="ContentRight"></div>
                   <h1 class="ContentTitle"><img src="images/image.gif" alt="" style="margin:0px 2px -3px 0px" class="CateIcon"/><b>标签云集</b></h1>
                   <h2 class="ContentAuthor">Tags Cloud</h2>
                 </div>
                 <div class="Content-body">
						<%
Dim log_Tag, log_TagItem
For Each log_TagItem IN Arr_Tags
    log_Tag = Split(log_TagItem, "||")

%>
									<a href="default.asp?tag=<%=Server.URLEncode(log_Tag(1))%>" title="共包含 <%=log_Tag(2)%> 篇日志"><span style="font-size:<%=getTagSize(log_Tag(2))%>px"><%=log_Tag(1)%>[<%=log_Tag(2)%>]</span></a>&nbsp;&nbsp;
								<%
Next

%>
			       </div>
               </div>       	 
       </div>
       <div id="mainContent-bottomimg"></div>
   </div>
   </div>
   <%Side_Module_Replace '处理系统侧栏模块信息%>
   <div id="sidebar">
	   <div id="innersidebar">
		     <div id="sidebar-topimg"><!--工具条顶部图象--></div>
			  <%=side_html%>
		     <div id="sidebar-bottomimg"></div>
	   </div>
   </div>
   <div style="clear: both;height:1px;overflow:hidden;margin-top:-1px;"></div>
  </div>
 <!--#include file="footer.asp" -->
 
 <%
Function getTagSize(c)
    Dim i
    For i = 1 To 10
        If Int(c)<i * 2.5 Then
            getTagSize = 12 + i
            Exit Function
        End If
    Next
    getTagSize = 22
End Function

%>
