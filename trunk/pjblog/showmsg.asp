<!--#include file="commond.asp" -->
<!--#include file="header.asp" -->
<!--内容-->
<%
'==================================
'  信息显示页面
'    更新时间: 2005-10-18
'==================================
If Not session(CookieName&"_ShowMsg") Then
    RedirectUrl("default.asp")
End If
session(CookieName&"_ShowMsg") = False

%>
 <div id="Tbody">
<br/><br/>
   <div style="text-align:center;">
    <div id="MsgContent" style="width:300px">
      <div id="MsgHead"><%=session(CookieName&"_title")%></div>
      <div id="MsgBody">
	   <div class="<%=session(CookieName&"_icon")%>"></div>
       <div class="MessageText"><%=session(CookieName&"_des")%></div>
	  </div>
	</div>
  </div><br/><br/>
 </div>
 <!--#include file="footer.asp" -->
