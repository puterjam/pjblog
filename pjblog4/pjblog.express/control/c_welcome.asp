<%
'=================================
' 后台欢迎页面
'=================================
Public Sub c_welcome
	%>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
		  <tr>
		 	<th class="CTitle"><%=Control.categoryTitle%></th>
		  </tr>
		  <tr>
		    <td class="CPanel">
		        <div id="updateInfo" style="background:ffffe1;border:1px solid #89441f;padding:4px;display:none"></div>
		    <script>
		      var CVersion = "<%=Sys.version%>";
		      var CDate = "<%=Sys.UpdateTime%>";
		    </script>
		    <script src="http://www.pjhome.net/updateN.js"></script>
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
			 <tr>
				 <td valign="top" style="padding:5px;width:140px"><img src="../images/Control/Icon/ControlPanel.png"/></td>
				 <td valign="top">
				    <div align="left" style="padding:5px;line-height:170%;clear:both;font-size:12px">
				     <b>当前软件版本:</b> PJBlog4 v<%=Sys.version%><br/>
				     <b>软件更新日期:</b> <%=Asp.DateToStr(Sys.UpdateTime,"mdy")%><br/>
				     <b>日志数量:</b> <%=blog_LogNums%> 篇<br/>
				     <b>评论数量:</b> <%=blog_CommNums%> 个<br/>
				     <b>引用数量:</b> <%=blog_TbCount%> 个<br/>
				     <b>留言数量:</b> <%=blog_MessageNums%>  个(需要留言本插件支持)<br/>
				     <b>会员总数:</b> <%=blog_MemNums%> 人<br/>
				     <b>访问次数:</b> <%=blog_VisitNums%> 次<br/>
				    </div>    
				 </td>
			 </tr>
			</table>
		
		</td></tr></table>
	<%
end Sub
%>