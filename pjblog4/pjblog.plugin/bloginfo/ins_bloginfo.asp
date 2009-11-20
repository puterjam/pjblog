<%
' *****************************************************
' 	用户控制面板
' *****************************************************
Public Function BlogInfo
	BlogInfo = ""
	BlogInfo = BlogInfo & "日志: <a href=""index_1.html""><b>" & blog_LogNums & "</b> 篇</a><br/>"
	BlogInfo = BlogInfo & "评论: <a href=""search.asp?searchType=Comments""><b>" & blog_CommNums & "</b> 个</a><br/>"
	BlogInfo = BlogInfo & "引用: <a href=""search.asp?searchType=trackback""><b>" & blog_TbCount & "</b> 个</a><br/>"
	BlogInfo = BlogInfo & "留言: <b>" & blog_MessageNums & "</b> 个<br/>"
	BlogInfo = BlogInfo & "会员: <a href=""member.asp""><b>" & blog_MemNums & "</b> 人</a><br/>"
	BlogInfo = BlogInfo & "访问: <b>" & blog_VisitNums & "</b> 次<br/>"
	BlogInfo = BlogInfo & "在线: <b>" & Init.getOnline & "</b> 人<br/>"
	BlogInfo = BlogInfo & "建站时间: <strong>2009-10-29</strong>"
End Function
%>
	PluginTempValue = ['plugin_BlogInfo', '<%=outputSideBar(BlogInfo)%>'];
	PluginOutPutString.push(PluginTempValue);