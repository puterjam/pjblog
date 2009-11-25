<%
' *****************************************************
' 	用户控制面板
' *****************************************************
Public Function BlogInfo
	Dim temps, temps1
	plus.open("sys.bloginfo")
	temps = Asp.UnCheckStr(Plus.getSingleTemplate("ins.bloginfo"))
	temps1 = ""
	temps1 = temps1 & "日志: <a href=""index_1.html""><b>" & blog_LogNums & "</b> 篇</a><br/>"
	temps1 = temps1 & "评论: <a href=""search.asp?searchType=Comments""><b>" & blog_CommNums & "</b> 个</a><br/>"
	temps1 = temps1 & "引用: <a href=""search.asp?searchType=trackback""><b>" & blog_TbCount & "</b> 个</a><br/>"
	temps1 = temps1 & "留言: <b>" & blog_MessageNums & "</b> 个<br/>"
	temps1 = temps1 & "会员: <a href=""member.asp""><b>" & blog_MemNums & "</b> 人</a><br/>"
	temps1 = temps1 & "访问: <b>" & blog_VisitNums & "</b> 次<br/>"
	temps1 = temps1 & "在线: <b>" & Init.getOnline & "</b> 人<br/>"
	temps1 = temps1 & "建站时间: <strong>2009-10-29</strong>"
	temps = Replace(temps, "<#bloginfo#>", temps1)
	BlogInfo = temps
End Function
%>
	PluginTempValue = ['ins.bloginfo', '<%=outputSideBar(BlogInfo)%>'];
	PluginOutPutString.push(PluginTempValue);