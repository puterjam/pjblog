<%
'=================================
' 数据库和插件管理
'=================================
Sub c_SQLFile
%>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
		  <tr>
		    <th class="CTitle"><%=categoryTitle%></th>
		  </tr>
		  <tr>
		    <td class="CPanel">
			<div class="SubMenu"><a href="?Fmenu=SQLFile">数据库管理</a> | <a href="?Fmenu=SQLFile&Smenu=Attachments">附件管理</a></div>
		    <div align="left" style="padding:5px;"><%getMsg%>
		     <%
		If Request.QueryString("Smenu") = "Attachments" Then
		
		%>
		   <form action="ConContent.asp" method="post" style="margin:0px" onSubmit="if (confirm('是否删除选择的文件或文件夹')) {return true} else {return false}">
		   <input type="hidden" name="action" value="Attachments"/>
		   <input type="hidden" name="whatdo" value="DelFiles"/>
		     <%
		Dim AttPath, ArrFolder, Arrfile, ArrFolders, Arrfiles, arrUpFolders, arrUpFolder, TempF
		TempF = ""
		AttPath = Request.QueryString("AttPath")
		If Len(AttPath)<1 Then
		    AttPath = "attachments"
		ElseIf bc(server.mapPath(AttPath), server.mapPath("attachments")) Then
		    AttPath = "attachments"
		End If
		ArrFolders = Split(getPathList(AttPath)(0), "*")
		Arrfiles = Split(getPathList(AttPath)(1), "*")
		response.Write "<div style=""font-weight:bold;font-size:14px;margin:3px;margin-left:0px"">"&AttPath&"</div><div style=""margin:3px;margin-left:0px;"">"
		If AttPath<>"attachments" Then
		    arrUpFolders = Split(AttPath, "/")
		    For i = 0 To UBound(arrUpFolders) -1
		        arrUpFolder = arrUpFolder&TempF&arrUpFolders(i)
		        TempF = "/"
		    Next
		End If
		If Len(arrUpFolder)>0 Then
		    response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=""?Fmenu=SQLFile&Smenu=Attachments&AttPath="&arrUpFolder&"""><img border=""0"" src=""images/file/folder_up.gif"" style=""margin:4px 3px -3px 0px""/>返回上级目录</a><br>"
		End If
		For Each ArrFolder in ArrFolders
		    response.Write "<input name=""folders"" type=""checkbox"" value="""&AttPath&"/"&ArrFolder&"""/>&nbsp;<a href=""?Fmenu=SQLFile&Smenu=Attachments&AttPath="&AttPath&"/"&ArrFolder&"""><img border=""0"" src=""images/file/folder.gif"" style=""margin:4px 3px -3px 0px""/>"&ArrFolder&"</a><br>"
		Next
		For Each Arrfile in Arrfiles
		    response.Write "<input name=""Files"" type=""checkbox"" value="""&AttPath&"/"&Arrfile&"""/>&nbsp;<a href="""&AttPath&"/"&Arrfile&""" target=""_blank"">"&getFileIcons(getFileInfo(AttPath&"/"&Arrfile)(1))&Arrfile&"</a>&nbsp;&nbsp;"&getFileInfo(AttPath&"/"&Arrfile)(0)&" | "&DateToStr(getFileInfo(AttPath&"/"&Arrfile)(2),"Y-m-d H:I:S")&" | "&getFileInfo(AttPath&"/"&Arrfile)(3)&"<br>"
		Next
		response.Write "</div>"
		
		%>
		    <div style="color:#f00">如果文件夹内存在文件，那么该文件夹将无法删除!</div>
			  	<div class="SubButton">
		      <input type="button" value="全选" class="button" onClick="checkAll()"/>  <input type="submit" name="Submit" value="删除所选的文件或文件夹" class="button"/> 
		     </div>
		     </form>
			  <%else%>
			  <b>数据库路径:</b> <%=Server.MapPath(AccessFile)%><br/>
			  <b>数据库大小:</b> <span id="accessSize"><%=getFileInfo(AccessFile)(0)%></span><br/>
			  <b>数据库操作:</b> <a href="?Fmenu=SQLFile&do=Compact">压缩修复</a> | <a href="?Fmenu=SQLFile&do=Backup">备份</a><br/>
			  <%
		Dim AccessFSO, AccessEngine, AccessSource
		'-------------压缩数据库-----------------
		If Request.QueryString("do") = "Compact" Then
		    Set AccessFSO = Server.CreateObject("Scripting.FileSystemObject")
		    If AccessFSO.FileExists(Server.Mappath(AccessFile)) Then
		        Response.Write "<div style='padding:4px 0px 4px 10px;border: 1px dotted #999;margin:2px;background:#ffffee'>"
		        Response.Write "压缩数据库开始，网站暂停一切用户的前台操作...<br/>"
		        Response.Write "关闭数据库操作...<br/>"
		        Call CloseDB
		        Application.Lock
		        FreeApplicationMemory
		        Application(CookieName & "_SiteEnable") = 0
		        Application(CookieName & "_SiteDisbleWhy") = "网站暂停中，请稍候几分钟后再来..."
		        Application.UnLock
		        Set AccessEngine = CreateObject("JRO.JetEngine")
		        AccessEngine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath(AccessFile), "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath(AccessFile & ".temp")
		        AccessFSO.CopyFile Server.Mappath(AccessFile & ".temp"), Server.Mappath(AccessFile)
		        AccessFSO.DeleteFile(Server.Mappath(AccessFile & ".temp"))
		        Set AccessFSO = Nothing
		        Set AccessEngine = Nothing
		        Application.Lock
		        Application(CookieName & "_SiteEnable") = 1
		        Application(CookieName & "_SiteDisbleWhy") = ""
		        Application.UnLock
		        Response.Write "压缩数据库完成...<br/>"
		        Application.Lock
		        Application(CookieName & "_SiteEnable") = 1
		        Application(CookieName & "_SiteDisbleWhy") = ""
		        Application.UnLock
		        Response.Write "网站恢复正常访问..."
		        Response.Write "</div>"
		        Response.Write "<script>document.getElementById('accessSize').innerText='"&getFileInfo(AccessFile)(0)&"'</script>"
		    End If
		End If
		'-------------备份数据库数据库-----------------
		If Request.QueryString("do") = "Backup" Then
		    Set AccessFSO = Server.CreateObject("Scripting.FileSystemObject")
		    If AccessFSO.FileExists(Server.Mappath(AccessFile)) Then
		        Response.Write "<div style='padding:4px 0px 4px 10px;border: 1px dotted #999;margin:2px;background:#ffffee'>"
		        Response.Write "备份数据库开始，网站暂停一切用户的前台操作...<br/>"
		        Response.Write "关闭数据库操作...<br/>"
		        Call CloseDB
		        Application.Lock
		        FreeApplicationMemory
		        Application(CookieName & "_SiteEnable") = 0
		        Application(CookieName & "_SiteDisbleWhy") = "网站暂停中，请稍候几分钟后再来..."
		        Application.UnLock
		        CopyFiles Server.Mappath(AccessFile), Server.Mappath("backup/Backup_" & DateToStr(Now(), "YmdHIS") & "_" & randomStr(8) &".mbk")
		        Application.Lock
		        Application(CookieName & "_SiteEnable") = 1
		        Application(CookieName & "_SiteDisbleWhy") = ""
		        Application.UnLock
		        Response.Write "压缩数据库完成...<br/>"
		        Application.Lock
		        Application(CookieName & "_SiteEnable") = 1
		        Application(CookieName & "_SiteDisbleWhy") = ""
		        Application.UnLock
		        Response.Write "网站恢复正常访问..."
		        Response.Write "</div>"
		        Response.Write "<script>document.getElementById('accessSize').innerText='"&getFileInfo(AccessFile)(0)&"'</script>"
		    End If
		    Set AccessFSO = Nothing
		End If
		'---------------还原数据库------------
		If Request.QueryString("do") = "Restore" Then
		    AccessSource = Request.QueryString("source")
		    Set AccessFSO = Server.CreateObject("Scripting.FileSystemObject")
		    If AccessFSO.FileExists(Server.Mappath(AccessSource)) Then
		        Response.Write "<div style='padding:4px 0px 4px 10px;border: 1px dotted #999;margin:2px;background:#ffffee'>"
		        Response.Write "还原数据库开始，网站暂停一切用户的前台操作...<br/>"
		        Response.Write "关闭数据库操作...<br/>"
		        Call CloseDB
		        Application.Lock
		        FreeApplicationMemory
		        Application(CookieName & "_SiteEnable") = 0
		        Application(CookieName & "_SiteDisbleWhy") = "网站暂停中，请稍候几分钟后再来..."
		        Application.UnLock
		        CopyFiles Server.Mappath(AccessFile), Server.Mappath(AccessFile & ".TEMP")
		        If DeleteFiles(Server.Mappath(AccessFile)) Then response.Write ("原数据库删除成功<br/>")
		        response.Write CopyFiles(Server.Mappath(AccessSource), Server.Mappath(AccessFile))
		        If DeleteFiles(Server.MapPath(AccessSource)) Then response.Write ("数据库备份删除成功<br/>")
		        If DeleteFiles(Server.Mappath(AccessFile & ".TEMP")) Then response.Write ("Temp备份删除成功<br/>")
		        Application.Lock
		        Application(CookieName & "_SiteEnable") = 1
		        Application(CookieName & "_SiteDisbleWhy") = ""
		        Application.UnLock
		        Response.Write "数据库还原完成...<br/>"
		        Application.Lock
		        Application(CookieName & "_SiteEnable") = 1
		        Application(CookieName & "_SiteDisbleWhy") = ""
		        Application.UnLock
		        Response.Write "网站恢复正常访问..."
		        Response.Write "</div>"
		        Response.Write "<script>document.getElementById('accessSize').innerText='"&getFileInfo(AccessFile)(0)&"'</script>"
		    End If
		    Set AccessFSO = Nothing
		End If
		'---------------删除备份数据库------------
		If Request.QueryString("do") = "DelFile" Then
		    AccessSource = Request.QueryString("source")
		    Set AccessFSO = Server.CreateObject("Scripting.FileSystemObject")
		    If AccessFSO.FileExists(Server.Mappath(AccessSource)) Then
		        Response.Write "<div style='padding:4px 0px 4px 10px;border: 1px dotted #999;margin:2px;background:#ffffee'>"
		        If DeleteFiles(Server.MapPath(AccessSource)) Then response.Write ("数据库备份删除成功<br/>")
		        Response.Write "</div>"
		    End If
		    Set AccessFSO = Nothing
		End If
		
		%>
			  	  <br/><b>数据库备份列表:</b> <br/>
			  <%
		Dim BackUpFiles, BackUpFile
		BackUpFiles = Split(getPathList("backup")(1), "*")
		For Each BackUpFile in BackUpFiles
		    response.Write "<a href=""backup/"&BackUpFile&""" target=""_blank"">"&getFileIcons(getFileInfo("backup/"&BackUpFile)(1))&BackUpFile&"</a>"
		    response.Write "&nbsp;&nbsp;<a href=""?Fmenu=SQLFile&do=DelFile&source=backup/"&BackUpFile&""" title=""删除备份文件"">删除</a> | <a href=""?Fmenu=SQLFile&do=Restore&source=backup/"&BackUpFile&""" title=""还原数据库"">还原数据库</a>"
		    response.Write " | "&getFileInfo("backup/"&BackUpFile)(0)&" | "&DateToStr(getFileInfo("backup/"&BackUpFile)(2),"Y-m-d H:I:S")&"<br/>"
		Next
		
		%>
			 <%end if%>
			 </div>
		 </td>
		  </tr></table>
<%
end Sub
%>