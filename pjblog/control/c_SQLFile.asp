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
			<div class="SubMenu"><a href="?Fmenu=SQLFile">数据库管理</a> | <a href="?Fmenu=SQLFile&Smenu=Attachments">附件管理</a> | <a href="?Fmenu=SQLFile&Smenu=Attachment">附件信息</a></div>
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
		%>
	    <table width="100%" cellpadding="2" cellspacing="1" class="TablePanel" style=" font-size:11px">
        <tr align="center">
		  <td nowrap="nowrap" class="TDHead">&nbsp;</td>
		  <td nowrap="nowrap" class="TDHead">附件所在文章</td>
		  <td nowrap="nowrap" class="TDHead">附件名</td>
		  <td nowrap="nowrap" class="TDHead">附件大小</td>
		  <td nowrap="nowrap" class="TDHead">附件更新时间</td>
		</tr>
		<%
		For Each Arrfile in Arrfiles
		    response.Write "<tr><td><input name=""Files"" type=""checkbox"" value="""&AttPath&"/"&Arrfile&""" "&checkit(Arrfile)&"</td><td>&nbsp;<a href="""&AttPath&"/"&Arrfile&""" target=""_blank"">"&getFileIcons(getFileInfo(AttPath&"/"&Arrfile)(9))&Arrfile&"</a></td><td>&nbsp;"&getFileInfo(AttPath&"/"&Arrfile)(0)&"</td><td>&nbsp;"&DateToStr(getFileInfo(AttPath&"/"&Arrfile)(10),"Y-m-d H:I:S")&"</td></tr>"
		Next
		response.Write "</table></div>"
		%>
		    <div style="color:#f00">如果文件夹内存在文件，那么该文件夹将无法删除!
			<br/>没有被日志引用的附件将被自动选中！</div>
			  	<div class="SubButton">
		      <input type="button" value="全选" class="button" onClick="checkAll()"/>  <input type="submit" name="Submit" value="删除所选的文件或文件夹" class="button"/>
		     </div>
		     </form>
			  <%
ElseIf Request.QueryString("Smenu") = "Attachment" Then%>
    <form action="ConContent.asp" method="post" style="margin:0px" name="bdkey" id="bdkey">
       <input type="hidden" name="action" value="Attachment2"/>
	    <table width="100%" cellpadding="2" cellspacing="1" class="TablePanel" style=" font-size:11px">
	   <%
	   dim uploadDB, upload_nums, PageCount
	   Set uploadDB = Server.CreateObject("ADODB.RecordSet")
       SQL="select * from blog_Files order by id desc"
       uploadDB.Open SQL,Conn,1,1
       IF not uploadDB.EOF Then
	     uploadDB.AbsolutePage = CurPage
		 upload_nums = uploadDB.RecordCount
	   %>
	     <tr><td colspan="8" style="border-bottom:1px solid #999;"><div class="pageContent">
		 <%=MultiPage(upload_nums,10,CurPage,"?Fmenu=SQLFile&Smenu=Attachment&","","float:left")%></div></td></tr>
        <tr align="center">
		  <td nowrap="nowrap" class="TDHead">&nbsp;</td>
          <td nowrap="nowrap" class="TDHead">ID</td>
          <td nowrap="nowrap" class="TDHead">附件大小</td>
          <td nowrap="nowrap" class="TDHead">附件类型</td>
          <td nowrap="nowrap" class="TDHead">附件地址</td>
          <td nowrap="nowrap" class="TDHead">附件点击次数</td>
          <td nowrap="nowrap" class="TDHead">附件更新日期</td>
          <td nowrap="nowrap" class="TDHead">白名单Code</td>
	   </tr>
	   <%
	   	Do Until uploadDB.EOF OR PageCount = uploadDB.PageSize
          If instr(uploadDB("FilesPath"),"http://") = 0 then
          If FileExist(uploadDB("FilesPath")) = True Then
            Dim FilesSize, FilesType, Filestime
            FilesSize = "<a href="""&uploadDB("FilesPath")&""" target=""_blank"">"&getFileInfo(uploadDB("FilesPath"))(0)&"</a>"
            FilesType = getFileInfo(uploadDB("FilesPath"))(9)
            Filestime = DateToStr(getFileInfo(uploadDB("FilesPath"))(10),"Y-m-d H:I:S")
          Else
            FilesSize = "文件不存在"
            FilesType = "未知"
            Filestime = "文件不存在"
          End If
          Else
            FilesSize = "<a href="""&uploadDB("FilesPath")&""" target=""_blank"">外链文件</a>"
            FilesType = Replace(right(uploadDB("FilesPath"),4),".","")
            Filestime = "外链文件"
          End If
	   %>
	   <tr align="center">
		  <td><input name="SelectFilesID" type="checkbox" value="<%=uploadDB("id")%>"/></td>
          <td><input name="FilesID" type="hidden" value="<%=uploadDB("id")%>"/><%=uploadDB("id")%></td>
          <td align="left">&nbsp;<%=FilesSize%></td>
          <td align="left">&nbsp;<%=getFileIcons(FilesType)%><%=FilesType%></td>
	      <td><input name="url" type="text" size="60" class="text" value="<%=uploadDB("FilesPath")%>" style="width:100%"/></td>
	      <td><input name="count" type="text" size="10" value="<%=uploadDB("FilesCounts")%>" class="text"/></td>
	      <td><%=Filestime%></td>
	      <td><%=right(md5(right(Ucase(uploadDB("FilesPath")),15)),10)%></td>
	   </tr>
	   <%
     	uploadDB.MoveNext
	    PageCount=PageCount+1
	loop
    uploadDB.Close
    Set uploadDB=Nothing
End If
%>
	    <tr align="center" bgcolor="#D5DAE0">
        <td colspan="8" class="TDHead" align="left" style="border-top:1px solid #999"><a name="AddLink"></a><img src="images/add.gif" style="margin:0px 2px -3px 2px"/>添加附件</td>
       </tr>
<%
Dim Rs, S_ID
Set Rs = Conn.Execute("Select top 1 id From blog_Files order by id desc")
If not rs.eof then S_ID = Rs(0) + 1 Else S_ID=1
Rs.Close : Set Rs = Nothing
%>
        <tr align="center">
		  <td></td>
          <td><input name="FilesID" type="hidden" value="-1"/><%=S_ID%></td>
          <td align="left"></td>
          <td></td>
	      <td><input name="url" type="text" size="60" class="text" value="" style="width:100%"/></td>
	      <td><input name="count" type="text" size="10" value="0" class="text"/></td>
	      <td></td>
	      <td></td>
	   </tr>
	  </table>
  <div class="SubButton" style="text-align:left;">
    	 <select name="S_Action">
			 <option value="SaveAll">保存所有附件</option>
			 <option value="DelSelect">删除所选附件</option>
		 </select>
      <input type="button" value="全选" class="button" onClick="checkAll()" style="margin-bottom:0px"/>
	  <input type="submit" name="Submit" value="提交" class="button" style="margin-bottom:0px"/>
	  <font color="red">白名单Code:Code的值是唯一的。例如：download.asp?id=1</font><font color="blue">&code=F24757B14A</font> <font color="red">带蓝色部分允许盗链，不带蓝色部分防盗链</font>
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
		    Set AccessFSO = New cls_FSO
		    If AccessFSO.FileExists(AccessFile) Then
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
		        AccessFSO.CopyFile AccessFile & ".temp", AccessFile
		        AccessFSO.DeleteFile(AccessFile & ".temp")
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
		    Set AccessFSO = New cls_FSO
		    If AccessFSO.FileExists(AccessFile) Then
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
		    Set AccessFSO = New cls_FSO
		    If AccessFSO.FileExists(AccessSource) Then
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
		        If AccessFSO.DeleteFile(AccessFile) Then response.Write ("原数据库删除成功<br/>")
		        response.Write CopyFiles(Server.Mappath(AccessSource), Server.Mappath(AccessFile))
		        If AccessFSO.DeleteFile(AccessSource) Then response.Write ("数据库备份删除成功<br/>")
		        If AccessFSO.DeleteFile(AccessFile & ".TEMP") Then response.Write ("Temp备份删除成功<br/>")
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
		    Set AccessFSO = New cls_FSO
		    If AccessFSO.FileExists(AccessSource) Then
		        Response.Write "<div style='padding:4px 0px 4px 10px;border: 1px dotted #999;margin:2px;background:#ffffee'>"
		        If AccessFSO.DeleteFile(AccessSource) Then response.Write ("数据库备份删除成功<br/>")
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

Function checkit(str)
Dim Rs
On Error Resume Next
Set Rs = Conn.Execute("select id from blog_Files where Lcase(FilesPath) like '%"&str&"%'")
If Rs.eof then
   checkit = "checked/></td><td><b>没有日志使用此附件</b></td>"
Else
   Set Rs = Conn.Execute("select log_id,log_Title from blog_Content where log_Intro like '%download.asp?id="&rs(0)&"%' or log_Content like '%download.asp?id="&rs(0)&"%'")
   Dim link
     If blog_postFile = 2 Then
      	link = caload(rs(0))
     Else
      	link = "article.asp?id="&rs(0)
     End If
   checkit = "/></td><td><a target=""_blank"" href="&link&">"&rs(1)&"</a></td>"
End If
Set Rs = Nothing
If err then checkit = "checked/></td><td><b>没有日志使用此附件</b></td>"
End Function
%>