<%
'=================================
' 服务器信息
'=================================
Sub c_status
		%>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
		  <tr>
			<th class="CTitle"><%=categoryTitle%></th>
		  </tr>
		  <tr>
		    <td class="CPanel">
		    <div align="left" style="padding:5px;line-height:150%">
		     <b>软件版本:</b> PJBlog3 v<%=blog_version%> - <%=DateToStr(blog_UpdateDate,"mdy")%><br/>
		     <b>服务器时间:</b> <%=DateToStr(Now(),"Y-m-d H:I A")%><br/>
		     <b>服务器物理路径:</b> <%=Request.ServerVariables("APPL_PHYSICAL_PATH")%><br/>
		     <b>服务器空间占用:</b> <%=GetTotalSize(Request.ServerVariables("APPL_PHYSICAL_PATH"),"Folder")%><br/>
		     <b>服务器CPU数量:</b> <%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%><br/>
		     <b>服务器IIS版本:</b> <%=Request.ServerVariables("SERVER_SOFTWARE")%><br/>
		     <b>脚本超时设置:</b> <%=Server.ScriptTimeout%><br/>
		     <b>脚本解释引擎:</b> <%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %><br/>
		     <b>服务器操作系统:</b> <%=Request.ServerVariables("OS")%><br/>
		     <b>服务器IP地址:</b> <%=Request.ServerVariables("LOCAL_ADDR")%><br/>
		     <b>客户端IP地址:</b> <%=Request.ServerVariables("REMOTE_ADDR")%><br/><br/>
		     
		     <b>关键组件:</b> (缺少关键组件的服务器会对PJBlog3运行有一定影响)<br/>
		     <b>　－ Scripting.FileSystemObject 组件:</b> <%=DisI(CheckObjInstalled("Scripting.FileSystemObject"))%><br/>
		     <b>　－ MSXML2.ServerXMLHTTP 组件:</b> <%=DisI(CheckObjInstalled("MSXML2.ServerXMLHTTP"))%><br/>
		     <b>　－ Microsoft.XMLDOM 组件:</b> <%=DisI(CheckObjInstalled("Microsoft.XMLDOM"))%><br/>
		     <b>　－ ADODB.Stream 组件:</b> <%=DisI(CheckObjInstalled("ADODB.Stream"))%><br/>
		     <b>　－ Scripting.Dictionary 组件:</b> <%=DisI(CheckObjInstalled("Scripting.Dictionary"))%><br/>
		     <br/>
		     
		     <b>其他组件: </b>(以下组件不影响PJBlog3运行)<br/>
		     <b>　－ Msxml2.ServerXMLHTTP.5.0 组件:</b> <%=DisI(CheckObjInstalled("Msxml2.ServerXMLHTTP.5.0"))%><br/>
		     <b>　－ Msxml2.DOMDocument.5.0 组件:</b> <%=DisI(CheckObjInstalled("Msxml2.DOMDocument.5.0"))%><br/>
		     <b>　－ FileUp.upload 组件:</b> <%=DisI(CheckObjInstalled("FileUp.upload"))%><br/>
		     <b>　－ JMail.SMTPMail 组件:</b> <%=DisI(CheckObjInstalled("JMail.SMTPMail"))%><br/>
		     <b>　－ GflAx190.GflAx 组件:</b> <%=DisI(CheckObjInstalled("GflAx190.GflAx"))%><br/>
		     <b>　－ easymail.Mailsend 组件:</b> <%=DisI(CheckObjInstalled("easymail.Mailsend"))%><br/>
		    </div>
		</td></tr></table>
<%
end Sub
%>