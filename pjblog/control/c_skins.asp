<%
'=================================
' 风格和插件
'=================================
Sub c_skins
	    Dim bmID, bMInfo

	    TypeArray = Array("sidebar", "content", "function")
		%>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
		  <tr>
		    <th class="CTitle"><%=categoryTitle%></th>
		  </tr>
		  <tr>
		    <td class="CPanel">
			<div class="SubMenu2">
			<b>风格设置: </b> <a href="?Fmenu=Skins">设置外观</a> <b>模板: </b> <a href="?Fmenu=CodeEditor&file=Template\Static.htm&referer=skins">编辑静态页面模板</a>  <br/> 
			<b>插件相关: </b> <a href="?Fmenu=Skins&Smenu=module">设置模块</a> | <a href="?Fmenu=Skins&Smenu=Plugins">已装插件管理</a> | <a href="?Fmenu=Skins&Smenu=PluginsInstall">安装模块插件</a> | <a href="?Fmenu=Skins&Smenu=PluginsOnline">在线安装插件</a></div><div id="GetHttpPage">
		
		<%
		If Request.QueryString("Smenu") = "module" Then
		
		%>
		<script>
			var lastFlag = 7;
			function filterModuleList(flag){
				var items = document.getElementsByTagName("tr");
				for (k in items) {
					if (items[k].name!="moduleItem") {
						continue;
					}
					items[k].style.display = "none";
					
					if (items[k].mType&flag || flag == 7){ //作者偷懒中...hardcode代码...
						items[k].style.display = "";
					}
				}
				
				document.getElementById("a" + lastFlag).style.fontWeight  = "normal";
				document.getElementById("a" + flag).style.fontWeight  = "bold";
				lastFlag = flag;
			}
		</script>
		   <div align="left" style="padding:5px;"><%getMsg%>
		   <form action="ConContent.asp" method="post" style="margin:0px">
		       <input type="hidden" name="action" value="Skins"/>
		       <input type="hidden" name="whatdo" value="UpdateModule"/>
		       <input type="hidden" name="DoID" value=""/>
		       
		       <div style="width:685px;">
			     	<div style="float:right;margin-top:15px;">
			     	<b>显示过滤: </b> <a id="a7" href="#" onclick="filterModuleList(7)" style="font-weight:bold">所有</a> | <a id="a1" href="#" onclick="filterModuleList(1)">隐藏</a> | <a id="a2" href="#" onclick="filterModuleList(2)">置顶</a> | <a id="a4" href="#" onclick="filterModuleList(4)">系统</a></div>
					<div class="SubButton" style="text-align:left;padding:5px;margin:0px">
						<select id="d1" name="doModule">
							 <option>对模块进行屏蔽和置顶设置</option>
							 <option>------------------------</option>
							 <option value="dohidden">&nbsp;&nbsp;- 设置隐藏</option>
							 <option value="cancelhidden">&nbsp;&nbsp;- 取消隐藏</option>
							 <option>------------------------</option>
							 <option value="doIndex">&nbsp;&nbsp;- 设置首页独享</option>
							 <option value="cancelIndex">&nbsp;&nbsp;- 取消首页独享</option>
						 </select>
						<input type="submit" name="Submit" value="保存" class="button" style="margin-bottom:0px"/> 
			     	</div>         
		     </div>  
		     	
		      <table border="0" cellpadding="2" cellspacing="1" class="TablePanel">
		        <tr align="center">
		          <td width="18" class="TDHead"><nobr>&nbsp;</nobr></td>
		          <td width="18" class="TDHead">&nbsp;</td>
		          <td width="18" class="TDHead">&nbsp;</td>
		          <td width="18" class="TDHead"><nobr>&nbsp;</nobr></td>
		          <td width="58" class="TDHead"><nobr>类型</nobr></td>
				  <td width="82" class="TDHead"><nobr>模块标识</nobr></td>
				  <td width="160" class="TDHead"><nobr>模块名称</nobr></td>
		          <td width="42" class="TDHead"><nobr>排序</nobr></td>
		          <td width="205"class="TDHead"><nobr>模块操作</nobr></td>
		          </tr>
		<%
		Dim blogModule,flag,moduleIcon
		Set blogModule = conn.Execute("select * from blog_module where type<>'function' order by type desc,SortID asc")
		Do Until blogModule.EOF
		flag = 0
		moduleIcon = blogModule("type")
		if blogModule("IsHidden") then flag = flag + 1
		if blogModule("IndexOnly") then flag = flag + 2
		if blogModule("IsSystem") then
			flag = flag + 4
			moduleIcon = "system"
		end if
		
		
		%>     <tr name="moduleItem" mType="<%=flag%>" id="tr_<%=blogModule("id")%>" align="center" style="background:<%if blogModule("type")="content" then response.write ("#a9c9e9")%>;">
		          <td><input type="checkbox" name="selectID" value="<%=blogModule("id")%>"/></td>
		          <td><%if blogModule("IsHidden") then response.write "<img src=""images/hidden.gif"" alt=""隐藏模块""/>" %></td>
			      <td><%if blogModule("IndexOnly") then response.write "<img src=""images/top.gif"" alt=""只在首页出现""/>" %></td>
		          <td><img name="Mimg_<%=blogModule("id")%>" src="images/<%=moduleIcon%>.gif" width="16" height="16"/></td>
		          <td align="left" style="padding-left:5px;"><input type="hidden" name="mID" value="<%=blogModule("id")%>"/>
			           <%if blogModule("IsSystem") then
			         		 response.write "<input name=""mType"" type=""hidden"" value="""&blogModule("type")&"""> &nbsp;<span style='color:#999'>系统</span> "
			            else
			            %>
				          <select name="mType" onChange="document.getElementById('tr_<%=blogModule("id")%>').style.backgroundColor=(this.value=='content')?'#a9c9e9':'';document.images['Mimg_<%=blogModule("id")%>'].src='images/'+this.value+'.gif'">
					           <option value="sidebar">侧边</option>
					           <option value="content" <%if blogModule("type")="content" then response.write ("selected=""selected""")%>>内容</option>
				          </select>
			        <%end if%>
		          </td>
		          <td><input name="mName" type="text" size="12" class="text" title="<%=blogModule("name")%>" value="<%=blogModule("name")%>" readonly="readonly" style="background:#ffe;"/></td>
		          <td><input name="mTitle" type="text" size="24" class="text" value="<%=blogModule("title")%>" <%if blogModule("name")="ContentList" then response.write ("readonly=""readonly"" style=""background:#e5e5e5;""")%>/></td>
		          <td><input name="mOrder" type="text" size="3" class="text" value="<%=blogModule("SortID")%>" <%if blogModule("name")="ContentList" then response.write ("readonly=""readonly"" style=""background:#e5e5e5;""")%>/></td>
		          <td align="left"><nobr>
		          <%if blogModule("name")<>"ContentList" then %>
		            <a href="?Fmenu=Skins&Smenu=editModule&miD=<%=blogModule("id")%>" title="可视化编辑模块"><img border="0" src="images/html.gif" style="margin:0px 2px -3px 0px"/>可视化编辑</a> <a href="?Fmenu=Skins&Smenu=editModuleNormal&miD=<%=blogModule("id")%>" title="编辑HTML源代码"><img border="0" src="images/code.gif" style="margin:0px 2px -3px 0px"/>编辑HTML</a> <%if not blogModule("IsSystem") then%><a href="javascript:DelModule(<%=blogModule("id")%>,'<%=blogModule("IsSystem")%>')" title="删除该模块"><img border="0" src="images/icon_del.gif" style="margin:0px 2px -3px 0px"/>删除</a><%end if%>
		          <%end if%>
		          </nobr></td>
		          </tr>
		      <%
		blogModule.movenext
		Loop
		
		%>
		        <tr align="center" bgcolor="#D5DAE0">
		         <td colspan="9" class="TDHead" align="left" style="border-top:1px solid #9EA9C5"><img src="images/add.gif" style="margin:0px 2px -3px 2px"/>添加新模块</td>
		        </tr>    
		        <tr align="center">
		          <td></td>
		          <td></td>
			      <td></td>
		         <td><img src="images/sidebar.gif" width="16" height="16"/></td>
		          <td><input type="hidden" name="mID" value="-1"/><select name="mType">
		           <option value="sidebar">侧边</option>
		           <option value="content">内容</option>
		          </select></td>
		          <td><input name="mName" type="text" size="12" class="text" /></td>
		          <td><input name="mTitle" type="text" size="24" class="text" /></td>
		          <td><input name="mOrder" type="text" size="3" class="text" /></td>
		          <td></td>
		          </tr>
		          </table>
				<div class="SubButton" style="text-align:left;padding:5px;margin:0px">
					<input type="submit" name="Submit" value="保存" class="button" style="margin-bottom:0px"/> 
		     </div>          
		               <div style="color:#f00">模块标识是唯一标记.一旦确定就无法修改.系统自带的模块不允许删除,内容模块只在首页有效.<br/><b>ContentList</b> 是系统自带的日志列表模块，不允许做任何修改</div>
		    </div>
		  </form>
		 <%
		'========================================================
		' 可视化编辑模块HTML代码
		'========================================================
		ElseIf Request.QueryString("Smenu") = "editModule" Then
		
		%>
		
		     <div align="center" style="padding:5px;"><%getMsg%>
		   <form action="ConContent.asp" method="post" style="margin:0px" onsubmit="return checkSystem()">
		     <%
		bmID = Request.QueryString("miD")
		If IsInteger(bmID) = True Then
		    Set bMInfo = conn.Execute("select * from blog_module where id="&bmID)
		    If bMInfo.EOF Or bMInfo.bof Then
		        session(CookieName&"_ShowMsg") = True
		        session(CookieName&"_MsgText") = "没找到符合条件的模块!"
		        RedirectUrl("ConContent.asp?Fmenu=Skins&Smenu=module")
		    Else
		
		%>
		      <input type="hidden" name="action" value="Skins"/>
		      <input type="hidden" name="whatdo" value="editModule"/>
		      <input type="hidden" name="DoID" value="<%=bmID%>"/>
		      <input type="hidden" name="DoName" value="<%=bMInfo("name")%>"/>
		      <input type="hidden" name="editType" value="fck"/>
		      <div style="font-weight:bold;font-size:14px;margin:3px;margin-left:0px;text-align:left"><img src="images/<%=bMInfo("type")%>.gif" width="16" height="16" style="margin:0px 4px -3px 0px"/>模块名称: <%=bMInfo("Title")%></div>
		      <%
		Dim sBasePath
		sBasePath = "fckeditor/"
		Dim oFCKeditor
		Set oFCKeditor = New FCKeditor
		oFCKeditor.BasePath = sBasePath
		oFCKeditor.Config("AutoDetectLanguage") = True
		oFCKeditor.Config("DefaultLanguage") = "zh-cn"
		oFCKeditor.Config("FormatSource") = True
		oFCKeditor.Config("FormatOutput") = True
		oFCKeditor.Config("EnableXHTML") = True
		oFCKeditor.Config("EnableSourceXHTML") = True
		oFCKeditor.Value = UnCheckStr(bMInfo("HtmlCode"))
		oFCKeditor.Create "HtmlCode"
		
		End If
		Else
		    session(CookieName&"_ShowMsg") = True
		    session(CookieName&"_MsgText") = "你提交了非法字符!"
		    RedirectUrl("ConContent.asp?Fmenu=Skins&Smenu=module")
		End If
		
		%>
			<div class="SubButton">
		      <input type="submit" name="Submit" value="保存模块代码" class="button" /> 
		      <input type="button" name="Submit" value="返回模块列表" class="button" onClick="location='ConContent.asp?Fmenu=Skins&Smenu=module'"/> 
		     </div>	
		   </div>
		   <script>
		   function checkSystem(){
		    <%if bMInfo("IsSystem") then%>
		     if (confirm("这个是系统内置的模块，修改不当，会使系统不正常。\n确定继续吗？")){
		      return true 
		     }
		     else
		     {return false}
		    <%end if%>
		    return true
		   }
		   </script>
		 <%
		'========================================
		' 编辑模块HTML代码
		'========================================
		ElseIf Request.QueryString("Smenu") = "editModuleNormal" Then
		%>
		     <div align="left" style="padding:5px;"><%getMsg%>
		   <form action="ConContent.asp" method="post" style="margin:0px" onsubmit="return checkSystem()">
		     <%
		bmID = Request.QueryString("Mid")
		If IsInteger(bmID) = True Then
		    Set bMInfo = conn.Execute("select * from blog_module where id="&bmID)
		    If bMInfo.EOF Or bMInfo.bof Then
		        session(CookieName&"_ShowMsg") = True
		        session(CookieName&"_MsgText") = "没找到符合条件的模块!"
		        RedirectUrl("ConContent.asp?Fmenu=Skins&Smenu=module")
		    Else
		
		%>
		      <input type="hidden" name="action" value="Skins"/>
		      <input type="hidden" name="whatdo" value="editModule"/>
		      <input type="hidden" name="DoID" value="<%=bmID%>"/>
		      <input type="hidden" name="DoName" value="<%=bMInfo("name")%>"/>
		      <input type="hidden" name="editType" value="normal"/>
		      <div style="font-weight:bold;font-size:14px;margin:3px;margin-left:0px;text-align:left"><img src="images/<%=bMInfo("type")%>.gif" width="16" height="16" style="margin:0px 4px -3px 0px"/>模块名称: <%=bMInfo("Title")%></div>
		      <textarea style="width:700px;height:200px" name="HtmlCode"><%=UnCheckStr(bMInfo("HtmlCode"))%></textarea>
		      <%
		End If
		Else
		    session(CookieName&"_ShowMsg") = True
		    session(CookieName&"_MsgText") = "你提交了非法字符!"
		    RedirectUrl("ConContent.asp?Fmenu=Skins&Smenu=module")
		End If
		
		%>
			<div class="SubButton" style="text-align:left">
		      <input type="submit" name="Submit" value="保存HTML代码" class="button" <%If Not CheckObjInstalled("ADODB.Stream") Then Response.Write("disabled")%>/>
		      <input type="button" name="Submit" value="返回模块列表" class="button" onClick="location='ConContent.asp?Fmenu=Skins&Smenu=module'"/> 
		     </div>	
		   </div>
		   <script>
		   function checkSystem(){
		    <%if bMInfo("IsSystem") then%>
		     if (confirm("这个是系统内置的模块，修改不当，会使系统不正常。\n确定继续吗？")){
		      return true 
		     }
		     else
		     {return false}
		    <%end if%>
		    return true
		   }
		   </script>
		<%
		ElseIf Request.QueryString("Smenu") = "PluginsInstall" Then
		    Bmodule = getModName
		    PluginsFolders = Split(getPathList("Plugins")(0), "*")
		%>
		    <div align="left" style="padding:5px;"><%getMsg%>
		      <table border="0" cellpadding="2" cellspacing="1" class="TablePanel">
		        <tr align="center">
		          <td width="18" class="TDHead">&nbsp;</td>
				  <td width="200" align="left" class="TDHead">名称</td>
				  <td width="82" class="TDHead">作者</td>
		          <td width="100" class="TDHead">发布日期</td>
		          <td width="200" class="TDHead">&nbsp;</td>
		          </tr>
			 <%
		If CheckObjInstalled(getXMLDOM()) Then
		    If CheckObjInstalled("Scripting.FileSystemObject") Then
		        Set PluginsXML = New PXML
		        For Each PluginsFolder in PluginsFolders
		            PluginsXML.XmlPath = "Plugins/"&PluginsFolder&"/install.xml"
		            PluginsXML.Open
		            If PluginsXML.getError = 0 Then
		                If Len(PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginName"))>0 And IsvalidPlugins(PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginName")) And (Not IsvalidValue(Bmodule, PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginName"))) And IsvalidValue(TypeArray, PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginType")) Then
		
		%>
		             <tr>
		               <td align="center"><img src="images/<%=PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginType")%>.gif" width="16" height="16"/></td>
		     		   <td align="left"><%=PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginTitle")%></td>
		     		   <td align="center"><%=PluginsXML.SelectXmlNodeText("PluginInstall/main/Author")%></td>
		               <td align="center"><%=PluginsXML.SelectXmlNodeText("PluginInstall/main/pubDate")%></td>
		               <td align="center">
		               <a href="?Fmenu=Skins&Smenu=InstallPlugins&Plugins=<%=PluginsFolder%>">安装插件</a> | 
		               <a href="?Fmenu=CodeEditor&file=plugins\<%=PluginsFolder%>\install.xml&referer=plugins">编辑插件</a> | 
		               <a href="javascript:alert('<%=PluginsXML.SelectXmlNodeText("PluginInstall/main/About")%>')">关于</a></td>
		             </tr>
		     	<%
		End If
		End If
		PluginsXML.CloseXml()
		Next
		Else
		    response.Write ("<tr><td colspan=""6"" align=""center""><div style=""background:#ffffe8;border:1px solid #95801c;padding:3px;text-align:left;"">你的系统不支持 <b>Scripting.FileSystemObject</b> 只能手动输入插件的文件夹名称</div>")
		    response.Write ("<div style=""text-align:left;padding:3px""><b>插件路径:</b> Plugins / <input id=""SPath"" type=""text"" size=""16"" class=""text"" value=""""/> <input type=""button"" value=""安装插件"" class=""button"" style=""margin-bottom:-2px"" onclick=""if (document.getElementById('SPath').value.length>0) {location='ConContent.asp?Fmenu=Skins&Smenu=InstallPlugins&Plugins='+document.getElementById('SPath').value}else{alert('请输入插件路径!')}""/></div>")
		    response.Write ("</td></tr>")
		End If
		Else
		    response.Write ("<tr><td colspan=""6"" align=""center""><div style=""background:#ffffe8;border:1px solid #95801c;padding:3px;text-align:left;"">你的系统不支持 <b>"&getXMLDOM()&"</b>，无法使用插件管理功能，请与服务商联系！</div></td></tr>")
		End If
		
		%>
		      </table>
		  <div style="color:#f00;text-align:left">
		  此处列出系统找到的合法的PJBlog3插件，安装插件前需要把插件连同其目录一起上传到Plugins文件夹内
		  <br/><b>注意:这里只列出没有安装的插件。</b>
		  <br/><%If Not CheckObjInstalled("ADODB.Stream") Then %><b>你的服务器不支持</b> <b><a href="http://www.google.com/search?hl=zh-CN&q=ADODB.Stream&btnG=Google+%E6%90%9C%E7%B4%A2&lr=" target="_blank"><b>ADODB.Stream</b></a> 组件,那将意味着大部分插件的无法正常工作</b><%End If%>
		  </div>    
		</div>
		<%
		'============================================================
		'    在线获取插件
		'============================================================
		ElseIf Request.QueryString("Smenu") = "PluginsOnline" Then
			GetOnlinePlugins
		'============================================================
		'    安装插件
		'============================================================
		ElseIf Request.QueryString("Smenu") = "InstallPlugins" Then
		    InstallPlugins
		    '============================================================
		    '    显示已经安装插件
		    '============================================================
		ElseIf Request.QueryString("Smenu") = "Plugins" Then
		    Dim Blog_Plugins
		    Set Blog_Plugins = conn.Execute("select * from blog_module where IsInstall=true order by id desc")
		
		%>
		    <div align="left" style="padding:5px;"><%getMsg%>
		      <table border="0" cellpadding="2" cellspacing="1" class="TablePanel">
		        <tr align="center">
		          <td width="18" class="TDHead">&nbsp;</td>
				  <td width="150" align="left" class="TDHead">名称</td>
		          <td width="150" class="TDHead">插件所在目录</td>
		          <td width="150" class="TDHead">安装时间</td>
		          <td width="240" class="TDHead">&nbsp;</td>
		        </tr>
		       <%do until Blog_Plugins.eof%>
		        <tr align="center">
		          <td ><img src="images/<%=Blog_Plugins("type")%>.gif" width="16" height="16"/></td>
				  <td align="left">
				  <%
					Set PluginsXML = New PXML
					PluginsXML.XmlPath = "Plugins/"&Blog_Plugins("InstallFolder")&"/install.xml"
					PluginsXML.Open
					If PluginsXML.getError = 0 Then
						response.write Blog_Plugins("title")
					Else
						response.write "<span style=""color:#999"">"&Blog_Plugins("title")&"--[安装文件丢失]</span>"
					End If
				  %>
                  </td>
		          <td align="left">Plugins/<%=Blog_Plugins("InstallFolder")%>/</td>
		          <td ><%=DateToStr(Blog_Plugins("InstallDate"),"Y-m-d H:I:S")%></td>
		          <td>
		          <%if len(Blog_Plugins("SettingXML"))>0 then %>
				          <a href="?Fmenu=Skins&Smenu=PluginsOptions&Plugins=<%=Blog_Plugins("name")%>">基本设置</a>
			          <%else%>
				          <span style="color:#999">基本设置</span>
		          <%end if%>
		           | 
		          <%if len(Blog_Plugins("ConfigPath"))>0 then %>
				          <a href="Plugins/<%=Blog_Plugins("InstallFolder")%>/<%=Blog_Plugins("ConfigPath")%>">高级设置</a>
			          <%else%>
				          <span style="color:#999">高级设置</span>
		          <%end if%>
		           | <a href="javascript:DelPlugins('<%=Blog_Plugins("InstallFolder")%>','<%=Blog_Plugins("type")%>')">反安装此插件</a></td>
		          </tr>
		           <%
		Blog_Plugins.movenext
		Loop
		Set Blog_Plugins = Nothing
		
		%>
		     </table>
		       <div style="color:#f00;text-align:left">
		       <%If Not CheckObjInstalled("ADODB.Stream") Then %>
			       <b>你的服务器不支持</b> <b><a href="http://www.google.com/search?hl=zh-CN&q=ADODB.Stream&btnG=Google+%E6%90%9C%E7%B4%A2&lr=" target="_blank"><b>ADODB.Stream</b></a> 组件,那将意味着大部分插件的无法正常工作</b>
		       <%else%>
			       <input type="button" name="button" value="修复插件" class="button" onClick="FixPlugins()"/>
		       <%End If%>
		<br/>
		 如果插件名称显示为灰色，请检查插件目录或者目录下install.xml文件是否存在；假如插件反安装不成功，请到 <b>数据库与附件-数据库管理</b> 压缩修复数据再反安装
		</div>
		  <%
		
		
		
		'============================================================
		'    反安装插件
		'============================================================
		ElseIf Request.QueryString("Smenu") = "UnInstallPlugins" Then
		    Dim UnPlugName, getCateID, DropTable, KeepTable, ModSetTemp1, getMod
		    PluginsFolder = CheckStr(Request.QueryString("Plugins"))
		    KeepTable = CBool(Request.QueryString("Keep"))
		    Set PluginsXML = New PXML
		    PluginsXML.XmlPath = "Plugins/"&PluginsFolder&"/install.xml"
		    PluginsXML.Open
		    If PluginsXML.getError = 0 Then
		        UnPlugName = PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginName")
		        Set ModSetTemp1 = New ModSet
		        ModSetTemp1.Open UnPlugName
		        ModSetTemp1.RemoveApplication
		        DropTable = PluginsXML.SelectXmlNodeText("PluginInstall/main/DropTable")
		        Set getMod = conn.Execute("select CateID from blog_module where name='"&UnPlugName&"'")
		        If getMod.EOF Then
		            session(CookieName&"_ShowMsg") = True
		            session(CookieName&"_MsgText") = "<font color=""#ff0000"">"&UnPlugName&"</font> 无法反安装,数据库没有找到相应的信息!"
		            RedirectUrl("ConContent.asp?Fmenu=Skins&Smenu=Plugins")
		        Else
		            getCateID = getMod(0)
		            If Len(getCateID)>0 Then conn.Execute("delete * from blog_Category where cate_ID="&getCateID)
		            delPlugins UnPlugName
		            If Len(DropTable)>0 And KeepTable = False Then
						If Instr(DropTable, ";") <> 0 Then
							Dim SplitDropTable, DropTableCount
							SplitDropTable = Split(DropTable, ";")
							For DropTableCount = 0 to UBound(SplitDropTable)
								conn.Execute("DROP TABLE "&SplitDropTable(DropTableCount))
							Next
						Else
							conn.Execute("DROP TABLE "&DropTable)
						End If
					End If
		            SubItemLen = Int(PluginsXML.GetXmlNodeLength("PluginInstall/SubItem/item"))
		
		            For tempI = 0 To SubItemLen -1
		                If Not PluginsXML.SelectXmlNodeText("PluginInstall/SubItem/item/PluginType") = "function" Then
		                    delPlugins UnPlugName&"SubItem"&(tempI + 1)
		                End If
		            Next
		            If Len(PluginsXML.SelectXmlNodeText("PluginInstall/main/SettingFile"))>0 Then
		                If KeepTable = False Then InstallPlugingSetting "", UnPlugName, "delete"
		            End If
		        End If
			Else
		        Set getMod = conn.Execute("select CateID from blog_module where InstallFolder='"&PluginsFolder&"'")
		        If getMod.EOF Then
		            session(CookieName&"_ShowMsg") = True
		            session(CookieName&"_MsgText") = "<font color=""#ff0000"">"&UnPlugName&"</font> 无法反安装,数据库没有找到相应的信息!"
		            RedirectUrl("ConContent.asp?Fmenu=Skins&Smenu=Plugins")
		        Else
		            getCateID = getMod(0)
		            If Len(getCateID)>0 Then conn.Execute("delete * from blog_Category where cate_ID="&getCateID)
					conn.Execute("delete * from blog_module where InstallFolder='"&PluginsFolder&"'")
		        End If
		    End If
		
		    PluginsXML.CloseXml()
		    log_module(2)
		    CategoryList(2)
		    FixPlugins(0)
			EmptyEtag
		    session(CookieName&"_ShowMsg") = True
		    session(CookieName&"_MsgText") = "<font color=""#ff0000"">"&UnPlugName&"</font> 插件反安装完成!"
		    RedirectUrl("ConContent.asp?Fmenu=Skins&Smenu=Plugins")
		
		    '============================================================
		    '    修复插件
		    '============================================================
		ElseIf Request.QueryString("Smenu") = "FixPlugins" Then
		    FixPlugins(1)
		    '============================================================
		    '    插件配置
		    '============================================================
		ElseIf Request.QueryString("Smenu") = "PluginsOptions" Then
		    Dim PluginsSetting, LoadSetXML, KeyLen, Si, LoadModSet, SelectTemp
		    Set PluginsSetting = conn.Execute("select top 1 * from blog_module where name='" + checkstr(Request.QueryString("Plugins")) + "'")
		    Set LoadSetXML = New PXML
		    Set LoadModSet = New ModSet
		    LoadModSet.Open(PluginsSetting("name"))
		    LoadSetXML.XmlPath = "Plugins/"&PluginsSetting("InstallFolder")&"/"&PluginsSetting("SettingXML")
		    LoadSetXML.Open
		    If LoadSetXML.getError = 0 Then
		        KeyLen = Int(LoadSetXML.GetXmlNodeLength("PluginOptions/Key"))
		        getMsg
		        Response.Write ("<div align=""center""><form action=""ConContent.asp"" method=""post"" style=""margin:0px"">")
		        Response.Write ("<input type=""hidden"" name=""action"" value=""Skins""/>")
		        Response.Write ("<input type=""hidden"" name=""whatdo"" value=""SavePluginsSetting""/>")
		        Response.Write ("<input type=""hidden"" name=""PluginsName"" value="""&PluginsSetting("name")&"""/>")
		        response.Write "<table border=""0"" cellpadding=""2"" cellspacing=""1"" class=""TablePanel"" style=""margin:6px"">"
		        response.Write("<tr><td colspan=""2"" align=""left"" style=""background:#e5e5e5;padding:6px""><div style=""font-weight:bold;font-size:14px;"">"&PluginsSetting("title")&" 的基本设置</div></td></tr>")
		        For tempI = 0 To KeyLen -1
		            response.Write "<tr><td align=""right"" width=""200"" valign=""top"" style=""padding-top:6px"">"&LoadSetXML.GetAttributes("PluginOptions/Key", "description", TempI)&"</td><td width=""300"">"
		            If LCase(LoadSetXML.GetAttributes("PluginOptions/Key", "type", TempI)) = "select" Then
		                response.Write "<select name="""&LoadSetXML.GetAttributes("PluginOptions/Key", "name", TempI)&""">"
		                For Si = 0 To LoadSetXML.SelectXmlNode("PluginOptions/Key", TempI).getElementsByTagName("option").Length -1
		                    If LoadModSet.getKeyValue(LoadSetXML.GetAttributes("PluginOptions/Key", "name", tempI)) = LoadSetXML.SelectXmlNode("PluginOptions/Key", TempI).getElementsByTagName("option").Item(Si).Attributes(0).Value Then SelectTemp = "selected"
		                    If LoadSetXML.SelectXmlNode("PluginOptions/Key", TempI).getElementsByTagName("option").Item(Si).Attributes.Length>0 Then
		                        Response.Write "<option "&SelectTemp&" value="""&LoadSetXML.SelectXmlNode("PluginOptions/Key", TempI).getElementsByTagName("option").Item(Si).Attributes(0).Value&""">"&LoadSetXML.SelectXmlNode("PluginOptions/Key", TempI).getElementsByTagName("option").Item(Si).text&"</option>"
		                    Else
		                        Response.Write "<option "&SelectTemp&">"&LoadSetXML.SelectXmlNode("PluginOptions/Key", TempI).getElementsByTagName("option").Item(Si).text&"</option>"
		                    End If
		                    SelectTemp = ""
		                Next
		                response.Write "</select></td></tr>"
		            ElseIf LCase(LoadSetXML.GetAttributes("PluginOptions/Key", "type", TempI)) = "textarea" Then
		                response.Write "<textarea name="""&LoadSetXML.GetAttributes("PluginOptions/Key", "name", tempI)&""" rows="""&LoadSetXML.GetAttributes("PluginOptions/Key", "rows", TempI)&""" cols="""&LoadSetXML.GetAttributes("PluginOptions/Key", "cols", TempI)&""">"&LoadModSet.getKeyValue(LoadSetXML.GetAttributes("PluginOptions/Key", "name", tempI))&"</textarea></td></tr>"
		            Else
		                response.Write "<input name="""&LoadSetXML.GetAttributes("PluginOptions/Key", "name", TempI)&""" type="""&LoadSetXML.GetAttributes("PluginOptions/Key", "type", TempI)&""" size="""&LoadSetXML.GetAttributes("PluginOptions/Key", "size", TempI)&""" value="""&LoadModSet.getKeyValue(LoadSetXML.GetAttributes("PluginOptions/Key", "name", tempI))&"""/></td></tr>"
		            End If
		        Next
		        response.Write "<tr><td colspan=""2"" align=""center""><input type=""submit"" name=""Submit"" value=""保存设置"" class=""button""/><input type=""button"" value=""放弃返回"" class=""button"" onclick=""history.go(-1)""/></td></tr>"
		        response.Write "</table>"
		        response.Write "</form></div>"
		    Else
		        response.Write "无法找到配置文件"
		    End If
		    Set LoadSetXML = Nothing
		    Set PluginsSetting = Nothing
		    '============================================================
		    '    设置外观
		    '============================================================
		Elseif Request.QueryString("Smenu")="delskin" then
		  delskin request.QueryString("SkinFolder")
		Else
		    Dim SkinFolders, SkinFolder
		    SkinFolders = Split(getPathList("skins")(0), "*")
		%>
		    <div align="left" style="padding:5px;"><%getMsg%>
		   <form action="ConContent.asp" method="post" style="margin:0px">
		       <input type="hidden" name="action" value="Skins"/>
		       <input type="hidden" name="whatdo" value="setDefaultSkin"/>
		       <input type="hidden" name="SkinName" value=""/>
		       <input type="hidden" name="SkinPath" value=""/>
		  </form>
		      <table border="0" cellpadding="2" cellspacing="1" class="TablePanel" width="800">
		        <tr>
			          <td width="700" class="TDHead" colspan="2">界面列表</td>
		        </tr>
			 <%
		If CheckObjInstalled(getXMLDOM()) And CheckObjInstalled("Scripting.FileSystemObject") Then
		    Dim SkinXML, k, SkinPreview
		    k = 2
		    Set SkinXML = New PXML
		    For Each SkinFolder in SkinFolders
		        SkinXML.XmlPath = "skins/"&SkinFolder&"/skin.xml"
		        SkinXML.Open
		        If SkinXML.getError = 0 Then
		            If k / 2 = Int(k / 2) Then response.Write "<tr>"
		            SkinPreview = "images/Control/skin.jpg"
		            If FileExist("skins/"&SkinFolder&"/Preview.jpg") Then SkinPreview = "skins/"&SkinFolder&"/Preview.jpg"
		
						%>
				 		<td width="50%" style='border-bottom:1px dotted #ccc'>
				 			<div class="<%if Lcase(blog_DefaultSkin)<>Lcase(SkinFolder) then response.write ("un")%>selectskin">
				 			<img src="<%=SkinPreview%>" alt="" border="0" class="skinimg"/>
				 			  <div class="skinDes">
				 			  <div style="height:38px;overflow:hidden"><b style="color:#004000"><%=SkinXML.SelectXmlNodeText("SkinName")%></b></div>
				 			  <span style="height:16px;overflow:hidden;cursor:default" title="设计者:<%=SkinXML.SelectXmlNodeText("SkinDesigner")%>"><B>设计者:</B> <%=SkinXML.SelectXmlNodeText("SkinDesigner")%></span><br/>
				 			  <B>发布时间:</B> <%=SkinXML.SelectXmlNodeText("pubDate")%><br/></div>
							 <%
							If LCase(blog_DefaultSkin) = LCase(SkinFolder) Then
							    'response.Write ("<div class=""cskin""><img src=""images/Control/select.gif"" alt="""" border=""0"" /></div>当前界面")
							    response.Write ("<a href=""?Fmenu=CodeEditor&file=skins\"&SkinFolder&"\layout.css&referer=skins""><img border=""0"" src=""images/icon_edit.gif"" style=""margin:0px 2px -3px 0px""/>编辑主题</a>") 
							Else
							    response.Write ("<a href=""javascript:setSkin('"&SkinFolder&"','"&SkinXML.SelectXmlNodeText("SkinName")&"')""><img border=""0"" src=""images/icon_apply.gif"" style=""margin:0px 2px -3px 0px""/>应用此主题</a>")
							    response.Write ("&nbsp;&nbsp;&nbsp;&nbsp;<a href=""?Fmenu=CodeEditor&file=skins\"&SkinFolder&"\layout.css&referer=skins""><img border=""0"" src=""images/icon_edit.gif"" style=""margin:0px 2px -3px 0px""/>编辑主题</a>")
							    '<a href="?Fmenu=CodeEditor">测试在线编辑器</a> 
							    response.Write ("&nbsp;&nbsp;&nbsp;&nbsp;<a href=""ConContent.asp?Fmenu=Skins&Smenu=delskin&SkinFolder="&server.URLEncode(SkinFolder)&""" onclick=""if (!window.confirm('确定要删除此主题吗？')){return false}""><img border=""0"" src=""images/icon_del.gif"" style=""margin:0px 2px -3px 0px""/>删除该主题</a>")
							End If
							
							%>
						  </div>
				 		</td>
					<%
						If k / 2<>Int(k / 2) Then response.Write "</tr>"
						k = k + 1
				End If
				SkinXML.CloseXml
			Next
		If k / 2<>Int(k / 2) Then response.Write "</tr>"
		
		Set SkinXML = Nothing
		Else
		    response.Write ("<tr><td colspan=""2"" align=""center""><div style=""background:#ffffe8;border:1px solid #95801c;padding:3px;text-align:left;"">你的系统不支持 <b>"&getXMLDOM()&"</b> 或 <b>Scripting.FileSystemObject</b> 只能手动输入Skin的文件夹名称</div>")
		    response.Write ("<div style=""text-align:left;padding:3px""><b>界面路径:</b> Skins / <input id=""SPath"" type=""text"" size=""16"" class=""text"" value=""" + blog_DefaultSkin + """/> <input type=""button"" value=""保存界面"" class=""button"" style=""margin-bottom:-2px"" onclick=""if (document.getElementById('SPath').value.length>0) {setSkin(document.getElementById('SPath').value,document.getElementById('SPath').value)}else{alert('请输入界面路径!')}""/></div>")
		    response.Write ("</td></tr>")
		End If
		
		%>
		      </table>
		</div>
		<%end if%></div>
		 </td>
		  </tr></table>
<%
end Sub
%>