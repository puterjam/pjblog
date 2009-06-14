<%
'=================================
' 站点基本设置
'=================================
Sub c_ceneral
	%>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
		  <tr>
		      <th class="CTitle"><%=categoryTitle%></th>
		  </tr>
		  <tr>
		    <td class="CPanel">
		    <div class="SubMenu"><a href="?Fmenu=General">设置基本信息</a> | <a href="?Fmenu=General&Smenu=UpLoadSet">附件基本设置</a> | <a href="?Fmenu=General&Smenu=visitors">查看访客记录</a> | <a href="?Fmenu=General&Smenu=Misc">初始化数据</a> | <a href="?Fmenu=General&Smenu=clear">清理服务器缓存</a></div>
		<%
		If Request.QueryString("Smenu") = "visitors" Then
		%>
			   <table border="0" cellpadding="2" cellspacing="1" class="TablePanel">
		<%
		If CheckStr(Request.QueryString("Page"))<>Empty Then
		    Curpage = CheckStr(Request.QueryString("Page"))
		    If IsInteger(Curpage) = False Or Curpage<0 Then Curpage = 1
		Else
		    Curpage = 1
		End If
		Dim bCounter, PageC
		Set bCounter = Server.CreateObject("ADODB.RecordSet")
		SQL = "SELECT * FROM blog_Counter order by coun_Time desc"
		bCounter.Open SQL, Conn, 1, 1
		PageC = 0
		If Not bCounter.EOF Then
		    bCounter.PageSize = 30
		    bCounter.AbsolutePage = CurPage
		    Dim bCounter_nums
		    bCounter_nums = bCounter.RecordCount
		    response.Write "<tr><td colspan=""6"" style=""border-bottom:1px solid #999;""><div class=""pageContent"">"&MultiPage(bCounter_nums, 30, CurPage, "?Fmenu=General&Smenu=visitors&", "", "float:left","")&"</div><div class=""Content-body"" style=""line-height:200%""></td></tr>"
		%>
		        <tr align="center">
		          <td width="100" nowrap="nowrap" class="TDHead">访客IP</td>
		          <td width="120" nowrap="nowrap" class="TDHead">访客操作系统</td>
		          <td nowrap="nowrap" class="TDHead">访客浏览器</td>
		          <td width="300" nowrap="nowrap" class="TDHead">访客介入地址</td>
		          <td class="TDHead">访客访问时间</td>
			   </tr>
			   <%
		Do Until bCounter.EOF Or PageC = bCounter.PageSize
		
		%>
		        <tr align="center">
		          <td nowrap><%=bCounter("coun_IP")%></td>
		          <td nowrap><%=bCounter("coun_OS")%></td>
		          <td nowrap><%=bCounter("coun_Browser")%></td>
		          <td nowrap align="left"><a href="<%=bCounter("coun_Referer")%>" target="_blank" title="<%=bCounter("coun_Referer")%>"><%=CutStr(bCounter("coun_Referer"),40)%></a></td>
		          <td nowrap><%=DateToStr(bCounter("coun_Time"),"Y-m-d H:I:S")%></td>
			   </tr>
		   <%
		bCounter.MoveNext
		PageC = PageC + 1
		Loop
		bCounter.Close
		Set bCounter = Nothing
		response.Write ("</table>")
		End If
		ElseIf Request.QueryString("Smenu") = "clear" Then
		        Response.Write "<div style='padding:4px 0px 4px 10px;border: 1px dotted #999;margin:2px;background:#ffffee'>"
		        Application.Lock
		        FreeApplicationMemory
		        Application.UnLock
		        Response.Write "<br/><span><b style='color:#040'>缓存清理完毕...	</b></span>"
		        Response.Write "</div>"
		ElseIf Request.QueryString("Smenu") = "Misc" Then
		%>
		<form action="ConContent.asp" method="post" style="margin:0px">
			<input type="hidden" name="action" value="General"/>
			<input type="hidden" name="whatdo" value="Misc"/>
		    <div align="left" style="padding:5px;"><%getMsg%>
		    
		   <b>1.基础数据初始化</b><br/>
		    &nbsp;&nbsp;&nbsp;<input type="checkbox" name="ReBulid" value="1" id="T1"/> <label for="T1">重建数据缓存</label><br/>
		    &nbsp;&nbsp;&nbsp;<input type="checkbox" name="ReTatol" value="1" id="T2"/> <label for="T2">重新统计网站数据</label><br/>
		    &nbsp;&nbsp;&nbsp;<input type="checkbox" name="CleanVisitor" value="1" id="T5"/> <label for="T5">清除访客记录</label><br/><br/>
		   <b>2.日志列表索引</b><br/>
		    &nbsp;&nbsp;&nbsp;<input type="checkbox" name="ReBulidIndex" value="1" id="T4"/> <label for="T4">重新建立日志列表翻页索引<span style="color:#666">（可以修复日志列表翻页错误的问题）</span></label><br/><br/>
		   <b>3.日志缓存和静态日志重建</b><span style="color:#f00">（这个过程可能会花很长时间,由你的日志数量来决定）</span><br/>
		    &nbsp;&nbsp;&nbsp;<input type="radio" name="ReBulidArticle" value="1" id="B1"/> <label for="B1">更新所有日志到文件，并且包含日志列表缓存 <span style="color:#666">（静态化所有日志内容数据，速度较慢）</span></label><br/>
		    &nbsp;&nbsp;&nbsp;<input type="radio" name="ReBulidArticle" value="2" id="B2"/> <label for="B2">只更新日志列表缓存<span style="color:#666">（在半静态和全静态之间切换的时候需要重新生成）</span></label><br/>
		    &nbsp;&nbsp;&nbsp;<input type="radio" name="ReBulidArticle" value="0" id="B3" checked	/> <label for="B3">什么都不做</label><br/><br/>
		    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input checked="checked" type="checkbox" name="silentMode" value="1" id="B4"/> <label for="B4">静态化日志安静模式 <span style="color:#666">（使用安静模式不出现进度条，速度较快）</span></label><br/><br/>
           <b>4.分段静态日志重建</b><br/>
			<%
			Dim ConurseDB, Start, Ends
			Set ConurseDB = Server.CreateObject("Adodb.Recordset")
				ConurseDB.open "blog_Content", Conn, 1, 2
				if ConurseDB.bof or ConurseDB.eof then
					Start = "0"
					Ends = "0"
				else
					ConurseDB.moveFirst
					Start = ConurseDB("log_ID")
					ConurseDB.moveLast
					Ends = ConurseDB("log_ID")
				end if
			ConurseDB.close
			Set ConurseDB = nothing
			%>
            &nbsp;&nbsp;&nbsp;<input type="checkbox" name="ReBulidPartStatus" value="1" id="C1"/> <label for="C1">分段静态化<span style="color:#666">（依据ID分段静态,减小服务器压力）</span></label><br/><br/>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;开始ID:<input onFocus="this.select();document.getElementById('C1').checked='checked'" type="text" name="ReBulidPartStatusSart" value="" id="C2" class="text" size="3"/> - 结束ID:<input onFocus="this.select();document.getElementById('C1').checked='checked'" type="text" name="ReBulidPartStatusEnd" value="" id="C3" class="text" size="3"/>&nbsp;&nbsp;<span style="color:#666">（您目前共有日志<%=blog_LogNums%>篇; 日志ID从<%=Start%>到<%=Ends%>）</span><br/><br />
            <script language="javascript" type="text/javascript">
				var Lists = eval("<%=BlogArticleID(1)%>");
	 			var CurrentIndex = 0;
				var IsStop = true;
				var pWidth = 0;
			</script>
            <style type="text/css">
				#msgbox{ height:17px; color:#0000ff}
				#processBar {width:80%;height:19px;border:1px #333 solid; margin-left:35px;}
				#process {width:0px;background-color: #ccc;text-align:center; height:19px; overflow:visible; text-align:left;}
				#percent{ color:#0000ff}
     		</style>
            <b>5.Ajax分段静态日志重建</b><br/><br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;※ Ajax重建日志,避免了ASP脚本超时的错误.&nbsp;&nbsp;&nbsp;<input type="button" class="text" value="开始Ajax重建日志" style=" font-size:11px; color:#333" onclick="StartHTML()" id="AjaxRebuildButton" />&nbsp;&nbsp;&nbsp;<input type="button" value="停止" onclick="StopHtml();" style=" font-size:11px; color:#333" class="text" />
            <br /><br />
            <div id="processBar">
  				<div id="process">
                </div>
  			</div><br />
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span id="msgbox">点击"开始Ajax重建日志"进行Ajax重建...</span>&nbsp;&nbsp;&nbsp;完成百分比:&nbsp;<span id="percent">0%</span>
            <br /><br /><br />
		   </div>
		   <div class="SubButton">
		      <input type="submit" name="Submit" value="保存配置" class="button"/>
		     </div>
			 </form>
		<%ElseIf Request.QueryString("Smenu") = "UpLoadSet" Then%>
			 	<form action="ConContent.asp" method="post">
			<input type="hidden" name="action" value="UpLoadSet"/>
			   <%getMsg
			   Dim UpLoadSet : UpLoadSet = Split(blog_UpLoadSet,"|")%>
		<fieldset>
		    <legend> 附件基本设置</legend>
		    <div align="left">
		      <table border="0" cellpadding="2" cellspacing="1">
		        <tr>
		          <td><div align="right"> 防盗链 </div></td>
		          <td align="left"><input name="Antimdown" type="checkbox" value="1" <%If UpLoadSet(0) = 1 then response.write ("checked=""checked""")%>/></td>
		        </tr>
		        <tr>
		          <td width="180"><div align="right"> 上传文件命名类型 </div></td>
		          <td>
				    <select name="FileNameMid">
				      <option value="0" <%If UpLoadSet(1) = 0 then response.write ("selected=""selected""")%>>上传日期时间</option>
				      <option value="1" <%If UpLoadSet(1) = 1 then response.write ("selected=""selected""")%>>文件原名</option>
				      <option value="2" <%If UpLoadSet(1) = 2 then response.write ("selected=""selected""")%>>上传日期时间_文件原名</option>
				      <option value="3" <%If UpLoadSet(1) = 3 then response.write ("selected=""selected""")%>>上传日期时间带-</option>
				      <option value="4" <%If UpLoadSet(1) = 4 then response.write ("selected=""selected""")%>>PJBlog默认</option>
				    </select>
				  </td>
		        </tr>
		        <tr>
		          <td width="180"><div align="right"> 上传文件命名方式 </div></td>
		          <td>
				    <select name="FileNameType">
				      <option value="0" <%If UpLoadSet(2) = 0 then response.write ("selected=""selected""")%>>类型</option>
				      <option value="1" <%If UpLoadSet(2) = 1 then response.write ("selected=""selected""")%>>前缀_类型</option>
				      <option value="2" <%If UpLoadSet(2) = 2 then response.write ("selected=""selected""")%>>类型_后缀</option>
				      <option value="3" <%If UpLoadSet(2) = 3 then response.write ("selected=""selected""")%>>前缀_类型_后缀</option>
				    </select>
				   <span style="color:#999">其中“类型”指上面设置的“上传文件命名类型”</span>
				  </td>
		        </tr>
				<tr>
		          <td width="180"><div align="right"> 上传文件命名前缀 </div></td>
		          <td align="left"><input name="FileNameLeft" type="text" size="20" class="text" value="<%=UpLoadSet(3)%>" maxlength="10"/></td>
		        </tr>
				<tr>
		          <td width="180"><div align="right"> 上传文件命名后缀 </div></td>
		          <td align="left"><input name="FileNameRight" type="text" size="20" class="text" value="<%=UpLoadSet(4)%>"/></td>
		        </tr>
		      </table>
		    </div>
			</fieldset>
			<fieldset>
		    <legend> 水印基本设置</legend>
		    <div align="left">
		      <table border="0" cellpadding="2" cellspacing="1">
		        <tr>
		          <td width="180"><div align="right"> 水印位置 </div></td>
		          <td>
				    <select name="Place">
				      <option value="0" <%If UpLoadSet(5) = 0 then response.write ("selected=""selected""")%>>默认位置</option>
				      <option value="1" <%If UpLoadSet(5) = 1 then response.write ("selected=""selected""")%>>顶部左侧</option>
				      <option value="2" <%If UpLoadSet(5) = 2 then response.write ("selected=""selected""")%>>顶部居中</option>
				      <option value="3" <%If UpLoadSet(5) = 3 then response.write ("selected=""selected""")%>>顶部右侧</option>
				      <option value="4" <%If UpLoadSet(5) = 4 then response.write ("selected=""selected""")%>>中部左侧</option>
				      <option value="5" <%If UpLoadSet(5) = 5 then response.write ("selected=""selected""")%>>中部居中</option>
				      <option value="6" <%If UpLoadSet(5) = 6 then response.write ("selected=""selected""")%>>中部右侧</option>
				      <option value="7" <%If UpLoadSet(5) = 7 then response.write ("selected=""selected""")%>>底部左侧</option>
				      <option value="8" <%If UpLoadSet(5) = 8 then response.write ("selected=""selected""")%>>底部居中</option>
				      <option value="9" <%If UpLoadSet(5) = 9 then response.write ("selected=""selected""")%>>底部右侧</option>
				      <option value="10" <%If UpLoadSet(5) = 10 then response.write ("selected=""selected""")%>>随机位置</option>
					</select>
				  </td>
		        </tr>
		        <tr>
		          <td><div align="right"> 水印位置计算边距 </div></td>
		          <td align="left"><input name="Calculate" type="checkbox" value="1" <%If UpLoadSet(6) = 1 then response.write ("checked=""checked""")%>/></td>
		        </tr>
				<tr>
		          <td width="180"><div align="right"> 水印离左边距离 </div></td>
		          <td align="left"><input name="SYX" type="text" size="10" class="text" value="<%=UpLoadSet(7)%>"/> px</td>
		        </tr>
				<tr>
		          <td width="180"><div align="right"> 水印离顶部距离 </div></td>
		          <td align="left"><input name="SYY" type="text" size="10" class="text" value="<%=UpLoadSet(8)%>"/> px</td>
		        </tr>
				<tr>
		          <td width="180"><div align="right"> 添加图片边框颜色 </div></td>
		          <td align="left"><input name="SYPenColor" type="text" size="10" class="text" value="<%=UpLoadSet(9)%>"/></td>
		        </tr>
				<tr>
		          <td width="180"><div align="right"> 添加图片边框宽度 </div></td>
		          <td align="left"><input name="SYPenWidth" type="text" size="10" class="text" value="<%=UpLoadSet(10)%>"/> px</td>
		        </tr>
				<tr>
		          <td width="180"><div align="right"> 水印水平边距 </div></td>
		          <td align="left"><input name="SYPaddingH" type="text" size="10" class="text" value="<%=UpLoadSet(11)%>"/> px</td>
		        </tr>
				<tr>
		          <td width="180"><div align="right"> 水印垂直边距 </div></td>
		          <td align="left"><input name="SYPaddingV" type="text" size="10" class="text" value="<%=UpLoadSet(12)%>"/> px</td>
		        </tr>
				<tr>
		          <td width="180"><div align="right"> 水印透明度 </div></td>
		          <td align="left"><input name="SYAlpha" type="text" size="10" class="text" value="<%=UpLoadSet(13)%>"/> px</td>
		        </tr>
		      </table>
		    </div>
			</fieldset>
			<fieldset>
		    <legend> 图片水印设置</legend>
		    <div align="left">
		      <table border="0" cellpadding="2" cellspacing="1">
				<tr>
		          <td width="180"><div align="right"> 水印图片 </div></td>
		          <td align="left"><input name="PicPath" type="text" size="30" class="text" value="<%=UpLoadSet(14)%>"/></td>
		        </tr>
				<tr>
		          <td width="180"><div align="right"> 图片水印宽度 </div></td>
		          <td align="left"><input name="PicWidth" type="text" size="10" class="text" value="<%=UpLoadSet(15)%>"/> px</td>
		        </tr>
				<tr>
		          <td width="180"><div align="right"> 图片水印高度 </div></td>
		          <td align="left"><input name="PicHeight" type="text" size="10" class="text" value="<%=UpLoadSet(16)%>"/> px</td>
		        </tr>
		      </table>
		    </div>
			</fieldset>
			<fieldset>
		    <legend> 文字水印设置</legend>
		    <div align="left">
		      <table border="0" cellpadding="2" cellspacing="1">
				<tr>
		          <td width="180"><div align="right"> 水印文字 </div></td>
		          <td align="left"><input name="Character" type="text" size="40" class="text" value="<%=UpLoadSet(17)%>"/></td>
		        </tr>
				<tr>
		          <td width="180"><div align="right"> 水印字体颜色 </div></td>
		          <td align="left"><input name="FontColor" type="text" size="10" class="text" value="<%=UpLoadSet(18)%>"/></td>
		        </tr>
				<tr>
		          <td width="180"><div align="right"> 水印字体大小 </div></td>
		          <td align="left"><input name="FontSize" type="text" size="10" class="text" value="<%=UpLoadSet(19)%>"/> px</td>
		        </tr>
		        <tr>
		          <td width="180"><div align="right"> 水印字体类型 </div></td>
		          <td>
				    <select name="FontFamily">
				      <option value="宋体" <%If UpLoadSet(20) = "宋体" then response.write ("checked=""checked""")%>>宋体</option>
				      <option value="黑体" <%If UpLoadSet(20) = "黑体" then response.write ("checked=""checked""")%>>黑体</option>
				      <option value="Arial" <%If UpLoadSet(20) = "Arial" then response.write ("checked=""checked""")%>>Arial</option>
				      <option value="Book Antiqua" <%If UpLoadSet(20) = "Book Antiqua" then response.write ("checked=""checked""")%>>Book Antiqua</option>
				      <option value="Century Gothic" <%If UpLoadSet(20) = "Century Gothic" then response.write ("checked=""checked""")%>>Century Gothic</option>
				      <option value="Courier New" <%If UpLoadSet(20) = "Courier New" then response.write ("checked=""checked""")%>>Courier New</option>
				      <option value="Georgia" <%If UpLoadSet(20) = "Georgia" then response.write ("checked=""checked""")%>>Georgia</option>
				      <option value="Impact" <%If UpLoadSet(20) = "Impact" then response.write ("checked=""checked""")%>>Impact</option>
				      <option value="Courier New" <%If UpLoadSet(20) = "Courier New" then response.write ("checked=""checked""")%>>Courier New</option>
				      <option value="Tahoma" <%If UpLoadSet(20) = "Tahoma" then response.write ("checked=""checked""")%>>Tahoma</option>
				      <option value="Times New Roman" <%If UpLoadSet(20) = "Times New Roman" then response.write ("checked=""checked""")%>>Times New Roman</option>
				      <option value="Verdana" <%If UpLoadSet(20) = "Verdana" then response.write ("checked=""checked""")%>>Verdana</option>
				    </select>
				  </td>
		        </tr>
		        <tr>
		          <td><div align="right"> 加粗 </div></td>
		          <td align="left"><input name="FontBold" type="checkbox" value="1" <%If UpLoadSet(21) = 1 then response.write ("checked=""checked""")%>/></td>
		        </tr>
		        <tr>
		          <td><div align="right"> 斜体 </div></td>
		          <td align="left"><input name="FontItalic" type="checkbox" value="1" <%If UpLoadSet(22) = 1 then response.write ("checked=""checked""")%>/></td>
		        </tr>
				<tr>
		          <td width="180"><div align="right"> 水印阴影颜色 </div></td>
		          <td align="left"><input name="FontShadowColor" type="text" size="10" class="text" value="<%=UpLoadSet(23)%>"/></td>
		        </tr>
				<tr>
		          <td width="180"><div align="right"> 水印阴影向右偏移量 </div></td>
		          <td align="left"><input name="FontShadowXOffset" type="text" size="10" class="text" value="<%=UpLoadSet(24)%>"/> px</td>
		        </tr>
				<tr>
		          <td width="180"><div align="right"> 水印阴影向下偏移量 </div></td>
		          <td align="left"><input name="FontShadowYOffset" type="text" size="10" class="text" value="<%=UpLoadSet(25)%>"/> px</td>
		        </tr>
		      </table>
		    </div>
			</fieldset>	<div class="SubButton">
		      <input type="submit" name="Submit" value="保存配置" class="button"/>
		     </div>
			 </form>
			 <%else%>
			 	<form action="ConContent.asp" method="post">
			<input type="hidden" name="action" value="General"/>
			<input type="hidden" name="whatdo" value="General"/>
			   <%getMsg%>
		<fieldset>
		    <legend> 站点基本信息</legend>
		    <div align="left">
		      <table border="0" cellpadding="2" cellspacing="1">
		        <tr>
		          <td width="180"><div align="right"> BLOG 名称 </div></td>
		          <td align="left"><input name="SiteName" type="text" size="30" class="text" value="<%=SiteName%>"/></td>
		        </tr>
				<tr>
		          <td width="180"><div align="right"> BLOG 副标题 </div></td>
		          <td align="left"><input name="blog_Title" type="text" size="50" class="text" value="<%=blog_Title%>"/></td>
		        </tr>
				<tr>
		          <td width="180"><div align="right"> 站长昵称 </div></td>
		          <td align="left"><input name="blog_master" type="text" size="10" class="text" value="<%=blog_master%>" maxlength="10"/></td>
		        </tr>
				<tr>
		          <td width="180"><div align="right"> 站长邮件地址 </div></td>
		          <td align="left"><input name="blog_email" type="text" size="50" class="text" value="<%=blog_email%>"/></td>
		        </tr>
		        <tr>
		          <td width="180"><div align="right"> BLOG 地址
		                  <div class="shuom">关系到<strong>RSS</strong>地址的可读性</div>
		          </div></td>
		          <td align="left"><input name="SiteURL" type="text" size="50" class="text" value="<%=SiteURL%>"/></td>
		        </tr>
				<tr>
		          <td width="180"><div align="right"> 首页 KeyWords 设置 </div></td>
		          <td align="left"><input name="blog_KeyWords" type="text" size="50" class="text" value="<%=blog_KeyWords%>"/></td>
		        </tr>
				<tr>
		          <td width="180"><div align="right"> 首页 Description 设置 </div></td>
		          <td align="left"><input name="blog_Description" type="text" size="50" class="text" value="<%=blog_Description%>"/></td>
		        </tr>
				<tr>
		          <td width="180"><div align="right"> 网站备案信息 </div></td>
		          <td align="left"><input name="blog_about" type="text" size="50" class="text" value="<%=blogabout%>"/></td>
		        </tr>  
		        <tr>
		          <td><div align="right">Blog对外开放</div></td>
		          <td align="left"><input name="SiteOpen" type="checkbox" value="1" <%if Application(CookieName & "_SiteEnable")=1 then response.write ("checked=""checked""")%>/></td>
		        </tr>
		      </table>
		    </div>
			</fieldset>
            
            <fieldset>
		    <legend> 整站功能的开启与关闭</legend>
		    <div align="left">
            	<table border="0" cellpadding="2" cellspacing="1">
                	<tr>
                    	<td width="180" align="right">使用密码保护功能</td>
                        <td><input name="blog_PasswordProtection" type="checkbox" value="1" <%if blog_PasswordProtection then response.write ("checked=""checked""")%>  /> </td>
                        <td><div class="shuom"><strong>[勾选]</strong>表示开启密码保护功能.</div></td>
                    </tr>
                    <tr>
		          		<td><div align="right"> 是否开启评论留言审核功能</div></td>
		          		<td align="left"><input name="AuditOpen" type="checkbox" value="1" <%if blog_AuditOpen then Response.Write ("checked=""checked""")%>/></td>
                        <td align="left"><div class="shuom"><strong>[勾选]</strong>表示开启评论留言审核功能</div></td>
		        	</tr>
                    <tr>
		          		<td><div align="right"> 是否开启Gravatar头像功能</div></td>
		          		<td align="left"><input name="GravatarOpen" type="checkbox" value="1" <%if blog_GravatarOpen then Response.Write ("checked=""checked""")%>/></td>
                        <td align="left"><div class="shuom"><strong>[勾选]</strong>表示开启Gravatar头像功能</div></td>
		        	</tr>  
                </table>
            </div>
            </fieldset>
            
            
            
			<fieldset>
		    <legend> 日志保存设置</legend>
		    <div align="left">
		      <table border="0" cellpadding="2" cellspacing="1">
		 	  <tr><td width="180" align="right" valign="top" style="padding-top:8px">日志输出模式</td><td>
		 	   
		 	   <label for="p1" >
			 	  <input id="p1" name="blog_postFile" type="radio" value="0" <%if blog_postFile = 0 then response.write ("checked=""checked""")%>/> 全动态模式 (不推荐)
			 	  <br/> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:#999">文章数据从数据库里直接获取</span> <br/>
		 	  </label>
		 	  
		 	  <label for="p2" >
			 	  <input id="p2" name="blog_postFile" type="radio" value="1" <%if blog_postFile = 1 then response.write ("checked=""checked""")%>/> 半静态模式 (适合喜欢个性化的用户) 
			 	  <br/> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:#999">把文章缓存成文件，并且插件功能不受影响. </span> <br/>
		 	  </label>
		 	  
		 	  <label for="p3" >
			 	  <input id="p3" name="blog_postFile" type="radio" value="2" <%if blog_postFile = 2 then response.write ("checked=""checked""")%>/> 全静态模式<span style="color:#f00;font-size:8px">new</span> (适合在乎seo和速度的用户) 
			 	  <br/> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:#999">把文章保存成全静态文件. 需要注意的是，这种模式下文章页面无法使用插件. </span> <br/>
		 	  </label>
			  <div style="border-top:1px solid #ccc;padding-top:5px;margin-top:5px">
		 	  <img src="images/notify.gif"/> <b>温馨提示:</b> 进行半静态和全静态切换后，需要到 <a href="ConContent.asp?Fmenu=General&Smenu=Misc">初始化数据</a> 里更新<b>日志列表缓存</b>。
		 	  </div>
		 	  <br/>
		 	  <%if not CheckObjInstalled("ADODB.Stream") then response.write "<b style='color:#f00'>需要 ADODB.Stream 组件支持</b>"%>
		 	  </td></tr>
		 	  <tr><td width="180" align="right">日志预览分割类型</td><td><select name="blog_SplitType"><option value="0">按照字符数量分割</option><option  value="1" <%if blog_SplitType then response.write ("selected=""selected""")%>>按照行数分割</option></select> <span style="color:#999">只对重新编辑日志或新建日志有效</span></td></tr>
			  <tr><td width="180" align="right">日志预览最大字符数</td><td><input name="blog_introChar" type="text" size="5" class="text" value="<%=blog_introChar%>"/> 个 <span style="color:#999">只对UBB编辑器有效</span></td></tr>
			  <tr><td width="180" align="right">日志预览切割行数</td><td><input name="blog_introLine" type="text" size="5" class="text" value="<%=blog_introLine%>"/> 行</td></tr>
              <tr><td width="180" align="right">草稿自动保存时间间隔</td><td><input name="blog_SaveTime" type="text" size="5" class="text" value="<%=blog_SaveTime%>"/> 秒</td></tr>
		      </table>
		    </div>
			</fieldset>	
			<fieldset>
		    <legend> 显示设置</legend>
		    <div align="left">
		      <table border="0" cellpadding="2" cellspacing="1">
			  <tr><td width="180" align="right">每页显示日志</td><td width="300"><input name="blogPerPage" type="text" size="5" class="text" value="<%=blogPerPage%>"/> 篇</td></tr>
			  <tr><td width="180" align="right">默认显示模式</td><td width="300"><select name="blog_DisMod"><option value="0">普通</option><option  value="1" <%if blog_DisMod then response.write ("selected=""selected""")%>>列表</option></select></td></tr>
			  <tr><td width="180" align="right">是否在首页显示图片友情链接</td><td width="300"><input name="blog_ImgLink" type="checkbox" value="1" <%if blog_ImgLink then response.write ("checked=""checked""")%>  /> </td></tr>
		      </table>
		    </div>
			</fieldset>
			
			<fieldset>
		    <legend> 评论设置</legend>
				<table border="0" cellpadding="2" cellspacing="1">
			    <tr><td width="180" align="right">每页显示评论数</td><td width="300"><input name="blogcommpage" type="text" size="5" class="text" value="<%=blogcommpage%>"/> 篇</td></tr>
				<tr><td width="180" align="right">发表评论时间间隔</td><td width="300"><input name="blog_commTimerout" type="text" size="5" class="text" value="<%=blog_commTimerout%>"/> 秒</td></tr>
				<tr><td width="180" align="right">发表评论字数限制</td><td width="300"><input name="blog_commLength" type="text" size="5" class="text" value="<%=blog_commLength%>"/> 字</td></tr>
				<tr><td width="180" align="right">发表评论必须输入验证码</td><td width="300"><input name="blog_validate" type="checkbox" value="1" <%if blog_validate then response.write ("checked=""checked""")%>  /> <span style="color:#999">可以让会员不用输入验证码，只有全动态模式有效</span> </td></tr>
				<tr><td width="180" align="right">禁用评论UBB代码</td><td width="300"><input name="blog_commUBB" type="checkbox" value="1" <%if blog_commUBB then response.write ("checked=""checked""")%>  /></td></tr>
				<tr><td width="180" align="right">禁用评论贴图</td><td width="300"><input name="blog_commIMG" type="checkbox" value="1" <%if blog_commIMG then response.write ("checked=""checked""")%> /></td></tr>
				</table>
			</fieldset>
			<fieldset>
		    <legend> Wap设置</legend>
		    <div align="left">
		      <table border="0" cellpadding="2" cellspacing="1">
			  <tr><td width="180" align="right">允许使用wap方式浏览Blog</td><td width="300"><input name="blog_wap" type="checkbox" value="1" <%if blog_wap then response.write ("checked=""checked""")%>  /></td></tr>
			  <tr><td width="180" align="right">允许wap使用简单HTML</td><td width="300"><input name="blog_wapHTML" type="checkbox" value="1" <%if blog_wapHTML then response.write ("checked=""checked""")%>  /></td></tr>
			  <tr><td width="180" align="right">允许wap显示图片</td><td width="300"><input name="blog_wapImg" type="checkbox" value="1" <%if blog_wapImg then response.write ("checked=""checked""")%>  /></td></tr>
			  <tr><td width="180" align="right">允许wap保留文章超链接</td><td width="300"><input name="blog_wapURL" type="checkbox" value="1" <%if blog_wapURL then response.write ("checked=""checked""")%>  /></td></tr>
			  <tr><td width="180" align="right">允许通过wap登录</td><td width="300"><input name="blog_wapLogin" type="checkbox" value="1" <%if blog_wapLogin then response.write ("checked=""checked""")%>  /></td></tr>
			  <tr><td width="180" align="right">允许通过wap发评论</td><td width="300"><input name="blog_wapComment" type="checkbox" value="1" <%if blog_wapComment then response.write ("checked=""checked""")%>  /></td></tr>
			  <tr><td width="180" align="right">Wap日志显示数量</td><td width="300"><input name="blog_wapNum" type="text" size="5" class="text" value="<%=blog_wapNum%>"/> 篇</td></tr>
		      </table>
		    </div>
			</fieldset>
			<fieldset>
		    <legend> 用户注册与过滤</legend>
		    <div align="left">
		      <table border="0" cellpadding="2" cellspacing="1">
		  		<tr><td width="180" align="right">不允许注册新用户</td><td width="300"><input name="blog_Disregister" type="checkbox" value="1" <%if blog_Disregister then response.write ("checked=""checked""")%>  /> </td></tr> 
		  	    <tr><td width="180" align="right">访客记录最大值</td><td width="300"><input name="blog_CountNum" type="text" size="5" class="text" value="<%=blog_CountNum%>"/>  <span style="color:#999">设置为0则不进行任何访客记录</span></td></tr>
		     <tr>
		         <td width="180" valign="top"><div align="right">注册名字过滤
		              <div class="shuom">用"|"分割需要过滤的名字</div>
		              <div class="shuom"></div>
		          </div></td>
		          <td align="left"><textarea name="Register_UserNames" cols="50" rows="5"><%=Register_UserNames%></textarea></td>
		        </tr>
		        <tr>
		          <td width="180" valign="top"><div align="right">IP过滤
		              <div class="shuom">以下Ip地址将无法访问Blog</div>
		              <div class="shuom">使用"|"分割IP地址,IP地址可以包含通配符号"*"用来禁止某个IP段的无法访问Blog</div>
		          </div></td>
		          <td align="left"><textarea name="FilterIPs" cols="50" rows="5"><%=FilterIPs%></textarea></td>
		        </tr>
		     </table>
		    </div>
			</fieldset>	<div class="SubButton">
		      <input type="submit" name="Submit" value="保存配置" class="button"/>
		     </div>
			 </form>
			 <%end if%>
			 </td>
		  </tr></table>
	<%
end Sub
%>