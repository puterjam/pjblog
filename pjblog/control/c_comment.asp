<%
'=================================
' 评论留言管理
'=================================
Sub c_comment
		%>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
		  <tr>
		    <th class="CTitle"><%=categoryTitle%></th>
		   </tr>
		  <tr>
		    <td class="CPanel">
			<div class="SubMenu2">
				<b>评论管理:</b> <a href="?Fmenu=Comment">评论管理</a> | <a href="?Fmenu=Comment&Smenu=msg">留言管理</a> | <a href="?Fmenu=Comment&Smenu=trackback">引用管理</a>
				<br/>
				<b>评论过滤:</b> <a href="?Fmenu=Comment&Smenu=spam" title="面向初级用户,提供简单的过滤功能">初级过滤功能</a> | <a href="?Fmenu=Comment&Smenu=reg" title="面向高级级用户,提供功能强大的过滤功能">高级过滤功能</a>
			</div>
		   
		   <%
		If Request.QueryString("Smenu") = "spam" Then
		    Dim spamXml, spamList
		    Set spamXml = New PXML
		    spamXml.XmlPath = "spam.xml"
		    spamXml.Open
		
		%>
		     	 <div align="left" style="padding-top:5px;"><%getMsg%>
		     	 <form action="ConContent.asp" method="post" style="margin:0px" onSubmit="return addSpanKey()" name="filter">
					    <input type="hidden" name="action" value="Comment"/>
					    <input type="hidden" name="doModule" value="updateKey"/>
					    <input type="hidden" name="keyList" id="keyList" value=""/>
		     	 <table border="0" cellpadding="2" cellspacing="1" class="TablePanel">
		     	 <tr><td class="TDHead">过滤关键字</td></tr>
		     	 <tr><td><div style="width:394px;overflow:hidden">
		     	 <%
		spamList = "<select name=""spamList"" id=""spamList"" size=""20"" multiple=""multiple"" style=""width:400px;margin:-3px 0 -3px -3px"">"
		For i = 0 To spamXml.GetXmlNodeLength("key") -1
		    spamList = spamList & "<option value=""" & spamXml.SelectXmlNodeItemText("key", i) & """>" & spamXml.SelectXmlNodeItemText("key", i) & "</option>"
		Next
		response.Write spamList
		Set spamXml = Nothing
		
		%>
		       </div></td></tr>
		       <tr><td style="padding-bottom:5px;background:#FAE1AF"><img src="images/add.gif" alt="" style="margin:0 5px -3px"/>添加关键字：<input id="keyWord" type="text" size="27" class="text" onKeyPress="this.style.backgroundColor='#fff';"/>
		       <input type="Submit" name="Submit" value="添加" class="button" style="margin-bottom:-2px"/><input type="button" name="button" value="移除" class="button" style="margin-bottom:-2px" onClick="clearKey()"/>
		       </td></tr>
		       </table>
		       		<div class="SubButton" style="text-align:left;padding:5px;margin:0px">
				       	<button accesskey="s" class="button" style="margin-bottom:0px;margin-left:-5px;" onClick="submitList()" >保存关键字列表(<u>S</u>)</button>
				       
			        </div>          
		            <div style="color:#f00"><b>友情提示:</b><br/> - 添加或清除关键字后必须 <b>保存关键字列表</b>，垃圾关键字列表才生效。<br/> - 使用逗号或者空格输入字符串可以一次添加多个关键字<br/> - enter键直接插入关键字 ,用ctrl或shift键多选清除关键字</div>
		      </div>
		      </form>
		      <%
		ElseIf Request.QueryString("Smenu") = "reg" Then
		    Dim reSpamXml, reSpamList
		    Set reSpamXml = New PXML
		    reSpamXml.XmlPath = "reg.xml"
		    reSpamXml.Open
		
		
		%>
				 <script src="common/reg.js" Language="javascript"></script>
		     	 <div align="left" style="padding-top:5px;"><%getMsg%>
		     	 	<form name="reFilter" action="ConContent.asp" method="post" style="margin:0px">
					   <input type="hidden" name="action" value="Comment"/>
					   <input type="hidden" name="doModule" value="updateRegKey"/>
					   <input type="hidden" name="keyList" id="keyList" value=""/>
				       <div align="left" style="padding-top:5px;">
				     	 <table border="0" cellpadding="2" cellspacing="1" class="TablePanel">
				     	 <tr>
					     	 <td class="TDHead" width="150" align="center">规则描述</td>
					     	 <td class="TDHead" width="300" align="center">正则表达式</td>
					     	 <td class="TDHead" align="center">允许匹配次数</td>
				     	 </tr>
				     	 <tr>
				     	 	<td colspan=3 id="reList"></td>
				     	 </tr>
				        <tr align="center" bgcolor="#D5DAE0">
				      	   <td colspan="3" class="TDHead" align="left" style="border-top:1px solid #9EA9C5"><img src="images/add.gif" style="margin:0px 2px -3px 2px"/>添加新规则</td>
				        </tr>
				     	<tr>
				     	 	<td colspan=3>
				     	 		<div>
				     	 			<input id="reTitle" type="text" class="text" style="width:150px"/>
				 		     	 	<input id="reRules" type="text" class="text" style="width:300px"/>
				     	 			<input id="reTimes" type="text" class="text" style="width:30px" maxlength="3"/>
				     	 			<input name="des" type="button" class="button" value="添加" onClick="addRules()" style="font-size:12px;margin:0px;margin-top:-2px;margin-bottom:-1px;height:20px;"/>
		    	 				</div>
				     	 	</td>
				     	 </tr>
				        <tr align="center" bgcolor="#D5DAE0">
				      	   <td colspan="3" class="TDHead" align="left" style="border-top:1px solid #9EA9C5"><img src="images/html.gif" style="margin:0px 2px -3px 2px"/>测试文本区域</td>
				        </tr>
				     	<tr>
				     	 	<td colspan=3>
				     	 		<textarea id="testArea" name="testArea" style="width:100%;height:100px"><%=reSpamXml.SelectXmlNodeItemText("text",0)%></textarea>
				     	 	</td>
				     	 </tr>
				     	 <tr></tr>
					</table>
		       		<div class="SubButton" style="text-align:left;padding:5px;margin:0px">
				       	<button accesskey="t" class="button" style="margin-bottom:0px;margin-left:-5px;" onClick="testRules()">测试(<u>T</u>)</button>
				        <button accesskey="s" class="button" style="margin-bottom:0px" onClick="submitReList()">保存(<u>S</u>)</button>
			        </div>          
		            <div style="color:#f00"><b>友情提示:</b>
			            <br/> - <b>允许匹配次数:</b> 允许规则匹配的最大次数
			            <br/> - 当匹配次数设置为0时,说明评论中不允许有次规则的文字
		            </div>
					</form>
					<script>
					<%
		For i = 0 To reSpamXml.GetXmlNodeLength("key") -1
		    response.Write "addRule('" & Replace(reSpamXml.GetAttributes("key", "des", i), "'", "\'") & "','" & Replace(reSpamXml.GetAttributes("key", "re", i), "'", "\'") & "','" & reSpamXml.GetAttributes("key", "times", i) & "');" & Chr(13)
		Next
		
		%>
					</script>
				</div>
				<%
		Set reSpamXml = Nothing
		Else
		
		%>
				<form action="ConContent.asp" method="post" style="margin:0px">
				   <input type="hidden" name="action" value="Comment"/>
				   <input type="hidden" name="doModule" value="DelSelect"/>
				      <div align="left" style="padding-top:5px;"><%getMsg%>
				      <%
		Dim blog_Comment, comm_Num, commArr, commArrLen, Pcount, aUrl, pSize, saveButton
		Pcount = 0
		saveButton = "<input type=""button"" value=""保存全部"" class=""button"" onclick=""updateComment()"" style=""font-weight:bold;margin:0px;margin-bottom:5px;float:right;margin-right:6px""/>"
		Set blog_Comment = Server.CreateObject("Adodb.RecordSet")
		If Request.QueryString("Smenu") = "trackback" Then
		    SQL = "SELECT tb_ID,tb_Intro,tb_Site,tb_PostTime,tb_Title,blog_ID,tb_URL,C.log_Title FROM blog_Content C,blog_Trackback T WHERE T.blog_ID=C.log_ID ORDER BY tb_PostTime desc"
		    aUrl = "?Fmenu=Comment&Smenu=trackback&"
		    pSize = 100
		    saveButton = ""
		    response.Write "<input type=""hidden"" name=""whatdo"" value=""trackback""/>"
		ElseIf Request.QueryString("Smenu") = "msg" Then
		    SQL = "SELECT book_ID,book_Content,book_Messager,book_PostTime,book_IP,book_reply FROM blog_book ORDER BY book_PostTime desc"
		    aUrl = "?Fmenu=Comment&Smenu=msg&"
		    pSize = 12
		    response.Write "<input type=""hidden"" name=""whatdo"" value=""msg""/>"
		Else '评论
		    SQL = "SELECT comm_ID,comm_Content,comm_Author,comm_PostTime,comm_PostIP,blog_ID,T.log_Title, C.comm_IsAudit from blog_Comment C,blog_Content T WHERE C.blog_ID=T.log_ID ORDER BY C.comm_PostTime desc"
		    aUrl = "?Fmenu=Comment&"
		    pSize = 15
		    response.Write "<input type=""hidden"" name=""whatdo"" value=""comment""/>"
		End If
		
		%>
					       <div style="height:24px;">
						       <input type="button" value="删除所选内容" onClick="DelComment()" class="button" style="margin:0px;margin-bottom:5px;float:right;"/> 
						       <input type="button" value="全选" onClick="checkAll()" class="button" style="margin:0px;margin-bottom:5px;float:right;margin-right:6px"/>
						       <%=saveButton%>
						       <div id="page1" class="pageContent"></div>
					       </div>
					       <div class="msgDiv">
							<%
		blog_Comment.Open SQL, Conn, 1, 1
		If blog_Comment.EOF And blog_Comment.BOF Then
		    response.Write "</div>"
		Else
		    blog_Comment.PageSize = pSize
		    blog_Comment.AbsolutePage = CurPage
		    comm_Num = blog_Comment.RecordCount
		
		    commArr = blog_Comment.GetRows(comm_Num)
		    blog_Comment.Close
		    Set blog_Comment = Nothing
		    commArrLen = UBound(commArr, 2)
		    'commArr(3,Pcount)
		    Do Until Pcount = commArrLen + 1 Or Pcount = pSize
		        If Request.QueryString("Smenu") = "trackback" Then
		
		%>
									        <div class="item"><input type="hidden" name="CommentID" value="<%=commArr(0,Pcount)%>"/>
										        <div class="title"><span class="blogTitle"><a href="default.asp?id=<%=commArr(5,Pcount)%>" target="_blank" title="<%=commArr(7,Pcount)%>"><%=CutStr(commArr(7,Pcount),25)%></a></span><input type="checkbox" name="selectCommentID" value="<%=commArr(0,Pcount)%>|<%=commArr(5,Pcount)%>" onClick="highLight(this)"/><img src="images/icon_trackback.gif" alt=""/><b><a href="<%=commArr(6,Pcount)%>" target="_blank"><%=commArr(2,Pcount)%></a></b> <span class="date">[<%=DateToStr(commArr(3,Pcount),"Y-m-d H:I:S")%>]</span></div>
										        <div class="contentTB">
										         <b>标题:</b> <%=checkURL(HTMLDecode(commArr(4,Pcount)))%><br/>
										         <b>链接:</b> <a href="<%=commArr(6,Pcount)%>" target="_blank"><%=commArr(6,Pcount)%></a><br/>
										         <b>摘要:</b> <%=checkURL(HTMLDecode(commArr(1,Pcount)))%><br/>
										        </div>
										    </div>
									      <%
		ElseIf Request.QueryString("Smenu") = "msg" Then
		
		%>
									        <div class="item"><input type="hidden" name="CommentID" value="<%=commArr(0,Pcount)%>"/>
										        <div class="title"><input type="checkbox" name="selectCommentID" value="<%=commArr(0,Pcount)%>" onClick="highLight(this)"/><img src="images/reply.gif" alt=""/><b><%=HtmlEncode(commArr(2,Pcount))%></b> <span class="date">[<%=DateToStr(commArr(3,Pcount),"Y-m-d H:I:S")%> | <%=commArr(4,Pcount)%>]</span></div>
										        <div class="content"><textarea name="message_<%=commArr(0,Pcount)%>" onFocus="focusMe(this)" onBlur="blurMe(this)" onMouseOver="overMe(this)" onMouseOut="outMe(this)"><%=UnCheckStr(commArr(1,Pcount))%></textarea></div>
										        <div class="reply"><h5>回复内容:<%if len(trim(commArr(5,Pcount)))<1 or IsNull(commArr(5,Pcount)) then response.write "<span class=""tip"">(无回复留言)</span>"%></h5><textarea name="reply_<%=commArr(0,Pcount)%>" onFocus="focusMe(this)" onBlur="blurMe(this)" onMouseOver="overMe(this)" onMouseOut="outMe(this)" onChange="checkMe(this);document.getElementById('edited_<%=commArr(0,Pcount)%>').value=1"><%=UnCheckStr(commArr(5,Pcount))%></textarea><input id="edited_<%=commArr(0,Pcount)%>" name="edited_<%=commArr(0,Pcount)%>" type="hidden" value="0"></div>
										    </div>
									      <%
		Else '评论
		
		%>
									        <div class="item"><input type="hidden" name="CommentID" value="<%=commArr(0,Pcount)%>"/>
										        <div class="title"><span class="blogTitle"><a href="default.asp?id=<%=commArr(5,Pcount)%>" target="_blank" title="<%=commArr(6,Pcount)%>"><%=CutStr(commArr(6,Pcount),25)%></a></span><input type="checkbox" name="selectCommentID" value="<%=commArr(0,Pcount)%>|<%=commArr(5,Pcount)%>" onClick="highLight(this)"/><img src="images/icon_quote.gif" alt=""/><b><%=HtmlEncode(commArr(2,Pcount))%></b> <span class="date">[<%=DateToStr(commArr(3,Pcount),"Y-m-d H:I:S")%> | <%=commArr(4,Pcount)%>]&nbsp;&nbsp;<select name="CommentAudit"><option value="1" style="background:#FABABA" <%if commArr(7,Pcount) = false then Response.Write("Selected=""selected""")%>>&nbsp;未审核</option><option value="0" style="background:#C5FDB7" <%if commArr(7,Pcount) = true then Response.Write("Selected=""selected""")%>>审核通过</option></select></span></div>
										        <div class="content"><textarea name="message_<%=commArr(0,Pcount)%>" onFocus="focusMe(this)" onBlur="blurMe(this)" onMouseOver="overMe(this)" onMouseOut="outMe(this)"><%=UnCheckStr(commArr(1,Pcount))%></textarea></div>
										    </div>
									      <%
		End If
		Pcount = Pcount + 1
		Loop
		
		%>
								</div>
								<div style="height:24px;">
								       <input type="button" value="删除所选内容" onClick="DelComment()" class="button" style="margin:0px;margin-bottom:5px;float:right;"/> 
								       <input type="button" value="全选" onClick="checkAll()" class="button" style="margin:0px;margin-bottom:5px;float:right;margin-right:6px"/>
								       <%=saveButton%>
								       <div id="page2" class="pageContent"><%=MultiPage(comm_Num,pSize,CurPage,aUrl,"","","")%></div>
								       <script>document.getElementById("page1").innerHTML=document.getElementById("page2").innerHTML</script>
							    </div>
					  </div>
				 </form>
				   <%
		End If
		Set blog_Comment = Nothing
		End If
		
		%>
	  </td></tr>
	  </table>
<%
end Sub
%>