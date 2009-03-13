<%
'=================================
' 日志分类管理
'=================================
Sub c_categories
	    Dim Arr_Category, blog_Cate, blog_Cate_Item, Icons_Lists, Icons_List, CateInOpstions
	    Dim CategoryListDB
		%>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
		  <tr>
		    <th class="CTitle"><%=categoryTitle%></th>
		  </tr>
		  <tr>
		    <td class="CPanel">
		    <div class="SubMenu"><a href="?Fmenu=Categories">设置日志分类</a> | <a href="?Fmenu=Categories&Smenu=move">批量移动日志</a> | <a href="?Fmenu=Categories&Smenu=del">批量删除日志</a> | <a href="?Fmenu=Categories&Smenu=tag">Tag管理</a></div>
		<%
		If Request.QueryString("Smenu") = "tag" Then
		%>
		   <form action="ConContent.asp" method="post" style="margin:0px">
		   <input type="hidden" name="action" value="Categories"/>
		   <input type="hidden" name="whatdo" value="Tag"/>
		      <div align="left" style="padding:5px;"><%getMsg%>
		
			   <table border="0" cellpadding="2" cellspacing="1" class="TablePanel">
		        <tr align="center">
				  <td width="16" nowrap="nowrap" class="TDHead">&nbsp;</td>
		          <td width="100" nowrap="nowrap" class="TDHead">Tag名称</td>
		          <td  nowrap="nowrap" class="TDHead">日志数量</td>
			   </tr>
			   <%
		Dim BTag
		Set BTag = conn.Execute("select * from blog_tag order by tag_id desc")
		Do Until BTag.EOF
		
		%>
			   <tr align="center">
				  <td><input name="selectTagID" type="checkbox" value="<%=BTag("tag_id")%>"/></td>
		          <td><input name="TagID" type="hidden" value="<%=BTag("tag_id")%>"/><input name="tagName" type="text" size="14" class="text" value="<%=BTag("tag_name")%>"/></td>
		          <td><input name="tagCount" type="text" size="2" class="text" value="<%=BTag("tag_count")%>" readonly style="background:#ffe"/> 篇</td>
			   </tr>
			   <%
		BTag.movenext
		Loop
		
		%>
			    <tr align="center" bgcolor="#D5DAE0">
		        <td colspan="3" class="TDHead" align="left" style="border-top:1px solid #9EA9C5"><a name="AddLink"></a><img src="images/add.gif" style="margin:0px 2px -3px 2px"/>添加Tag</td>
		       </tr>	
		        <tr align="center">
				  <td></td>
		          <td><input name="TagID" type="hidden" value="-1"/><input name="tagName" type="text" size="14" class="text"/></td>
		          <td></td>
			   </tr>
			  </table>
		  <div class="SubButton" style="text-align:left;">
		    	 <select name="doModule">
					 <option value="SaveAll">保存所有Tag</option>
					 <option value="DelSelect">删除所选Tag</option>
				 </select>
			  <input type="submit" name="Submit" value="提交" class="button" style="margin-bottom:0px"/> 
		     </div>	</form>
		    <div style="color:#f00">提示:在发表和编辑日志的同时,也可以直接输入多个tag.系统会自动添加不存在的tag</div>
		     
		    </div>
		</td></tr></table>
		</div>
		<%
		ElseIf Request.QueryString("Smenu") = "move" Then
		    Set CategoryListDB = conn.Execute("select * from blog_Category order by cate_local asc, cate_Order desc")
		    Do While Not CategoryListDB.EOF
		        If Not CategoryListDB("cate_OutLink") Then
		            CateInOpstions = CateInOpstions&"<option value="""&CategoryListDB("cate_ID")&""">&nbsp;&nbsp;"&CategoryListDB("cate_Name")&" ["&CategoryListDB("cate_count")&"]</option>"
		        End If
		        CategoryListDB.movenext
		    Loop
		    Set CategoryListDB = Nothing
		%>
		  <form action="ConContent.asp" method="post" style="margin:0px" onSubmit="return CheckMove()">
		   <input type="hidden" name="action" value="Categories"/>
		   <input type="hidden" name="whatdo" value="move"/>
		    <div align="left" style="padding:5px;"><%getMsg%>
		    
		<select name="source">
		<option value="null" style="color:#333">源日志分类</option>
		<option value="null" style="color:#333">-----------------</option>
		
		<%=CateInOpstions%>
		</select> 移动到 
		<select name="target">
		<option value="null" style="color:#333">目标日志分类</option>
		<option value="null" style="color:#333">-----------------</option>
		<%=CateInOpstions%>
		</select> <input type="submit" name="Submit" value="移动日志" class="button" style="margin-bottom:-1px;"/>
		    </div>
		     </form></td>
		  </tr>
			 <%
		ElseIf Request.QueryString("Smenu") = "del" Then
		    Dim TempOsel, FilterO
		    Set CategoryListDB = conn.Execute("select * from blog_Category order by cate_local asc, cate_Order desc")
		    FilterO = checkstr(request.QueryString("filter"))
		    If Len(FilterO)<1 Then FilterO = -1
		    Do While Not CategoryListDB.EOF
		        If Not CategoryListDB("cate_OutLink") Then
		            If Int(FilterO) = CategoryListDB("cate_ID") Then TempOsel = "selected"
		            CateInOpstions = CateInOpstions&"<option value="""&CategoryListDB("cate_ID")&""" "&TempOsel&">&nbsp;&nbsp;&nbsp; - "&CategoryListDB("cate_Name")&" ["&CategoryListDB("cate_count")&"]</option>"
		            TempOsel = ""
		        End If
		        CategoryListDB.movenext
		    Loop
		    Set CategoryListDB = Nothing
		
		%>
		  <form action="ConContent.asp" method="post" style="margin:0px">
		   <input type="hidden" name="action" value="Categories"/>
		   <input type="hidden" name="whatdo" value="batdel"/>
		    <div align="left" style="padding:5px;"><%getMsg%>
				&nbsp;过滤器: <select name="filter" onChange="location='?Fmenu=Categories&Smenu=del&filter='+this.value">
					<option value="-1">显示所有日志</option>
						<%=CateInOpstions%>
					<option value="-3" <%if int(FilterO)=-3 then response.write "selected"%>>&nbsp;&nbsp;显示所有隐藏日志</option>
					<option value="-2" <%if int(FilterO)=-2 then response.write "selected"%>>&nbsp;&nbsp;显示所有草稿</option>
				</select>
				
				<table style="font-size:12px;margin:8px 0px 8px 0px">
				<%
		Dim DelContent
		If Int(FilterO) = -1 Then
		    SQL = "select * from blog_content order by log_posttime desc"
		ElseIf Int(FilterO) = -2 Then
		    SQL = "select * from blog_content where log_IsDraft=true order by log_posttime desc"
		ElseIf Int(FilterO) = -3 Then
		    SQL = "select * from blog_content where log_IsShow=false order by log_posttime desc"
		Else
		    SQL = "select * from blog_content where log_CateID="&Int(FilterO)&" and log_IsDraft=false order by log_posttime desc"
		End If
		Set DelContent = conn.Execute(SQL)
		If DelContent.EOF And DelContent.bof Then
		
		%>
						 <tr><td>没有找到符合条件的查询</td></tr>
						 <%Else
		    Dim TempImg
		    Do Until DelContent.EOF
		        If DelContent("log_IsShow") = False Then TempImg = "<img src=""images/icon_lock.gif"" alt="""" border=""0"" style=""margin:0px 0px -3px 2px""/>"
		        If Int(FilterO) = -2 Then
		
		%>
								 <tr><td><input name="CID" type="checkbox" value="<%=DelContent("log_ID")%>"/></td><td><a href="blogedit.asp?id=<%=DelContent("log_ID")%>" target="_blank"><%=DelContent("log_ID")%>. <%=DelContent("log_Title")%> <%=TempImg%></a></td></tr>
								 <%
		Else
		
		%>
								 <tr><td><input name="CID" type="checkbox" value="<%=DelContent("log_ID")%>"/></td><td><a href="default.asp?id=<%=DelContent("log_ID")%>" target="_blank"><%=DelContent("log_ID")%>. <%=DelContent("log_Title")%> <%=TempImg%></a></td></tr>
								 <%
		End If
		TempImg = ""
		DelContent.movenext
		Loop
		End If
		DelContent.Close
		Set DelContent = Nothing
		
		%>
				</table>
		&nbsp;<input type="button" value="全选" onClick="checkAll()" class="button" style="margin-bottom:-1px;"/>
		<input type="submit" name="Submit" value="删除选中的日志" class="button" style="margin-bottom:-1px;"/>
		    </div>
		     </form>
		     <br/>
			 <%Else
		    Set CategoryListDB = conn.Execute("select * from blog_Category order by cate_local asc, cate_Order asc")
		    If CheckObjInstalled("Scripting.FileSystemObject") Then Icons_Lists = Split(getPathList("images\icons")(1), "*")
		
		%>
			 <script>
			 	var il = [];
		        <%
		For Each Icons_List in Icons_Lists
		    response.Write ("il.push('"&Icons_List&"');" & Chr(13))
		Next
		
		%>
		       
		       function fillList(o){
			     	var v = o.getAttribute("dv");
			     	for (var i=0;i<il.length;i++){
			     		var n = new Option(il[i],"images/icons/" + il[i]);
			     		o.options.add(n);
			     	}

			     	if (!v) {o.selectedIndex = 0; }else {o.value  = v;}
		       }
		       
		       function fillAllList(){
		       		var ci = document.getElementsByName("Cate_icons");
		       		for (var i=0;i<ci.length;i++)
		       			fillList(ci[i]);
		       		fillList(document.getElementsByName("New_Cate_icons")[0]);
		       }
			 </script>
		   <form action="ConContent.asp" method="post" style="margin:0px">
		   <input type="hidden" name="action" value="Categories"/>
		   <input type="hidden" name="whatdo" value="Cate"/>
		   <input type="hidden" name="DelCate" value=""/>
		    <div align="left" style="padding:5px;"><%getMsg%>
		      <table border="0" cellpadding="2" cellspacing="1" class="TablePanel">
		        <tr align="center">
		          <td class="TDHead" nowrap>分类图标</td>
		          <td class="TDHead" nowrap>标题</td>
		          <td class="TDHead" nowrap>提示说明</td>
				  <%If blog_postFile = 2 Then%>
				  <td class="TDHead" nowrap>文章存放目录名</td>
				  <%end if%>
		          <td width="180" class="TDHead" nowrap>外部链接</td>
		          <td width="29" class="TDHead" nowrap>排序</td>
		          <td class="TDHead" nowrap>位置</td>
		          <td class="TDHead" nowrap>保密</td>
		          <td class="TDHead" nowrap>日志数量</td>
		          <td align="center" class="TDHead">&nbsp;</td>
		        </tr>
		        <%
		Do While Not CategoryListDB.EOF
		
		%>
		        <tr id="Catetr_<%=CategoryListDB("cate_ID")%>" style="background:<%
		If Int(CategoryListDB("cate_local")) = 1 Then
		    response.Write ("#a9c9e9")
		ElseIf Int(CategoryListDB("cate_local")) = 2 Then
		    response.Write ("#bcf39e")
		Else
		
		End If
		
		%>">
		          <td align="center" nowrap>
		          <img id="" name="CateImg_<%=CategoryListDB("cate_ID")%>" src="<%=CategoryListDB("cate_icon")%>" width="16" height="16" />
		         <%if CheckObjInstalled("Scripting.FileSystemObject") then%>
		          <select name="Cate_icons" dv="<%=CategoryListDB("cate_icon")%>" onChange="document.images['CateImg_<%=CategoryListDB("cate_ID")%>'].src=this.value;" style="width:120px;"></select>
		          <%else%>
		          <input name="Cate_icons" type="text" class="text" value="<%=CategoryListDB("cate_icon")%>" size="18" onChange="document.images['CateImg_<%=CategoryListDB("cate_ID")%>'].src=this.value"/>
		          <%end if%>
		          </td>
		          <td><input name="Cate_Name" type="text" class="text" value="<%=CategoryListDB("cate_Name")%>" size="14"/></td>
		          <td align="left"><input name="Cate_Intro" type="text" class="text" value="<%=CategoryListDB("cate_Intro")%>" size="20"/></td>
				  <!--分类名-->
				  <%If blog_postFile = 2 Then%>
				  <td><input name="Cate_Part" type="text" class="text" value="<%=CategoryListDB("cate_Part")%>" size="14" style="ime-mode:disabled" onblur="checkpart(this.value)"/><input name="OldCate_Name" type="hidden" value="<%=CategoryListDB("cate_Part")%>"/></td>
				  <%end if%>
				  <!--分类名-->
		          <td align="left"><input name="cate_URL" type="text" value="<%=CategoryListDB("cate_URL")%>" size="30" class="text" <%if CategoryListDB("cate_count")>0 Then response.write "readonly=""readonly"" style=""background:#e5e5e5"""%>/></td>
		          <td align="left"><input name="cate_Order" type="text" value="<%=CategoryListDB("cate_Order")%>" size="2" class="text"/></td>
		          <td align="center">
		           <select name="Cate_local" onChange="getElementById('Catetr_<%=CategoryListDB("cate_ID")%>').style.backgroundColor=(this.value==1)?'#a9c9e9':(this.value==2)?'#bcf39e':''">
			            <option value="0">同时</option>
			            <option value="1" <%if int(CategoryListDB("cate_local"))=1 then response.write "selected=""selected"""%>>顶部</option>
			            <option value="2" <%if int(CategoryListDB("cate_local"))=2 then response.write "selected=""selected"""%>>侧边</option>
		           </select>
		          </td>
		          <td>
		           <select name="cate_Secret" <%If CategoryListDB("cate_OutLink") Then response.write "disabled=""disabled"""%>>
		            <option value="0" style="background:#0f0">公开</option>
		            <option value="1" style="background:#f99" <%if int(CategoryListDB("cate_Secret")) then response.write "selected=""selected"""%>>保密</option>
		           </select>
		           <%If CategoryListDB("cate_OutLink") Then response.write "<input type=""hidden"" name=""cate_Secret"" value=""0""/>"%>
		           </td>
		          <td align="left" nowrap><input type="text" class="text" name="cate_count" value="<%=CategoryListDB("cate_count")%>" size="2" readonly="readonly" style="background:#ffe"/> 篇</td>
		          <td align="center">
				   <%if not CategoryListDB("cate_Lock") then response.write ("<a href=""javascript:CheckDel("&CategoryListDB("cate_ID")&")"" title=""删除该分类""><img border=""0"" src=""images/icon_del.gif"" width=""16"" height=""16"" /></a>")%>
		           <input type="hidden" name="Cate_ID" value="<%=CategoryListDB("cate_ID")%>"/>
		          </td>
		        </tr>
		        <%
		CategoryListDB.movenext
		Loop
		Set CategoryListDB = Nothing
		
		%>
		        <tr align="center" bgcolor="#D5DAE0">
		         <td colspan="<%If blog_postFile = 2 Then%>10<%else%>9<%end if%>" class="TDHead" align="left" style="border-top:1px solid #9EA9C5"><img src="images/add.gif" style="margin:0px 2px -3px 2px"/>添加日志分类</td>
		        </tr>
		        <tr>
		          <td align="center" nowrap><img name="CateImg" src="images/icons/1.gif" width="16" height="16" />
		          <%if CheckObjInstalled("Scripting.FileSystemObject") then%>
			          <select name="New_Cate_icons" onChange="document.images['CateImg'].src=this.value" style="width:120px;"></select>
		          <%else%>
		          <input name="New_Cate_icons" type="text" class="text" value="images/icons/1.gif" size="18" onChange="document.images['CateImg'].src=this.value"/>
		          <%end if%>
		          </td>
		          <td><input name="New_Cate_Name" type="text" size="14" class="text"/></td>
		          <td align="left"><input name="New_Cate_Intro" type="text" class="text" size="20"/></td>
				  <%If blog_postFile = 2 Then%>
				  <td><input name="New_Cate_Part" type="text" size="14" class="text" style="ime-mode:disabled"/></td>
				  <%end if%>
		          <td align="left"><input name="New_cate_URL" type="text" size="30" class="text"/></td>
		          <td align="left"><input name="New_cate_Order" type="text" size="2" class="text"/></td>
		          <td align="center">
		           <select name="New_Cate_local">
		            <option value="0">同时</option>
		            <option value="1">顶部</option>
		            <option value="2">侧边</option>
		           </select>
		          </td>
		          <td>
		          <select name="New_cate_Secret">
		            <option value="0" style="background:#0f0">公开</option>
		            <option value="1" style="background:#f99">保密</option>
		           </select>
		          </td>
		          <td align="center">&nbsp;</td>
		          <td align="center">&nbsp;</td>
		        </tr>
		      </table>
		      <script>fillAllList()</script>
		    </div>
		    <div style="color:#f00">
		    <%If blog_postFile = 2 Then%>
		    您当前的日志模式是 <b>【全静态化】</b> 方式, 您可以配置 <b>【文章存放目录名】</b> 来修改分类中的文章存储的目录,方便搜索引擎. <br/>例如: <%=siteURL%>article/文章存放目录名/0000-0-0-1.html <br />
		     <br />如果修改了 <b>【文章存放目录名】</b> 请到 <a href="ConContent.asp?Fmenu=General&Smenu=Misc" style="color:#00f">站点基本设置-初始化数据</a> ,重新生成所有日志到文件 ,不修改则前台显示日志的路径错误</div>
			<%else%>
			如果分类中存在日志，则无法使用外部连接.删除日志分类的时假如分类中存在日志,那么日志也会被删除
			<%end if%>
			<div class="SubButton">
		      <input type="submit" name="Submit" value="保存日志分类" class="button"/> 
		     </div>   
		     </form></td>
		  </tr>
		<%end if%>
		</td></tr>
		</table>
<%
end Sub
%>
<script language="javascript" type="text/javascript">
function checkpart(val){
   if ( val.indexOf("<",0) != -1 || val.indexOf(">",0) != -1 || val.indexOf("/",0) != -1 || val.indexOf("\\",0) != -1 || val.indexOf(":",0) != -1 || val.indexOf("*",0) != -1 || val.indexOf("?",0) != -1 || val.indexOf("|",0) != -1 || val.indexOf("\"",0) != -1) alert("文件夹名不能包含特殊字符")
}
</script>