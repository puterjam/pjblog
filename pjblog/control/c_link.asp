<%
'=================================
' 好友链接管理
'=================================
Sub c_link
	Dim ToolBar
	ToolBar = "<div class='SubMenu'><b>友情链接:</b> <a href='?Fmenu=Link'>链接列表</a> | <a href='?Fmenu=Link&Smenu=LinkClass'>分类管理</a></div>"
	
If Request.QueryString("Smenu") = "LinkClass" Then
%>
		<form name="LinkClass" action="ConContent.asp" method="post" style="margin:0px">
            	<input type="hidden" name="action" value="Links"/>
	  		 	<input type="hidden" name="whatdo" value="LinkClassUpdate"/>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
		  <tr>
			<th class="CTitle"><%=categoryTitle%></th>
		  </tr>
          <tr>
		    <td class="CPanel">
            <%=ToolBar%>
            <%getMsg%>
                <table border="0" cellpadding="2" cellspacing="1" class="TablePanel">
                	<tr align="center">
	         	 		<td width="16" nowrap="nowrap" class="TDHead">&nbsp;</td>
	          			<td width="120" nowrap="nowrap" class="TDHead">分类名称</td>
	          			<td width="180" nowrap="nowrap" class="TDHead">分类说明</td>
	         	 		<td class="TDHead">排序</td>
		   			</tr>
                    <%
						Dim ClassLinkAdo
						Set ClassLinkAdo = Conn.Execute("Select * From blog_LinkClass Order By LinkClass_Order ASC")
						If ClassLinkAdo.Bof Or ClassLinkAdo.Eof Then
							Response.Write("")
						Else
							Do While Not ClassLinkAdo.Eof
					%>
                    <tr align="center">
                    	<td><input type="checkbox" name="SelectID" value="<%=ClassLinkAdo("LinkClass_ID")%>" /><input type="hidden" value="<%=ClassLinkAdo("LinkClass_ID")%>" name="LinkClass_ID" /></td>
                        <td><input type="text" value="<%=ClassLinkAdo("LinkClass_Name")%>" name="LinkClass_Name" size="30" class="text" /></td>
                        <td><input type="text" value="<%=ClassLinkAdo("LinkClass_Title")%>" name="LinkClass_Title" size="30" class="text" /></td>
                        <td><input type="text" value="<%=ClassLinkAdo("LinkClass_Order")%>" name="LinkClass_Order" size="2" class="text" /></td>
                    </tr>
                    <%
							ClassLinkAdo.MoveNext
							Loop
						End If
						Set ClassLinkAdo = Nothing
					%>
                    <tr align="center" bgcolor="#D5DAE0">
	         			<td colspan="7" class="TDHead" align="left" style="border-top:1px solid #9EA9C5"><a name="AddLink"></a><img src="images/add.gif" style="margin:0px 2px -3px 2px"/>添加新友情链接分类</td>
	        		</tr>	
	        		<tr align="center">
	          			<td>&nbsp;</td>
                        <td><input type="text" value="" name="New_LinkClass_Name" size="30" class="text" /></td>
                        <td><input type="text" value="" name="New_LinkClass_Title" size="30" class="text" /></td>
                        <td><input type="text" value="" name="New_LinkClass_Order" size="2" class="text" /></td>
		  			</tr>
                </table>
            </td>
          </tr>
          <tr>
          		<td>
                	<div class="SubButton">
	      		<input type="submit" name="Submit" value="保存友情链接分类" class="button"/><input type="button" name="dellinks" value="删除友情链接分类" class="button" onclick="DelLikClasses()"/> 
	     	</div>
                </td>
          </tr>
          </table>
          <script language="javascript" type="text/javascript">
		  		function DelLikClasses(){
					if (confirm("确定要删除所选友情链接分类吗？")){
		    			document.forms["LinkClass"].whatdo.value = "DelLinkClasses";
						document.forms["LinkClass"].submit();
					}
				}
		  </script>
         </form>
<%
Else
		%>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
		  <tr>
			<th class="CTitle"><%=categoryTitle%></th>
		  </tr>
		  <tr>
		    <td class="CPanel">
            <%=ToolBar%>
            <form name="filter" action="ConContent.asp" method="post" style="margin:0px">
			<%
		Dim disLink, disCount, disLinkClassID
		If session(CookieName&"_disLink") = "" Then session(CookieName&"_disLink") = "All"
		If session(CookieName&"_disCount") = "" Then session(CookieName&"_disCount") = "10"
		disLink = session(CookieName&"_disLink")
		disCount = session(CookieName&"_disCount")
		disLinkClassID = session(CookieName&"_disLinkClassID")
		Dim FilterWhere
		If CheckStr(Request.QueryString("Page"))<>Empty Then
		    Curpage = CheckStr(Request.QueryString("Page"))
		    If IsInteger(Curpage) = False Or Curpage<0 Then Curpage = 1
		Else
		    Curpage = 1
		End If
		
		%>
		<div class="SubMenu">过滤器: 
	   <input type="hidden" name="action" value="Links"/>
	   <input type="hidden" name="whatdo" value="Filter"/>
       <select name="disLinkClass" onchange="document.forms['filter'].submit()">
       		<option value="0">所有分类</option>
       		<%=SelectOutOption(disLinkClassID)%>
       </select>
	  	<select name="disLink" onChange="document.forms['filter'].submit()">
		  <option value="All">显示所有链接</option>
		  <option value="Allow" <%if disLink="Allow" then response.write ("selected=""selected""")%>>已通过验证链接</option>
		  <option value="NoAllow" <%if disLink="NoAllow" then response.write ("selected=""selected""")%>>未通过验证链接</option>
		  <option value="Top" <%if disLink="Top" then response.write ("selected=""selected""")%>>置顶链接</option>
		  <option value="NoTop" <%if disLink="NoTop" then response.write ("selected=""selected""")%>>未置顶链接</option>
		</select> 每页显示 	<select name="disCount" onChange="document.forms['filter'].submit()">
		  <option <%if int(disCount)=100 then response.write ("selected=""selected""")%>>100</option>
		  <option <%if int(disCount)=50 then response.write ("selected=""selected""")%>>50</option>
		  <option <%if int(disCount)=25 then response.write ("selected=""selected""")%>>25 </option>
		  <option <%if int(disCount)=10 then response.write ("selected=""selected""")%>>10</option>
		</select>
		条 <a href="#AddLink">添加新友情链接</a></div> </form>
	   <form name="Link" action="ConContent.asp" method="post" style="margin:0px">
	    <div align="left" style="padding:5px;"><%getMsg%>
	       <input type="hidden" name="action" value="Links"/>
	       <input type="hidden" name="whatdo" value="SaveLink"/>
	       <input type="hidden" name="ALinkID" value=""/>
	       <input type="hidden" name="Page" value="<%=Curpage%>"/>       
		   <table border="0" cellpadding="2" cellspacing="1" class="TablePanel">
	<%
	Select Case disLink
	    Case "All"
	        FilterWhere = ""
	    Case "Allow"
	        FilterWhere = " where link_IsShow=true"
	    Case "NoAllow"
	        FilterWhere = " where link_IsShow=false"
	    Case "Top"
	        FilterWhere = " where link_IsMain=true"
	    Case "NoTop"
	        FilterWhere = " where link_IsMain=false"
	    Case Else
	        FilterWhere = ""
	End Select
	Dim bLink
	If FilterWhere = "" Then
		If len(disLinkClassID) = 0 or disLinkClassID = "" or disLinkClassID = empty or disLinkClassID = null or session(CookieName&"_disLinkClassID") = 0 then
			FilterWhere = FilterWhere & ""
		Else
			FilterWhere = FilterWhere & " where Link_ClassID=" & disLinkClassID
		End If
	Else
		If len(disLinkClassID) = 0 or disLinkClassID = "" or disLinkClassID = empty or disLinkClassID = null or session(CookieName&"_disLinkClassID") = 0 then
			FilterWhere = FilterWhere & ""
		Else
			FilterWhere = FilterWhere & " and Link_ClassID=" & disLinkClassID
		End If
	End If
	Set bLink = Server.CreateObject("ADODB.RecordSet")
	SQL = "SELECT * FROM blog_Links"&FilterWhere&" order by link_IsShow desc,link_Order desc"
	bLink.Open SQL, Conn, 1, 1
	If Not bLink.EOF Then
	    bLink.PageSize = disCount
	    bLink.AbsolutePage = CurPage
	    Dim bLink_nums
	    bLink_nums = bLink.RecordCount
	    Dim MultiPages, PageCount
	    response.Write "<tr><td colspan=""8"" style=""border-bottom:1px solid #999;""><div class=""pageContent"">"&MultiPage(bLink_nums, disCount, CurPage, "?Fmenu=Link&Smenu=&", "", "float:left","")&"</div><div class=""Content-body"" style=""line-height:200%""></td></tr>"
	End If
	%>
	        <tr align="center">
	          <td width="16" nowrap="nowrap" class="TDHead">&nbsp;</td>
	          <td width="16" nowrap="nowrap" class="TDHead">&nbsp;</td>
	          <td width="120" nowrap="nowrap" class="TDHead">网站名称</td>
	          <td width="180" nowrap="nowrap" class="TDHead">网站地址</td>
	          <td width="250" nowrap="nowrap" class="TDHead">Logo图片地址</td>
              <td nowrap="nowrap" class="TDHead">分类</td>
	          <td class="TDHead">排序</td>
	          <td class="TDHead">操作</td>
		   </tr>
		   <%
	If Not bLink.EOF Then
	    Do Until bLink.EOF Or PageCount = bLink.PageSize
	        If Not bLink("link_IsShow") Then
	
	%>
	        <tr align="center" bgcolor="#FCF4BC">
	          <td><img src="images/slink.gif" alt="没有通过验证链接"/></td>
	          <td><input type="checkbox" name="checklinkID" value="<%=bLink("link_ID")%>"></td>
	          <td><input name="LinkID" type="hidden" value="<%=bLink("link_ID")%>"/><input name="LinkName" type="text" size="18" class="text" value="<%=bLink("link_Name")%>"/></td>
	          <td><input name="LinkURL" type="text" size="30" class="text" value="<%=bLink("link_URL")%>"/></td>
	          <td><input name="LinkLogo" type="text" size="40" class="text" value="<%=bLink("link_Image")%>"/></td>
              <td>
              	<select name="LinkClass">
                	<option value="0">未分类</option>
                	<%=SelectOutOption(bLink("Link_ClassID"))%>
              	</select>
              </td>
	          <td><input name="LinkOrder" type="text" class="text" size="2" value="<%=bLink("link_Order")%>"/></td>
	          <td><a href="#" onClick="ShowLink(<%=bLink("link_ID")%>)" title="通过该链接的验证"><img border="0" src="images/alink.gif" width="16" height="16" style="margin:0px 2px -3px 0px"/>通过</a> <a href="<%=bLink("link_URL")%>" target="_blank" title="查看该链接"><img border="0" src="images/icon_trackback.gif" width="16" height="16" style="margin:0px 2px -3px 0px"/>查看</a> <a href="#" onClick="Dellink(<%=bLink("link_ID")%>)" title="删除该链接"><img border="0" src="images/icon_del.gif" width="16" height="16" style="margin:0px 2px -3px 0px"/>删除</a> </td>
	        </tr>
			<%else%>
	        <tr align="center">
	          <td><%if bLink("link_IsMain") then response.write ("<img src=""images/urlInTop.gif"" alt=""置顶链接""/>") else response.write ("&nbsp;")%></td>
	          <td><input type="checkbox" name="checklinkID" value="<%=bLink("link_ID")%>"></td>
	          <td><input name="LinkID" type="hidden" value="<%=bLink("link_ID")%>"/><input name="LinkName" type="text" size="18" class="text" value="<%=bLink("link_Name")%>"/></td>
	          <td><input name="LinkURL" type="text" size="30" class="text" value="<%=bLink("link_URL")%>"/></td>
	          <td><input name="LinkLogo" type="text" size="40" class="text" value="<%=bLink("link_Image")%>"/></td>
              <td>
              	<select name="LinkClass">
                	<option value="0">未分类</option>
                	<%=SelectOutOption(bLink("Link_ClassID"))%>
              	</select>
              </td>
	          <td><input name="LinkOrder" type="text" class="text" size="2" value="<%=bLink("link_Order")%>"/></td>
	          <td><%if bLink("link_IsMain") then response.write ("<a href=""#"" onclick=""Toplink("&bLink("link_ID")&")"" title=""取消该链接在首页置顶""><img border=""0"" src=""images/ct.gif"" style=""margin:0px 2px -3px 0px""/>取消</a> ") else response.write ("<a href=""#"" onclick=""Toplink("&bLink("link_ID")&")"" title=""把该链接在首页置顶""><img border=""0"" src=""images/it.gif"" style=""margin:0px 2px -3px 0px""/>置顶</a> ")%>
			  <a href="<%=bLink("link_URL")%>" target="_blank" title="查看该链接"><img border="0" src="images/icon_trackback.gif" width="16" height="16" style="margin:0px 2px -3px 0px"/>查看</a> <a href="#" onClick="Dellink(<%=bLink("link_ID")%>)" title="删除该链接"><img border="0" src="images/icon_del.gif" width="16" height="16" style="margin:0px 2px -3px 0px"/>删除</a> </td>
		   </tr>
		<%End If
		bLink.MoveNext
		PageCount = PageCount + 1
		Loop
		bLink.Close
		Set bLink = Nothing
		End If
		%>
	        <tr align="center" bgcolor="#D5DAE0">
	         <td colspan="8" class="TDHead" align="left" style="border-top:1px solid #9EA9C5"><a name="AddLink"></a><img src="images/add.gif" style="margin:0px 2px -3px 2px"/>添加新友情链接</td>
	        </tr>	
	        <tr align="center">
	          <td>&nbsp;</td>
	          <td>&nbsp;</td>
	          <td><input name="new_LinkID" type="hidden" value="-1"/><input name="new_LinkName" type="text" size="18" class="text"/></td>
	          <td><input name="new_LinkURL" type="text" size="30" class="text" /></td>
	          <td><input name="new_LinkLogo" type="text" size="40" class="text" /></td>
              <td>
              	<select name="new_LinkClass">
                	<%=SelectOutOption(0)%>
              	</select>
              </td>
	          <td><input name="new_LinkOrder" type="text" class="text" size="2" /></td>
	          <td>&nbsp;</td>
		   </tr>
		 	</table>
	  </div>
	  <div class="SubButton">
	      <input type="submit" name="Submit" value="保存友情链接" class="button"/><input type="button" name="dellinks" value="删除友情链接" class="button" onclick="DelLiks()"/> 
	     </div>
		 <script language="javascript" type="text/javascript">
		 function DelLiks(){
		    if (confirm("确定要删除所选友情链接吗？")){
		    document.forms["Link"].whatdo.value = "DelLinks";
			document.forms["Link"].submit();
			}
		 }
		 </script>
	 </td>
	  </tr></table></div></form>
<%
End If
end Sub
%>