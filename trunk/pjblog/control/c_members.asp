<%
'=================================
' 权限和会员管理
'=================================
Sub c_members
	    Dim blog_Status, blog_Statu, StatusItem, blog_Status_Len
		%>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
		  <tr>
		    <th class="CTitle"><%=categoryTitle%></th>
		  </tr>
		  <tr>
		    <td class="CPanel">
			<div class="SubMenu"><a href="?Fmenu=Members">权限管理</a> | <a href="?Fmenu=Members&Smenu=Users">帐户管理</a></div>
		<%
		If Request.QueryString("Smenu") = "Users" Then
		
		%>
		<%getMsg%>
		 <form action="ConContent.asp" method="post" style="margin:0px">
		   <input type="hidden" name="action" value="Members"/>
		   <input type="hidden" name="whatdo" value="SaveUserRight"/>
		   <input type="hidden" name="DelID" value=""/>
		<table border="0" cellpadding="2" cellspacing="1" class="TablePanel" style="margin:5px;">
		<%
		blog_Status = Application(CookieName&"_blog_rights")
		Dim FindUser, FindUserFilter
		FindUser = Request.QueryString("User")
		If Len(FindUser)<1 Then
		    FindUserFilter = ""
		Else
		    FindUserFilter = " AND M.mem_Name='" & FindUser & "'"
		End If
		If CheckStr(Request.QueryString("Page"))<>Empty Then
		    Curpage = CheckStr(Request.QueryString("Page"))
		    If IsInteger(Curpage) = False Or Curpage<0 Then Curpage = 1
		Else
		    Curpage = 1
		End If
		Dim bMember, PageCM
		Set bMember = Server.CreateObject("ADODB.RecordSet")
		SQL = "SELECT M.*,S.stat_name,S.stat_title FROM blog_Member as M,blog_status as S where M.mem_Status=S.stat_name"&FindUserFilter&" order by mem_RegTime desc"
		bMember.Open SQL, Conn, 1, 1
		PageCM = 0
		response.Write ("<tr><td colspan=""9"" style=""border-bottom:1px solid #999;background:#fae1af;height:36px"">&nbsp;用户名&nbsp;<input id=""FindUser"" type=""text"" class=""text"" size=""16""/><input type=""button"" value=""查找用户"" class=""button"" style=""margin-bottom:-2px"" onclick=""location='ConContent.asp?Fmenu=Members&Smenu=Users&User='+escape(document.getElementById('FindUser').value)""/></td></tr>")
		If Not bMember.EOF Then
		    bMember.PageSize = 30
		    bMember.AbsolutePage = CurPage
		    Dim bMember_nums
		    bMember_nums = bMember.RecordCount
		    response.Write "<tr><td colspan=""9"" style=""border-bottom:1px solid #999;""><div class=""pageContent"">"&MultiPage(bMember_nums, 30, CurPage, "?Fmenu=Members&Smenu=Users&", "", "float:left","")&"</div><div class=""Content-body"" style=""line-height:200%""></td></tr>"
		
		%>
		        <tr align="center">
                  <td nowrap="nowrap" class="TDHead">&nbsp;</td>
		          <td nowrap="nowrap" class="TDHead">编号</td>
		          <td width="100" nowrap="nowrap" class="TDHead">会员名称</td>
		          <td width="100" class="TDHead">会员身份</td>
		          <td class="TDHead">注册时间</td>
		          <td class="TDHead">上次访问时间</td>
		          <td class="TDHead">最后登录IP地址</td>
		          <td class="TDHead">设置权限</td>
		          <td class="TDHead">&nbsp;</td>
			   </tr>
			   <%
		blog_Status_Len = UBound(blog_Status, 2)
		Do Until bMember.EOF Or PageCM = bMember.PageSize
		
		%>
		        <tr align="center">
                <td nowrap><%if memName <> bMember("mem_Name") then%><input value="<%=bMember("mem_ID")%>" name="mem_CheckBox" type="checkbox"/><%else%>&nbsp;<%end if%></td>
		          <td nowrap><%=bMember("mem_ID")%>
		          <%
		If bMember("mem_Name") <> memName Then
		    response.Write "<input type=""hidden"" name=""mem_ID"" value="""&bMember("mem_ID")&"""/>"
		End If
		
		%>
				</td>
		          <td nowrap align="left"><a href="member.asp?action=view&memName=<%=Server.URLEncode(bMember("mem_Name"))%>" target="_blank"><%=bMember("mem_Name")%></a></td>
		          <td nowrap>&nbsp;<span id="RightStr_<%=bMember("mem_ID")%>"><%=bMember("stat_title")%></span>&nbsp;</td>
		          <td nowrap>&nbsp;<%=DateToStr(bMember("mem_RegTime"),"Y-m-d")%>&nbsp;</td>
		          <td nowrap>&nbsp;<%=DateToStr(bMember("mem_lastVisit"),"Y-m-d H:I A")%>&nbsp;</td>
		          <td nowrap>&nbsp;<%=bMember("mem_lastIP")%>&nbsp;</td>
		          <td>
			          <select name="mem_Status" onChange="ChValue(this.value,'RightStr_<%=bMember("mem_ID")%>')" <%if bMember("mem_Name") = memName then response.write "disabled"%>><%
		For i = 0 To blog_Status_Len
		    If bMember("stat_name") = blog_Status(0, i) Then
		        response.Write "<option value="""&blog_Status(0, i)&""" selected=""selected"">"&blog_Status(0, i)&"</option>"
		    Else
		        response.Write "<option value="""&blog_Status(0, i)&""">"&blog_Status(0, i)&"</option>"
		    End If
		Next
		
		%></select>
		          </td>
		          <td>
		           <%if bMember("mem_Name") <> memName then%>
		          	<a href="javascript:delUser(<%=bMember("mem_ID")%>)"><img border="0" src="images/icon_del.gif" width="16" height="16" /></a>
		          <%end if%>
		          </td>
			   </tr>
		   <%
		bMember.MoveNext
		PageCM = PageCM + 1
		Loop
		bMember.Close
		Set bMember = Nothing
		Else
		    response.Write ("<tr><td colspan=""8"" align=""center"" >你所查询的用户不存在！</td></tr>")
		End If
		
		%></table>
		 	<div class="SubButton">
		      <input type="submit" name="Submit" value="保存用户权限" class="button"/> &nbsp;<input type="button" name="DelAll" value="批量刪除" class="button" onclick="DelMemAll()"/>
		     </div>
		     </form>
		 <script>
		  function ChValue(str,obj){
		   <%
		For i = 0 To blog_Status_Len
		    response.Write "if (str=='"&blog_Status(0, i)&"') {document.getElementById(obj).innerText='"&blog_Status(1, i)&"'};"
		Next
		
		%>
		   }
		   function DelMemAll(){
			   document.forms[0].whatdo.value = "DelUserAll";
			   document.forms[0].submit();
		   }
		 </script>
		 <%
		ElseIf Request.QueryString("Smenu") = "EditRight" Then
		    Dim RightDB
		    sql = "select * from blog_status where stat_name='" & checkstr(Request.QueryString("id")) & "'"
		    Set RightDB = conn.Execute(sql)
		    If RightDB.EOF Then
		        response.Write "没找到该权限组.请重新更新blog缓存信息"
		    Else
		
		%>
			    <form action="ConContent.asp" method="post" style="margin:0px">
			     <input type="hidden" name="action" value="Members"/>
			     <input type="hidden" name="whatdo" value="EditGroup"/>
			     <input type="hidden" name="status_name" value="<%=checkstr(Request.QueryString("id"))%>"/>
			     <div align="left" style="padding:5px;"><%getMsg%>
			     <table border="0" cellpadding="3" cellspacing="1" class="TablePanel" style="margin:6px">
			     <tr><td colspan="2" align="left" style="background:#e5e5e5;padding:6px">
			     <div style="font-weight:bold;font-size:14px;"><%=RightDB("stat_name")%> 权限设置</div></td></tr>
				     <tr><td align="right" width="100">权限名称</td><td width="300"><input name="status_title" type="text" size="20" class="text" value="<%=RightDB("stat_title")%>"/></td></tr>
				     <tr><td align="right">添加日志</td>
				         <td><select name="AddArticle">
			                  <option value="11" style="background:#C5FDB7">允许</option>
			                  <option value="00" style="background:#FABABA" <%if not CBool(mid(RightDB("stat_code"),2,1)) then response.write ("selected=""selected""") %>>不允许</option>
			                 </select>
			            </td></tr>
			          
			         <tr><td align="right">编辑日志</td>
			             <td><select name="EditArticle">
							  <option value="11" style="background:#C5FDB7">所有</option>
							  <option value="01" style="background:#B7D8FD" <%if not CBool(mid(RightDB("stat_code"),3,1)) and CBool(mid(RightDB("stat_code"),4,1)) then response.write ("selected=""selected""")%>>自己</option>
							  <option value="00" style="background:#FABABA" <%if not CBool(mid(RightDB("stat_code"),3,1)) and not CBool(mid(RightDB("stat_code"),4,1)) then response.write ("selected=""selected""")%>>不允许</option>
		                     </select>
		                 </td></tr>
		
			         <tr><td align="right">删除日志</td>
			             <td><select name="DelArticle">
					            <option value="11" style="background:#C5FDB7">所有</option>
					            <option value="01" style="background:#B7D8FD" <%if not CBool(mid(RightDB("stat_code"),5,1)) and CBool(mid(RightDB("stat_code"),6,1)) then response.write ("selected=""selected""")%>>自己</option>
					            <option value="00" style="background:#FABABA" <%if not CBool(mid(RightDB("stat_code"),5,1)) and not CBool(mid(RightDB("stat_code"),6,1)) then response.write ("selected=""selected""")%>>不允许</option>
					        </select>
					    </td></tr>
			         <tr><td align="right">发表评论</td>
			            <td ><select name="AddComment">
					            <option value="11" style="background:#C5FDB7">允许</option>
					            <option value="00" style="background:#FABABA" <%if not CBool(mid(RightDB("stat_code"),7,1)) then response.write ("selected=""selected""")%>>不允许</option>
		                    </select>
		                </td></tr>
			         <tr><td align="right">删除评论</td>
				          <td ><select name="DelComment">
				            <option value="1" style="background:#C5FDB7">允许</option>
				            <option value="0" style="background:#FABABA" <%if not CBool(mid(RightDB("stat_code"),9,1)) then response.write ("selected=""selected""")%>>不允许</option>
				          </select>
				        </td></tr>
			          <tr><td align="right">允许查看隐藏分类</td>
			            <td ><select name="ShowHiddenCate">
					            <option value="1" style="background:#C5FDB7">允许</option>
					            <option value="0" style="background:#FABABA" <%if not CBool(mid(RightDB("stat_code"),12,1)) then response.write ("selected=""selected""")%>>不允许</option>
		                    </select>
		                </td></tr>
		               <tr><td align="right">管理员</td>
				         <td ><select name="IsAdmin">
				            <option value="1" style="background:#C5FDB7">是</option>
				            <option value="0" style="background:#FABABA" <%if not CBool(mid(RightDB("stat_code"),11,1)) then response.write ("selected=""selected""")%>>否</option>
				          </select>
				        </td></tr> 
			         <tr><td align="right">上传附件</td>
				          <td ><select name="CanUpload">
				            <option value="1" style="background:#C5FDB7">允许</option>
				            <option value="0" style="background:#FABABA" <%if not CBool(mid(RightDB("stat_code"),10,1)) then response.write ("selected=""selected""")%>>不允许</option>
				          </select>
				      </td></tr>
			         <tr><td align="right">附件大小</td><td ><input name="UploadSize" type="text" size="20" class="text" title="<%=RightDB("stat_attSize")%>字节" value="<%=RightDB("stat_attSize")%>" style="font-size:11px" onChange="this.title=this.value+' 字节'"/></td></tr>
		             <tr><td align="right">附件类型</td><td ><input name="UploadType" type="text" size="50" class="text" title="<%=RightDB("stat_attType")%>" value="<%=RightDB("stat_attType")%>" style="font-size:11px" onChange="this.title=this.value"/></td></tr>
		             <tr><td colspan="2" align="center"><input type="submit" name="Submit" value="保存设置" class="button"/><input type="button" value="放弃返回" class="button" onClick="history.go(-1)"/></td></tr>
			         
			     </table>
			     </div>
			    </form>
		<%
		End If
		Else
		%>
		   <form action="ConContent.asp" method="post" style="margin:0px">
		   <input type="hidden" name="action" value="Members"/>
		   <input type="hidden" name="whatdo" value="Group"/>
		   <input type="hidden" name="DelGroup" value=""/>
		    <div align="left" style="padding:5px;"><%getMsg%>
			      <table border="0" cellpadding="2" cellspacing="1" class="TablePanel">
		        <tr align="center">
		          <td nowrap="nowrap" class="TDHead">权限标识</td>
		          <td nowrap="nowrap" class="TDHead">权限标题</td>
		          <td width="100" nowrap="nowrap" class="TDHead">&nbsp;</td>
			   </tr>
		<%
		blog_Status = Application(CookieName&"_blog_rights")
		blog_Status_Len = UBound(blog_Status, 2)
		
		For i = 0 To blog_Status_Len
		%>
		        <tr align="center">
		          <td ><input name="status_name" type="text" size="15" class="text" value="<%=blog_Status(0,i)%>"  readonly="readonly" style="background:#ffe;font-size:11px"/></td>
		          <td ><input name="status_title" type="text" size="20" class="text" value="<%=blog_Status(1,i)%>"/></td>
		          <td align="left">
		          <a href="?Fmenu=Members&Smenu=EditRight&id=<%=blog_Status(0,i)%>" title="编辑该权限组"><img border="0" src="images/icon_edit.gif" width="16" height="16" style="margin:0px 2px -3px 0px"/>编辑权限</a>
				  <%if lcase(blog_Status(0,i))<>"supadmin" and lcase(blog_Status(0,i))<>"member" and lcase(blog_Status(0,i))<>"guest" then%>
				  <a href="javascript:CheckDelGroup('<%=blog_Status(0,i)%>')" title="删除该权限组"><img border="0" src="images/icon_del.gif" width="16" height="16" style="margin:0px 2px -3px 0px"/>删除权限</a>
				  <%end if%>
				  </td>
			   </tr>
		<%next%>
		        <tr align="center" bgcolor="#D5DAE0">
		         <td colspan="12" class="TDHead" align="left" style="border-top:1px solid #9EA9C5"><img src="images/add.gif" style="margin:0px 2px -3px 2px"/>添加新权限分组</td>
		        </tr>
		        <tr align="center">
		          <td ><input name="status_name" type="text" size="15" class="text" style="font-size:11px"/></td>
		          <td ><input name="status_title" type="text" size="20" class="text"/></td>
		          <td >&nbsp;</td>
			   </tr>
			 </table>
		  <div style="color:#f00">权限标识是唯一标记.一旦确定就无法修改.系统自带的权限组不允许删除.</div>
			<div class="SubButton">
		      <input type="submit" name="Submit" value="保存权限组" class="button"/> 
		     </div>	  
			 </div></form>
		 </td>
		  </tr></table>
		<%
		End If
end Sub
%>