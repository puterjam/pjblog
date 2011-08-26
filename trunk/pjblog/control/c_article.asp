<%
'=================================
' 日志分类管理
'=================================
Sub c_article
%>
<%getMsg%>
    <table width="100%" border="0" align="center" cellpadding="6" cellspacing="1" bgcolor="#CCCCCC" class="CContent">
	    <tr>
		    <td bgcolor="#FFFFFF" class="CTitle"><b>日志管理</b></td>
   		</tr>
	    <%IF Request.QueryString("type")="LogMG" Then%>
	    <tr>
		    <td align="center" bgcolor="#FFFFFF" height="48">
		    <%
		        Dim Log_Dele, Log_source_ID, fso, logTag, WebFso
		        If Request.form("moveto") = 1 Then
		            Log_Dele = Split(Request.form("Log_Dele"),", ")
		            For i = 0 To UBound(Log_Dele)
						If blog_postFile = 2 Then
							Dim p_targFolder, p_file, p_filename
							p_targFolder = Conn.Execute("Select cate_Part From blog_Category Where cate_ID=" & CheckStr(Request.form("source")))(0)
							p_file = Alias(Log_Dele(i))
							p_file = Replace(p_file, "\", "/")
							p_filename = Split(p_file, "/")(Ubound(Split(p_file, "/")))
							Set WebFso = New cls_FSO
							If Len(p_targFolder) = 0 Then
								WebFso.MoveFile Alias(Log_Dele(i)), "article/" & p_filename
							Else
								WebFso.MoveFile Alias(Log_Dele(i)), "article/" & p_targFolder & "/" & p_filename
							End If
							Set WebFso = Nothing
						End if
		                Log_source_ID = conn.execute("select log_CateID from blog_Content where log_ID="&Log_Dele(i))(0)
		                conn.execute ("update blog_Content set log_CateID="&CheckStr(Request.form("source"))&" where log_ID="&Log_Dele(i))
		                conn.execute ("update blog_Category set cate_count=cate_count+1 where cate_ID="&CheckStr(Request.form("source")))
		                conn.execute ("update blog_Category set cate_count=cate_count-1 where cate_ID="&Log_source_ID)
						PostArticle Log_Dele(i), False
		            Next
		             session(CookieName&"_ShowMsg") = True
		            session(CookieName&"_MsgText") = "日志移动成功！"
			        FreeMemory
		            RedirectUrl("ConContent.asp?Fmenu=Article")
		        Else
		            Log_Dele=split(Request.form("Log_Dele"),", ")
		            Set fso = CreateObject("Scripting.FileSystemObject")
					Set WebFso = New cls_FSO
		            for i=0 to ubound(Log_Dele)
		                Log_source_ID = conn.execute("select log_CateID from blog_Content where log_ID="&Log_Dele(i))(0)
						logTag = conn.execute("select log_Tag from blog_Content where log_ID="&Log_Dele(i))(0)
						If fso.FileExists(server.MapPath(Alias(Log_Dele(i)))) Then
		                    WebFso.DeleteFile(Alias(Log_Dele(i)))
							WebFso.DeleteFile("cache/"&Log_Dele(i)&".asp")
		                End If
		                conn.execute ("update blog_Category set cate_count=cate_count-1 where cate_ID="&Log_source_ID)
						conn.execute ("update blog_Info set blog_LogNums=blog_LogNums-1 where blog_ID=1")
		                conn.execute("DELETE * from blog_Content where log_ID="&Log_Dele(i))
		                autoDeleteTag logTag
		            Next
					Set WebFso = Nothing
		             session(CookieName&"_ShowMsg") = True
		            session(CookieName&"_MsgText") = "日志删除成功！"
			        FreeMemory
					Session(CookieName&"_LastDo")="DelArticle"
		            RedirectUrl("ConContent.asp?Fmenu=Article&cate_ID="&Request.form("cate_ID")&"")
		       End If
		    %>
		    </td>
	    </tr>
    <%Else%>
	    
	    <tr>
		    <td class="CPanel">
		    <div class="SubMenu">
		    <%
		        Dim Log_cate,Log_cateid
				Log_cateid=Request("cate_ID")
				If Log_cateid="" Then
					Log_cateid=0
				End If
		        Set Log_cate=Server.CreateObject("ADODB.RecordSet")
		        Sql="select * from blog_Category where not cate_OutLink"
		        Log_cate.Open Sql,conn,1,1
		        If Log_cate.eof and Log_cate.bof Then
		            response.write "暂未添加分类！"
		        Else
					Dim Log_c
					Set Log_c=Server.CreateObject("ADODB.RecordSet")
					Log_c.Open "SELECT COUNT(*) as Mycount FROM blog_Content",Conn
		                If Log_cateid>0 Then
							Response.write "<a href=ConContent.asp?Fmenu=Article&Smenu=>查看全部("&Log_c("MyCount")&")</a> | "
		                Else
							Response.write "<a href=ConContent.asp?Fmenu=Article&Smenu=><b>查看全部("&Log_c("MyCount")&")</b></a> | "
		                End If
					Log_c.Close
					Set Log_c=nothing
		            Do While Not Log_cate.eof
		                If Log_cate("cate_ID")=int(Log_cateid) Then
		                    response.write "<a href='ConContent.asp?Fmenu=Article&cate_ID="&Log_cate("cate_ID")&"&Smenu='><b>"&Log_cate("cate_Name")&"</b></a>("&Log_cate("cate_Count")&") | "
		                Else
		                    response.write "<a href='ConContent.asp?Fmenu=Article&cate_ID="&Log_cate("cate_ID")&"&Smenu='>"&Log_cate("cate_Name")&"</a>("&Log_cate("cate_Count")&") | "
		                End If
		            Log_cate.MoveNext
		            Loop	
		        End If
		        Log_cate.Close
		        Set Log_cate=Nothing
		    %></div>
		    
		    </td>
		  </tr>
		    <tr>
			    <td valign="top" class="CPanel" style="padding-left:20px;padding-right:20px">
				<div style="float:right;margin-right:6px;">
					 <b>排序:</b> 发表时间 <a href="?Fmenu=Article&cate_ID=<%=Request.QueryString("cate_ID")%>&Log_sort=1"><img src="images/sort_up.gif" alt="正序排列" border=0/></a>
					 <a href="?Fmenu=Article&cate_ID=<%=Request.QueryString("cate_ID")%>&Log_sort=2"><img src="images/sort_down.gif" alt="倒序排列" border=0/></a>
					 <span style="color:#999">|</span> 评论 <a href="?Fmenu=Article&cate_ID=<%=Request.QueryString("cate_ID")%>&Log_sort=5"><img src="images/sort_up.gif" alt="正序排列" border=0/></a>
					 <a href="?Fmenu=Article&cate_ID=<%=Request.QueryString("cate_ID")%>&Log_sort=6"><img src="images/sort_down.gif" alt="倒序排列" border=0/></a>
					 <span style="color:#999">|</span> 引用 <a href="?Fmenu=Article&cate_ID=<%=Request.QueryString("cate_ID")%>&Log_sort=7"><img src="images/sort_up.gif" alt="正序排列" border=0/></a> 
					 <a href="?Fmenu=Article&cate_ID=<%=Request.QueryString("cate_ID")%>&Log_sort=8"><img src="images/sort_down.gif" alt="倒序排列" border=0/></a> 
					 <span style="color:#999">|</span> 访问 <a href="?Fmenu=Article&cate_ID=<%=Request.QueryString("cate_ID")%>&Log_sort=3"><img src="images/sort_up.gif" alt="正序排列" border=0/></a>
					 <a href="?Fmenu=Article&cate_ID=<%=Request.QueryString("cate_ID")%>&Log_sort=4"><img src="images/sort_down.gif" alt="倒序排列" border=0/></a>
				</div>
			    <div style="text-align:left;margin-bottom:5px;"><a href="blogpost.asp" target="_blank"><img style="margin: 0px 2px -3px 0px" src="images/add.gif" border="0" /><b>发表新日志</b></a></div>
			    
					<form action="ConContent.asp?Fmenu=Article&type=LogMG" method="post" name="ph_Category" id="ph_Category" style="margin:0px;">
			               <input type="hidden" name="doModule" value="DelSelect"/>
			               <input type="hidden" name="cate_ID" value="<%=Request.QueryString("cate_ID")%>"/>
				    <%
				        If CheckStr(Request.QueryString("Page"))<>Empty Then
				            Curpage=CheckStr(Request.QueryString("Page"))
				            If IsInteger(Curpage)=False OR Curpage<0 Then Curpage=1
				        Else
				            Curpage=1
				        End If
				    
				    Dim Log_List
				    Set Log_List=Server.CreateObject("ADODB.RecordSet")
				    
				    If Request.QueryString("cate_ID")<>Empty Then
				        Sql="select log_ID,log_CateID,log_Title,log_PostTime,log_CommNums,log_QuoteNums,log_ViewNums,cate_ID,cate_Name from blog_Content c inner join blog_Category l on c.log_CateID=l.cate_ID Where log_CateID="&Request.QueryString("cate_ID")&""
				    Else
				        Sql="select log_ID,log_CateID,log_Title,log_PostTime,log_CommNums,log_QuoteNums,log_ViewNums,cate_ID,cate_Name from blog_Content c inner join blog_Category l on c.log_CateID=l.cate_ID"
				    End If
				    If Request.QueryString("Log_sort")<>Empty Then
				        Select Case Request.QueryString("Log_sort")
				            Case 1
				                Sql=Sql&" order by log_PostTime"
				            Case 2
				                Sql=Sql&" order by log_PostTime desc"
				            Case 3
				                Sql=Sql&" order by log_ViewNums"
				            Case 4
				                Sql=Sql&" order by log_ViewNums desc"
				            Case 5
				                Sql=Sql&" order by log_CommNums"
				            Case 6
				                Sql=Sql&" order by log_CommNums desc"
				            Case 7
				                Sql=Sql&" order by log_QuoteNums"
				            Case 8
				                Sql=Sql&" order by log_QuoteNums desc"
				        End Select
				    Else
				        Sql=Sql&" order by log_ID desc"	
				    End If
				    
				    Log_List.Open Sql,conn,1,1
				    
				    If not Log_List.eof Then
				        Dim Log_PageCM
				        Log_PageCM=0
				        Log_List.PageSize=15
				        Log_List.AbsolutePage=CurPage
				        Dim Log_List_nums
				        Log_List_nums=Log_List.RecordCount
						Dim urlLink    '根据动静态判断输出连接  JieLiao
				    %>
				    <table border="0" cellpadding="2" cellspacing="1" class="TablePanel">
				    <tr align="center">
					    <td class="TDHead" nowrap>&nbsp;</td>
					    <%If Request.QueryString("cate_ID")=Empty Then%>
					    <td class="TDHead" nowrap>日志分类</td>
					    <%End If%>
					    <td class="TDHead" width="100%">日志标题</td>
					    <td class="TDHead" nowrap>发表时间</td>
					    <td class="TDHead" nowrap>评论</td>
					    <td class="TDHead" nowrap>引用</td>
					    <td class="TDHead" nowrap>访问</td>
					    <td class="TDHead" nowrap>操作</td>
				    </tr>
				    <%
					Do Until Log_List.EOF OR Log_PageCM=Log_List.PageSize
						If blog_postFile = 2 Then
							urlLink = Alias(Log_List(0))
						Else 
							urlLink = "article.asp?id="&Log_List(0)
						End If

				    %>
				    <tr bgcolor="#FFFFFF">
				    <td align="center"><input name="Log_Dele" type="checkbox" id="Log_Dele" value=<%=Log_List(0)%>></td>
				    <%If Request.QueryString("cate_ID")=Empty Then%>
				    <td><nobr><a href=ConContent.asp?Fmenu=Article&cate_ID=<%=Log_List(1)%>>【<%=Log_List(8)%>】</a></nobr></td>
				    <%End If%>
				    <td><a target="_blank" href="<%=urlLink%>"><%=Log_List(2)%></a></td>
				    <td><nobr><%=DateToStr(Log_List(3),"Y-m-d H:I:S")%></nobr></td>
				    <td align="center">
				    <%
				    If Log_List(4)>0 Then
				    %>
				    <a href="<%=urlLink%>#comm_top" target="_blank"><%=Log_List(4)%></a>
				    <%
				    Else
				    %>
				    0
				    <%End If
				    %>
				    </td>
				    <td align="center"><nobr><%=Log_List(5)%></nobr></td>
				    <td align="center"><nobr><%=Log_List(6)%></nobr></td>
				    <td align="center"><nobr><a target="_blank" href="blogedit.asp?id=<%=Log_List(0)%>"><img style="margin: 0px 2px -3px 0px" src="images/icon_edit.gif" border="0" />编辑</a></nobr>
				    </select>
				    </td>
				    </tr>
				    <%
				    Log_List.MoveNext
				    Log_PageCM=Log_PageCM+1
				    Loop
				    %>
				    <tr><td colspan="7" bgcolor="#ffffff">
			            <input type="button" value="全选" onClick="checkAll()" class="button" style="margin:0px;margin-bottom:5px;margin-right:6px"/>
			            <input type="button" value="删除所选内容" onClick="DelComment()" class="button" style="margin:0px;margin-bottom:5px;"/>
			            <input type="hidden" value="0" name="moveto">
			           <input type="submit" value="将所选内容移至" onClick="moveto.value=1" class="button" style="margin:0px;margin-bottom:5px;"/>
			           <select name="source"  style="margin:0px;margin-bottom:5px;">
					    <%
					    Dim Log_CategoryListDB,Log_CateInOpstions
					            set Log_CategoryListDB=conn.execute("select * from blog_Category order by cate_local asc, cate_Order desc")
					             do while not Log_CategoryListDB.eof
					              If not Log_CategoryListDB("cate_OutLink") Then
					               Log_CateInOpstions=Log_CateInOpstions&"<option value="""&Log_CategoryListDB("cate_ID")&""">&nbsp;&nbsp;"&Log_CategoryListDB("cate_Name")&" ["&Log_CategoryListDB("cate_count")&"]</option>"
					              End If
					              Log_CategoryListDB.movenext
					             loop
					             set Log_CategoryListDB=nothing
					    %>
					        <%=Log_CateInOpstions%>
					       </select>
				    </td></tr>				   
			   	 	</table>

    	    <%
	        response.write "<div class=""pageContent"">"&MultiPage(Log_List_nums,Log_List.PageSize,CurPage,"?Fmenu=Article&Log_sort="&Request.QueryString("Log_sort")&"&cate_ID="&Request.QueryString("cate_ID")&"&","","float:left")&"</div>"
	    Else
	        response.write ("该分类暂无日志不存在！")
	    End If
	    Log_List.close
	    Set Log_List=Nothing
	    %>
    <%End IF%>
    </form>
   		 </td>
   	 </tr>
    </table>
<%
End Sub
%>