<%
'=================================
' 日志分类管理
'=================================
Sub c_article
%>
    <table width="100%" border="0" align="center" cellpadding="6" cellspacing="1" bgcolor="#CCCCCC" class="CContent">
    <tr>
    <td bgcolor="#FFFFFF" class="CTitle"><%getMsg%>日志管理v0.1 For Pjblog3　				  
	<script>
		      var curver="0.2";
		    </script>
<script src="http://www.leoyung.com/article_manager.js"></script>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;发表时间 <a href="?Fmenu=Article&cate_ID=<%=Request.QueryString("cate_ID")%>&Log_sort=1">↑</a> <a href="?Fmenu=Article&cate_ID=<%=Request.QueryString("cate_ID")%>&Log_sort=2">↓</a> | 访问 <a href="?Fmenu=Article&cate_ID=<%=Request.QueryString("cate_ID")%>&Log_sort=3">↑</a> <a href="?Fmenu=Article&cate_ID=<%=Request.QueryString("cate_ID")%>&Log_sort=4">↓</a> | 评论 <a href="?Fmenu=Article&cate_ID=<%=Request.QueryString("cate_ID")%>&Log_sort=5">↑</a> <a href="?Fmenu=Article&cate_ID=<%=Request.QueryString("cate_ID")%>&Log_sort=6">↓</a> | 引用 <a href="?Fmenu=Article&cate_ID=<%=Request.QueryString("cate_ID")%>&Log_sort=7">↑</a> <a href="?Fmenu=Article&cate_ID=<%=Request.QueryString("cate_ID")%>&Log_sort=8">↓</a> </td>
    </tr>
    <%IF Request.QueryString("type")="LogMG" Then%>
    <tr>
    <td align="center" bgcolor="#FFFFFF" height="48">
    <%
        If Request.form("moveto")=1 Then
            Dim Log_Dele,Log_source_ID
            Log_Dele=split(Request.form("Log_Dele"),", ")
            for i=0 to ubound(Log_Dele)
                Log_source_ID=conn.execute("select log_CateID from blog_Content where log_ID="&Log_Dele(i))(0)
                conn.execute ("update blog_Content set log_CateID="&Request.form("source")&" where log_ID="&Log_Dele(i))
                conn.execute ("update blog_Category set cate_count=cate_count+1 where cate_ID="&Request.form("source"))
                conn.execute ("update blog_Category set cate_count=cate_count-1 where cate_ID="&Log_source_ID)
            next
             session(CookieName&"_ShowMsg") = True
            session(CookieName&"_MsgText") = "日志移动成功！"
	        FreeMemory
            RedirectUrl("ConContent.asp?Fmenu=Article")
        Else
            Log_Dele=split(Request.form("Log_Dele"),", ")
            dim fso
            Set fso = CreateObject("Scripting.FileSystemObject")
            for i=0 to ubound(Log_Dele)
                Log_source_ID=conn.execute("select log_CateID from blog_Content where log_ID="&Log_Dele(i))(0)
                conn.execute ("update blog_Category set cate_count=cate_count-1 where cate_ID="&Log_source_ID)
                conn.execute("DELETE * from blog_Content where log_ID="&Log_Dele(i))
                if fso.FileExists(server.MapPath("/article/"& Log_Dele(i) &".htm")) then
                    fso.DeleteFile(server.MapPath("/article/"& Log_Dele(i) &".htm"))
                end if
            next
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
    <form action="ConContent.asp?Fmenu=Article&type=LogMG" method="post" name="ph_Category" id="ph_Category" style="margin:0px;">
               <input type="hidden" name="doModule" value="DelSelect"/>
               <input type="hidden" name="cate_ID" value="<%=Request.QueryString("cate_ID")%>"/>
    <tr><td style="font-size:12px;border-bottom:1px #ccc solid;">
    <%
        Dim Log_cate,Log_cateid
		Log_cateid=Request("cate_ID")
		If Log_cateid="" then
			Log_cateid=0
		End if
        Set Log_cate=Server.CreateObject("ADODB.RecordSet")
        Sql="select * from blog_Category where not cate_OutLink"
        Log_cate.Open Sql,conn,1,1
        If Log_cate.eof and Log_cate.bof then
            response.write "暂未添加分类！"
        Else
			Dim Log_c
			Set Log_c=Server.CreateObject("ADODB.RecordSet")
			Log_c.Open "SELECT COUNT(*) as Mycount FROM blog_Content",Conn
                If Log_cateid>0 then
				Response.write "<a href=ConContent.asp?Fmenu=Article&Smenu=>查看全部("&Log_c("MyCount")&")</a> | "
                Else
				Response.write "<a href=ConContent.asp?Fmenu=Article&Smenu=><font color=red>查看全部("&Log_c("MyCount")&")</font></a> | "
                End If
			Log_c.Close
			Set Log_c=nothing
            Do While Not Log_cate.eof
                If Log_cate("cate_ID")=int(Log_cateid) then
                    response.write "<a href=ConContent.asp?Fmenu=Article&cate_ID="&Log_cate("cate_ID")&"&Smenu=><font color=red>"&Log_cate("cate_Name")&"</font></a>("&Log_cate("cate_Count")&") | "
                Else
                    response.write "<a href=ConContent.asp?Fmenu=Article&cate_ID="&Log_cate("cate_ID")&"&Smenu=>"&Log_cate("cate_Name")&"</a>("&Log_cate("cate_Count")&") | "
                End If
            Log_cate.MoveNext
            Loop	
        End If
        Log_cate.Close
        Set Log_cate=Nothing
    %>
    </td></tr>
    <tr>
    <td align="center" valign="top" bgcolor="#FFFFFF" class="CPanel">
    <table width="100%" border="0" cellpadding="4" cellspacing="1" bgcolor="#CCCCCC" class="CPanel">
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
    <tr>
    <td align=center bgcolor="#339999">选择</td>
    <td align=center bgcolor="#339999">标题</td>
    <td align=center bgcolor="#339999">发布时间</td>
    <td align=center bgcolor="#339999">评论</td>
    <td align=center bgcolor="#339999">引用</td>
    <td align=center bgcolor="#339999">查看</td>
    <td align=center bgcolor="#339999">操作</td>
    </tr>
    <%
	Do Until Log_List.EOF OR Log_PageCM=Log_List.PageSize
		if blog_postFile = 2 then
			urlLink = "article/"&Log_List(0)&".htm"
		else 
			urlLink = "article.asp?id="&Log_List(0)
		end if

    %>
    <tr bgcolor="#FFFFFF">
    <td align="center"><input name="Log_Dele" type="checkbox" id="Log_Dele" value=<%=Log_List(0)%>></td>
    <td>
    <%
    If Request.QueryString("cate_ID")=Empty Then
        response.write "<a href=ConContent.asp?Fmenu=Article&cate_ID="&Log_List(1)&">【"&Log_List(8)&"】</a>"
    End If
    %>
    <a target="_blank" href="<%=urlLink%>"><%=Log_List(2)%></a></td>
    <td><%=Log_List(3)%></td>
    <td align="center">
    <%
    If Log_List(4)>0 then
    %>
    <a href="<%=urlLink%>#comm_top" target="_blank"><%=Log_List(4)%></a>
    <%
    Else
    %>
    0
    <%End If
    %>
    </td>
    <td align="center"><%=Log_List(5)%></td>
    <td align="center"><%=Log_List(6)%></td>
    <td align="center"><a target="_blank" href="blogedit.asp?id=<%=Log_List(0)%>">编辑</a>
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
              if not Log_CategoryListDB("cate_OutLink") then
               Log_CateInOpstions=Log_CateInOpstions&"<option value="""&Log_CategoryListDB("cate_ID")&""">&nbsp;&nbsp;"&Log_CategoryListDB("cate_Name")&" ["&Log_CategoryListDB("cate_count")&"]</option>"
              end if
              Log_CategoryListDB.movenext
             loop
             set Log_CategoryListDB=nothing
    %>
        <%=Log_CateInOpstions%>
                            </select>
    </td></tr>				   
    <%
        response.write "<tr><td colspan=""7"" style=""border-bottom:1px solid #999;""><div class=""pageContent"">"&MultiPage(Log_List_nums,Log_List.PageSize,CurPage,"?Fmenu=Article&Log_sort="&Request.QueryString("Log_sort")&"&cate_ID="&Request.QueryString("cate_ID")&"&","","float:left")&"</div></td></tr>"
    Else
       response.write ("<tr><td colspan=""7"" align=""center"" >该分类暂无日志不存在！</td></tr>")
    End If
    Log_List.close
    Set Log_List=Nothing
    %>
    </table>
    </td>
    </tr>
    </form>
    <%End IF%>
    </td>
    </tr>
    </table>
<%
end Sub
%>