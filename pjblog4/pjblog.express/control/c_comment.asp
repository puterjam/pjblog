<%
Public Sub c_Comment
	Dim Aduit, Reply, SqlStr, Sql, AduitCons, ReplyCons, i
	Dim Rs, CommentArray, PageLeft, PageRight, LenCount, PageSize, PageStr, page
	PageSize = 10
	page = Asp.CheckStr(Request.QueryString("page"))
	If Not Asp.IsInteger(page) Then page = 1
	Aduit = Session(Sys.CookieName & "Aduit")
	If Asp.IsBlank(Aduit) Then
		Aduit = False
		AduitCons = ""
	Else
		If Not Asp.IsInteger(Aduit) Then Aduit = 0
		Aduit = CBool(Aduit)
		If Aduit Then
			AduitCons = "Where comm_IsAudit=True"
		Else
			AduitCons = "Where comm_IsAudit=False"
		End If
	End If
	
	
	Reply = Session(Sys.CookieName & "Reply")
	If Asp.IsBlank(Reply) Then
		Reply = False
		ReplyCons = ""
	Else
		If Not Asp.IsInteger(Reply) Then Reply = 0
		Reply = CBool(Reply)
		If Reply Then
			ReplyCons = "And replay_content<>''"
		Else
			ReplyCons = "And replay_content=''"
		End If
	End If
	
	SqlStr = "comm_ID, blog_ID, comm_Content, comm_Author, comm_PostTime, comm_PostIP, comm_IsAudit, comm_Email, comm_WebSite, reply_author, reply_time, replay_content"
	Sql = "Select " & SqlStr & " From blog_Comment " & AduitCons & " " & ReplyCons & " Order By comm_PostTime Desc"
	
	Set Rs = Server.CreateObject("Adodb.RecordSet")
	Rs.open Sql, Conn, 1, 1
	If Rs.Bof Or Rs.Eof Then
		ReDim CommentArray(0, 0)
	Else
		CommentArray = Rs.GetRows
	End If
	Rs.Close
	Set Rs = Nothing
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
	<tr>
    	<th class="CTitle"><%=control.categoryTitle%></th>
  	</tr>
    <tr><td>
    <div style="margin:20px">
    	
        <div style="border:1px solid #dfdfdf; float:right; width:220px;">
        	<div style="background:url(../images/Control/gray-grad.png) repeat-x; padding:5px 10px 5px 10px;">筛选条件</div>
            <div style="padding:5px 10px 10px 10px; line-height:25px">1. 根据你的审核情况显示评论模式<br />
            	<select style="width:200px">
                	<option value="0" <%If Not Aduit Then Response.Write("selected")%>>未审核评论</option>
                    <option value="1" <%If Aduit Then Response.Write("selected")%>>已审核评论</option>
                </select><br />
               2. 根据你的回复情况显示评论模式<br />
               	<select style=" width:200px">
                	<option value="0" <%If Not Reply Then Response.Write("selected")%>>未回复评论</option>
                    <option value="1" <%If Reply Then Response.Write("selected")%>>已回复评论</option>
                </select><br />
               3. 某日志评论 <input type="text" value="" size="1" class="text"> 确定(日志ID)<br />
               4. 还原显示默认设置
            </div>
        </div>
        
    	<div style="margin-right:245px">
    	<table cellpadding="3" cellspacing="0" class="CeeTable" width="100%">
    	<thead>
        	<tr>
            	<th width="50"><input type="checkbox"></th>
                <th>描述</th>
                <th width="250">&nbsp;</th>
            </tr>
        </thead>
        <tfoot>
        	<tr>
            	<th width="50"><input type="checkbox"></th>
                <th>描述</th>
                <th width="250">&nbsp;</th>
            </tr>
        </tfoot>
        <tbody>
<%
	If UBound(CommentArray, 1) = 0 Then
		Response.Write("<tr><td colspan=""3"" align=""center"">没有找到信息</td></tr>")
	Else
		LenCount = UBound(CommentArray, 2)
		PageLeft = (page - 1) * PageSize
		PageRight = page * PageSize - 1
		If PageLeft < 0 Then PageLeft = 0
		If PageRight > LenCount Then PageRight = LenCount
		For i = PageLeft To PageRight
		'SqlStr = "comm_ID, blog_ID, comm_Content, comm_Author, comm_PostTime, comm_PostIP, comm_IsAudit, comm_Email, comm_WebSite, reply_author, reply_time, replay_content"
		'  				0		1			2			3			4				5			6				7			8
		'9		10				11
%>
        	<tr class="active">
            	<td style="padding-left:8px"><input type="checkbox" value="" name=""></td>
                <td><%=CommentArray(3, i)%></td>
                <td width="250"><%=CommentArray(5, i)%> | <%=Asp.DateToStr(CommentArray(4, i), "Y-m-d H:I:S")%></td>
            </tr>
            <tr class="active" id="comcontent_<%=CommentArray(0, i)%>">
            	<td style="padding-left:8px">&nbsp;</td>
                <td><%=CommentArray(2, i)%></td>
                <td width="250">删除 | <a href="javascript:ceeevio.ManComm.Reply.ReplyBox(<%=CommentArray(0, i)%>);" class="activeButton">回复</a> | 审核 | 查看原文</td>
            </tr>
            <tr class="active second">
            	<td style="padding-left:8px">&nbsp;</td>
                <td>
                	<%If Not Asp.IsBlank(CommentArray(8, i)) Then%>&lt; <a href="<%=CommentArray(8, i)%>"><%=CommentArray(8, i)%></a> &gt;<%End If%>
					<%If Not Asp.IsBlank(CommentArray(7, i)) Then%>&lt; <a href="mailto:<%=CommentArray(7, i)%>"><%=CommentArray(7, i)%></a> &gt;<%End If%>
                </td>
                <td width="250">&nbsp;</td>
            </tr>
<%
		Next
	End If
%>
        </tbody>
      </table>
      </div>
      
    </div>
    </td></tr>
</table>
<%
End Sub
%>