<%
Public Sub c_link
	Dim Rs, Sql, LinkArray, i
	Dim linkLeft, linkRight, linkLen, linkPageStr
	linkPageStr = ""
	Dim CateID, IsShow, IsMain, Order, PerPage, page, PageCount
	
	CateID = Session(Sys.CookieName & "CateID") : If Not Asp.IsInteger(CateID) Then CateID = "" Else CateID = "And Link_ClassID=" & CateID
	IsShow = Session(Sys.CookieName & "IsShow") : IsMain = Session(Sys.CookieName & "IsMain") : Order = Session(Sys.CookieName & "Order")
	PerPage = Session(Sys.CookieName & "PerPage") : If Not Asp.IsInteger(PerPage) Then PerPage = 10
	page = Asp.CheckStr(Request.QueryString("page")) : If Not Asp.IsInteger(page) Then page = 1
	If Not Asp.IsInteger(IsShow) Then
		IsShow = ""
	Else
		If CBool(IsShow) Then
			IsShow = "And link_IsShow=true"
		Else
			IsShow = "And link_IsShow=false"
		End If
	End If
	
	If Not Asp.IsInteger(IsMain) Then
		IsMain = ""
	Else
		If CBool(IsMain) Then
			IsMain = "And link_IsMain=true"
		Else
			IsMain = "And link_IsMain=false"
		End If
	End If
	
	If Not Asp.IsInteger(Order) Then
		Order = ""
	Else
		If CBool(Order) Then
			Order = ",link_Order desc"
		Else
			Order = ",link_Order asc"
		End If
	End If
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
	<tr>
    	<th class="CTitle"><%=control.categoryTitle%></th>
  	</tr>
    <tr><td>
    <div style="width:220px; float:left">
    	<!--分类筛选-->
    	<div style=" margin:10px; border:1px solid #dfdfdf;">
        	<div style="background:url(../images/Control/gray-grad.png) repeat-x; padding:5px 10px 5px 10px; text-align:right"><span style="float:left; font-weight:bold">分类筛选</span><a href="javascript:ceeevio.Link.AddLinkClass.Box();" class="newadd">添加</a></div>
            <div style="line-height:20px; padding:5px 10px 5px 10px;">已有分类 (你可以添加分类)<br />
            	<select style="width:175px;" onchange="ceeevio.Link.SetSession('CateID', jQuery(this).val())">
                	<option value="">不筛选分类(默认)</option>
                    <%=CateSelect(Session(Sys.CookieName & "CateID"))%>
                </select><br />
			</div>
        </div>
        <script language="javascript">
			var CateString = '<%=CateSelect("")%>';
		</script>
        <!--其他筛选-->
        <div style=" margin:10px; border:1px solid #dfdfdf;">
        	<div style="background:url(../images/Control/gray-grad.png) repeat-x; padding:5px 10px 5px 10px;"><span style="font-weight:bold">其他筛选</span></div>
            <div style="line-height:20px; padding:5px 10px 5px 10px;">你可以根据不同条件来进行筛选<br />
            	<div>1. 每页链接数 <input type="text" value="<%If Not Asp.IsInteger(Session(Sys.CookieName & "PerPage")) Then Response.Write("10") Else Response.Write(Session(Sys.CookieName & "PerPage"))%>" class="text" size="1" id="PageSize"> <a href="javascript:;" onclick="ceeevio.Link.SetSession('PerPage', $('#PageSize').val())">确定</a></div>
                <div>2. <a href="javascript:;" onclick="ceeevio.Link.SetSession('IsShow', 1)">已通过链接</a> &nbsp;&nbsp;<a href="javascript:;" onclick="ceeevio.Link.SetSession('IsShow', 0)">未通过链接</a></div>
                 <div>3. <a href="javascript:;" onclick="ceeevio.Link.SetSession('IsMain', 1)">已置顶链接</a> &nbsp;&nbsp;<a href="javascript:;" onclick="ceeevio.Link.SetSession('IsMain', 0)">未置顶链接</a></div>
                 <div>4. 链接顺序 &nbsp;<a href="javascript:;" onclick="ceeevio.Link.SetSession('Order', 0)">顺序</a> &nbsp;&nbsp;<a href="javascript:;" onclick="ceeevio.Link.SetSession('Order', 1)">倒序</a></div>
                 <div>5. <a href="javascript:;" onclick="ceeevio.Link.ShowAll()">显示全部</a></div>
			</div>
        </div>
        
    </div>
    
	<div style="margin:10px 20px 10px 220px" id="rightbox">
    
    <table cellpadding="3" cellspacing="0" class="CeeTable" width="100%">
    	<thead>
        	<tr>
            	<th width="50"><input type="checkbox"></th>
                <th width="150">站点名</th>
                <th>链接</th>
                <th width="300">logo</th>
                <th width="100"><a href="javascript:ceeevio.Link.AddLink.Box();" style="font-weight:normal" class="newadd">+ 添加新链接</a></th>
            </tr>
        </thead>
        <tfoot>
        	<tr>
            	<th width="50"><input type="checkbox"></th>
                <th width="150">站点名</th>
                <th>链接</th>
                <th width="300">logo</th>
                <th width="100"><a href="javascript:javascript:ceeevio.Link.AddLink.Box();" style="font-weight:normal" class="newadd">+ 添加新链接</a></th>
            </tr>
        </tfoot>
        <tbody>
<%	
	Sql = "Select link_ID, link_Name, link_URL, link_Image, link_Order, link_IsShow, link_IsMain, Link_ClassID, Link_Type From blog_Links Where Link_Type=false " & CateID & " " & IsShow & " " & IsMain & " Order By link_IsShow desc" & Order
	'				0		1			2			3			4			5			6			7
	'8
	'Response.Write(Sql)
	Set Rs = Server.CreateObject("Adodb.RecordSet")
	Rs.open Sql, Conn, 1, 1
	If Rs.Bof Or Rs.Eof Then
		ReDim LinkArray(0, 0)
	Else
		LinkArray = Rs.GetRows
	End If
	Rs.Close
	Set Rs = Nothing
	If UBound(LinkArray, 1) = 0 Then
		Response.Write("<tr><td colspan=""4"" align=""center"">没有找到信息</td></tr>")
	Else
		linkLen = UBound(LinkArray, 2)
		If (linkLen + 1) Mod PerPage = 0 Then
			PageCount = Int((linkLen + 1) / PerPage)
		Else
			PageCount = Int((linkLen + 1) / PerPage) + 1
		End If
		If Int(page) > Int(PageCount) Then page = PageCount
		linkLeft = (page - 1) * PerPage
		linkRight = page * PerPage - 1
		If linkLeft < 0 Then linkLeft = 0
		If linkRight > linkLen Then linkRight = linkLen
		For i = linkLeft To linkRight
		
%>
		<tr class="active" id="link_<%=LinkArray(0, i)%>">
        	<td width="50"><input type="checkbox" value="" name="SelectID" style="margin-left:4px"></td>
            <td width="150"><div id="name_<%=LinkArray(0, i)%>"><%=Trim(LinkArray(1, i))%></div></td>
            <td><div id="url_<%=LinkArray(0, i)%>"><%=Trim(LinkArray(2, i))%></div></td>
            <td width="300"><div id="image_<%=LinkArray(0, i)%>"><%=Trim(LinkArray(3, i))%></div></td>
            <td width="100"><div id="order_<%=LinkArray(0, i)%>"><%=Trim(LinkArray(4, i))%></div></td>
        </tr>
        <tr class="active second">
        	<td width="50">&nbsp;</td>
            <td width="150"><div id="tool_<%=LinkArray(0, i)%>"><a href="javascript:ceeevio.Link.DelLink(<%=LinkArray(0, i)%>);" class="linkdel">删除</a> | <a href="javascript:ceeevio.Link.editBox(<%=LinkArray(0, i)%>);" class="linkeditbutton">修改</a></div></td>
            <td>
            <%If LinkArray(5, i) Then%>
            	<a href="javascript:ceeevio.Link.Show(<%=LinkArray(0, i)%>, 'unshow');" class="linkdel" id="doshow_<%=LinkArray(0, i)%>">取消通过</a>
            <%Else%>
            	<a href="javascript:ceeevio.Link.Show(<%=LinkArray(0, i)%>, 'show');" class="linkdel" id="doshow_<%=LinkArray(0, i)%>">通过</a>
            <%End If%>
             |  
             <%If LinkArray(6, i) Then%>
             	<a href="javascript:ceeevio.Link.main(<%=LinkArray(0, i)%>, 'unmain')" class="linkdel" id="domain_<%=LinkArray(0, i)%>">取消置顶</a>
             <%Else%>
             	<a href="javascript:ceeevio.Link.main(<%=LinkArray(0, i)%>, 'tomain')" class="linkdel" id="domian_<%=LinkArray(0, i)%>">置顶</a>
             <%End If%>
              |  <a href="<%=Trim(LinkArray(2, i))%>" target="_blank">访问</a> | 
              转移到 <select id="move_<%=LinkArray(0, i)%>" onchange="ceeevio.Link.move(<%=LinkArray(0, i)%>, $(this).val())"><%=CateSelect(LinkArray(7, i))%></select></td>
            <td align="center">&nbsp;</td>
            <td align="center"><div id="edit_<%=LinkArray(0, i)%>" style="text-align:center; display:none"><input type="button" value="修改" class="button" onClick="ceeevio.Link.postLinkedit(<%=LinkArray(0, i)%>)"></div></td>
        </tr>
<%
		Next
		linkPageStr = MultiPage(linkLen + 1, PerPage, page, "?Fmenu=Link&page={$page}", "float:right", "", false)
	End If
%>
    </tbody>
    </table>
    <div class="page" style="margin:10px 10px 10px 10px;"><span style="float:left">共有链接 <%=Sys.doGet("Select Count(*) From blog_Links Where Link_Type=false")(0)%> 个</span><%=linkPageStr%></div>
    </div>
   </td></tr>
</table>
<%
Control.getMsg
End Sub

Public Function CateSelect(Sess)
	Dim Rs
	Set Rs = Conn.Execute("Select * From blog_Links Where Link_Type=true")
	If Rs.Bof Or Rs.Eof Then
		CateSelect = ""
	Else
		CateSelect = ""
		Do While Not Rs.Eof
			If Not Asp.IsInteger(Sess) Then
				CateSelect = CateSelect & "<option value=""" & Rs("link_ID").value & """>" & Rs("link_Name").value & "</option>"
			Else
				If Int(Rs("link_ID").value) = Int(Sess) Then
					CateSelect = CateSelect & "<option value=""" & Rs("link_ID").value & """ selected=""selected"">" & Rs("link_Name").value & "</option>"
				Else
					CateSelect = CateSelect & "<option value=""" & Rs("link_ID").value & """>" & Rs("link_Name").value & "</option>"
				End If
			End If
		Rs.MoveNext
		Loop
	End If
	Set Rs = Nothing
End Function
%>