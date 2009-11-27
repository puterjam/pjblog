<%
Public Sub c_categories
	Dim Rs, CateRow, i, joinStr
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
	<tr>
    	<th class="CTitle"><%=control.categoryTitle%></th>
  	</tr>
  	<tr>
    	<td class="CPanel">
            <div class="SubMenu">
                <a href="?Fmenu=Categories">设置日志分类</a> | 
                <a href="?Fmenu=Categories&Smenu=move">批量移动日志</a> | 
                <a href="?Fmenu=Categories&Smenu=del">批量删除日志</a> | 
                <a href="?Fmenu=Categories&Smenu=tag">Tag管理</a>
            </div>
<%
If Request.QueryString("Smenu") = "move" Then

Else
%>
			<form action="#" method="post" id="CateList">
			<table cellpadding="3" cellspacing="0" style="margin-top:10px; width:100%" id="CateTable">
            	<tr>
                	<td class="SecTd" width="29">&nbsp;</td>
                	<td class="SecTd" width="29" nowrap>排序</td>
                    <td class="SecTd">&nbsp;</td>
                    <td class="SecTd">标题</td>
                    <td class="SecTd">提示说明</td>
                    <td class="SecTd">文章目录名</td>
                    <td class="SecTd">外部链接</td>
                    <td class="SecTd">&nbsp;</td>
                    <td class="SecTd">&nbsp;</td>
                    <td class="SecTd">&nbsp;</td>
                </tr>
<%
	Set Rs = Conn.Execute("Select cate_ID, cate_Order, cate_icon, cate_Name, cate_Intro, cate_Folder, cate_URL, cate_local, cate_Secret, cate_count, cate_OutLink, cate_Lock From blog_Category Order By cate_local asc, cate_Order asc")
		CateRow = Rs.GetRows
		Rs.Close
	Set Rs = Nothing
	joinStr = Strinel
%>
	<script language="javascript">
		var carr = new Array(), canSubmit = true;
		carr.push("<%=Replace(CommonTipPic(joinStr, "")(0), Chr(34), "'")%>");
		carr.push("<%=Replace(CommonTipPic(joinStr, "")(1), Chr(34), "'")%>");
		$(document).ready(function(){
				$("#CateList").ajaxForm({
					dataType : "json",
					beforeSubmit : function(){
						if (!canSubmit){alert('请将新分类内容填写完整后提交!') ; return false;}
						else if (doWhich = null){return false;}
						else{
							effect.WarnTip.open({
								title 	: 	'<strong>waiting...</strong>',
								html	:	'请等待,正在提交数据...'
							});
							return true;
						}
					},
					success : function(data){
						effect.WarnTip.open({
							title 	: 	'恭喜 操作成功',
							html	:	data.Info
						});
						setTimeout("window.location.reload();", 1000);
					}
				});
			}						  
		);
		function whatdo(str){
			$("#CateList").attr("action", "../pjblog.logic/control/log_category.asp?action=" + str);
			$("#CateList").submit();
		}
	</script>
<%
	If UBound(CateRow, 1) = 0 Then
%>
				<tr><td colspan="10">没有分类信息,请添加!</td></tr>
<%
	Else
		For i = 0 To UBound(CateRow, 2)
%>
                <tr id="Catetr_<%=Int(CateRow(0, i))%>" class="ceo SecTd">
                	<td width="25"><%If Not CateRow(11, i) Then%><input type="checkbox" value="<%=Int(CateRow(0, i))%>" name="SelectID" /><%else%>&nbsp;<%End If%><input type="hidden" value="<%=Int(CateRow(0, i))%>" name="Cate_ID" /></td>
                	<td class="SecTd" width="29"><input type="text" value="<%=Int(CateRow(1, i))%>" name="cate_Order" class="text" size="2" /></td>
                    <td>
                    <img id="CateImg_<%=Int(CateRow(0, i))%>" src="<%=Trim(CateRow(2, i))%>" width="16" height="16" />
                    <Select name="Cate_icons" onChange="$('#CateImg_<%=Int(CateRow(0, i))%>').attr('src', this.options[this.options.selectedIndex].value)" style="width:120px;">
                    	<%=CommonTipPic(joinStr, Trim(CateRow(2, i)))(0)%>
                    </Select>
                    </td>
                    <td><input name="Cate_Name" type="text" class="text" value="<%=Trim(CateRow(3, i))%>" size="14"/></td>
                    <td><input name="Cate_Intro" type="text" class="text" value="<%=Trim(CateRow(4, i))%>" size="20"/></td>
                    <td><input name="Cate_Part" type="text" class="text" value="<%=Trim(CateRow(5, i))%>" size="16" onblur="ReplaceInput(this, window.event)" onkeyup="ReplaceInput(this, window.event)" /></td>
                    <td><input name="cate_URL" type="text" value="<%=Trim(CateRow(6, i))%>" size="30" class="text" <%If Int(CateRow(9, i)) > 0 Then Response.Write "readonly=""readonly"" style=""background:#e5e5e5""" %> /></td>
                    <td>
                    <select name="Cate_local" onChange="document.getElementById('Catetr_<%=Int(CateRow(0, i))%>').style.backgroundColor=(this.value==1)?'#a9c9e9':(this.value==2)?'#bcf39e':''">
			            <option value="0">同时</option>
			            <option value="1" <%If Int(CateRow(7, i)) = 1 Then Response.Write "selected=""selected"""%>>顶部</option>
			            <option value="2" <%If Int(CateRow(7, i)) = 2 Then Response.Write "selected=""selected"""%>>侧边</option>
		           	</select>
                    </td>
                    <td>
                    <select name="cate_Secret" <%If CateRow(10, i) Then Response.Write "disabled=""disabled"""%>>
		            	<option value="0" style="background:#0f0">公开</option>
		            	<option value="1" style="background:#f99" <%if Int(CateRow(8, i)) then Response.Write "selected=""selected"""%>>保密</option>
		           	</select>
                    <%If CateRow(10, i) Then%><input type="hidden" value="<%=Int(CateRow(8, i))%>" name="cate_Secret"><%End If%>
                    </td>
                    <td><input type="text" class="text" name="cate_count" value="<%=Int(CateRow(9, i))%>" size="2" readonly="readonly" style="background:#ffe"/> 篇</td>
                </tr>
<%
		Next
%>
				<tr style=" line-height:25px;" id="AddRowMark">
                	<td colspan="2">&nbsp;</td>
                    <td colspan="8"><div class="AddNew"><a href="javascript:void(0);" class="AddA" onClick="ceeevio.category.add();" id="Addbutton">添加新分类</a></div></td>
                </tr>
                <tr style=" line-height:30px;">
                	<td>&nbsp;</td>
                    <td colspan="9">
                    <input type="button" value="保存分类" class="button" onclick="whatdo('edit')">
                    <input type="button" value="批量删除" class="button" onclick="whatdo('del')">
                    </td>
                </tr>
<%
	End If
%>
			</table>
            </form>
<%
End If
%>
        </td>
	</tr>
</table>
<%
Control.getMsg
End Sub

Public Function Strinel
	Dim fso
	Set fso = New cls_fso
		Strinel = fso.FileItem("../images/icons")
	Set fso = Nothing
End Function

Public Function CommonTipPic(ByVal Arrays, ByVal Str)
	Dim JoinArray, i, o, t, c
	t = "" : o = ""
	JoinArray = Split(Arrays, "|")
	For i = 1 To UBound(JoinArray)
		If "../images/icons/" & JoinArray(i) = Str Then
			t = t & "<option value=""../images/icons/" & JoinArray(i) & """ selected=""selected"">" & JoinArray(i) & "</option>"
			o = "../images/icons/" & JoinArray(i)
		Else
			c = "<option value=""../images/icons/" & JoinArray(i) & """>" & JoinArray(i) & "</option>"
			If i = UBound(JoinArray) Then
				If Len(o) = 0 Then c = "<option value=""../images/icons/" & JoinArray(i) & """ selected>" & JoinArray(i) & "</option>"
			End If
			t =t & c
		End If
	Next
	If Len(o) = 0 Then
		CommonTipPic = Array(t, "../images/icons/" & JoinArray(UBound(JoinArray)))
	Else
		CommonTipPic = Array(t, o)
	End If
End Function
%>