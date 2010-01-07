<%
Public Sub c_categories
	Dim Rs, CateRow, i, joinStr
%>
<script language="javascript" type="text/javascript">
$(function(){
	$("span[rel='editOrder']").hover(
		function(){
			try{mses.obj.text(mses.value);}catch(e){}
			var c = $(this).text();
			mses = {obj : $(this), value : c}
			$(this).html("<input type=\"text\" value=\"" + c + "\" class=\"text\" size=\"4\" />");
		},
		function(){
			var c = mses.obj.find("input[type='text']").val();
			if (c != mses.value){
				jConfirm('确定保存该排序?', '确认对话框', function(r) {
					if (r){
						mses.obj.text(c);
						mses = null;
					}else{
						mses.obj.text(mses.value);
						mses = null;
					}
				});	
			}else{
				mses.obj.text(mses.value);
				mses = null;
			}
		}
	);
	
	$(".spanMan").toggle(
		function(){
			$(this).parents("tr:first").next().find("td > .showHide").hide("fast");
			$(this).attr({
						 src : "../../images/Control/shouqi.gif",
						 title : "展开"
			});
		},
		function(){
			$(this).parents("tr:first").next().find("td > .showHide").show("slow");
			$(this).attr({
						 src : "../../images/Control/xiala.gif",
						 title : "收起"
			});
		}
	);
});
var mses = null;
</script>
<style type="text/css">
.content h5{
	color:#1d6bb2;
	font-size:14px;
	margin:0;
}
h5.category{
	background: url(../../images/Control/Icon/toppiao.gif) no-repeat;
	line-height:28px;
	height:28px;
	padding-left:35px;
}
.ConZone{
	margin-top:10px;
}
.cateFiled{
	padding:10px 20px;
}
.cateFiled table tr td{
	line-height:30px;
}
.cateSlide td{
	border-top:1px solid #EAEDF0;
}
.edox{ border:1px solid #B6C3D4; padding:0px 2px 2px 2px; border-top:1px solid #fff; z-index:99}
.spanMan{
	cursor:pointer;
}
</style>
<h5 class="category">分类管理</h5>
<div class="ConZone">
	<div class="cateFiled">
    <div style="margin-bottom:10px;">
    	<input type="button" value="帮助" class="button" onclick="$('.yellowbox').slideToggle('fast')">
        <input type="button" value="添加" class="button" onclick="conMain.category.add(this)" />
    </div>
    <%=yellowBox("分类模块注意事项:<br />1. 添加分类请点击＂添加＂按钮, 新增成功后请选择 <strong>是否刷新页面</strong> 来显示该数据.<br />2. 分类的信息现实模式带有上下展开和收起功能, 分类数据量多的用户可以使用收起.<br />3. 简单更新分类的排序,只需要鼠标移动到右侧排序序号上便能自动更新,填写完成后根据提示进行确认.<br />4. 删除分类时须谨慎, 删除分类的同时也将删除该分类下的所有日志,切勿盲目操作!")%>
        	<table cellpadding="3" cellspacing="0" width="100%" style="margin-top:10px;">
<%
	Set Rs = Conn.Execute("Select cate_ID, cate_Order, cate_icon, cate_Name, cate_Intro, cate_Folder, cate_URL, cate_local, cate_Secret, cate_count, cate_OutLink, cate_Lock From blog_Category Order By cate_local asc, cate_Order asc")
	'								0			1			2			3		4				5			6
'	7			8			9			10				11
		CateRow = Rs.GetRows
		Rs.Close
	Set Rs = Nothing
	joinStr = Strinel
%>
<%
	If UBound(CateRow, 1) = 0 Then
	Else
		For i = 0 To UBound(CateRow, 2)
%>
		
            	<tr class="cateSlide">
                	<td width="30" align="center"><img src="<%=CateRow(2, i)%>" /></td>
                    <td>
                    	<strong style="color:#0F5FBB; margin-right:10px;"><%=CateRow(3, i)%></strong> <%If Len(CateRow(4, i)) > 0 Then%>＂<%=CateRow(4, i)%>＂<%End If%>
                    </td>
                    <td>
                    	<span style="float:right; height:30px; line-height:30px;">
                        	<img src="../../images/Control/yincang.gif" width="8" height="8">
                            &nbsp;&nbsp;
                            <img src="../../images/Control/xiala.gif" width="9" height="8" class="spanMan" title="收起">
                        </span>
                    </td>
                </tr>
            	<tr>
                	<td></td>
                    <td>
                    <div class="showHide" style="line-height:18px; padding:0px 10px;border-left:2px solid #d6d6d6;">
                    	<%If CateRow(10, i) Then%><div>外部链接地址 : <%=CateRow(6, i)%></div><%End If%>
                        <%If Not CateRow(10, i) Then%><div>存放文章目录 : <%=CateRow(5, i)%></div><%End If%>
                        <div>分类所在位置 : <%
												If Int(CateRow(7, i)) = 1 Then
													Response.Write("顶部")
												ElseIf Int(CateRow(7, i)) = 1 Then
													Response.Write("侧边")
												Else
													Response.Write("同时")
												End If
											%>
                        </div>
                        <div>分类状态 : <%
											If Int(CateRow(8, i)) Then
												Response.Write("保密")
											Else
												Response.Write("公开")
											End If
										%>
                        </div>
                    </div>
                    </td>
                    <td>
                    	<div class="hideData">
                        	<div><%=CateRow(0, i)%></div>
                            <div><%=CateRow(1, i)%></div>
                            <div><%=CateRow(2, i)%></div>
                            <div><%=CateRow(3, i)%></div>
                            <div><%=CateRow(4, i)%></div>
                            <div><%=CateRow(5, i)%></div>
                            <div><%=CateRow(6, i)%></div>
                            <div><%=CateRow(7, i)%></div>
                            <div><%=CateRow(8, i)%></div>
                            <div><%=CateRow(9, i)%></div>
                            <div><%=CateRow(10, i)%></div>
                            <div><%=CateRow(11, i)%></div>
                        </div>
                    </td>
                </tr>
                <tr>
                	<td>&nbsp;</td>
                	<td><span><a href="javascript:;" onclick="conMain.category.expert(this)">编辑</a> | <a href="javascript:;" onclick="conMain.category.del(<%=CateRow(0, i)%>)">删除</a> <%If CateRow(10, i) Then%>| <strong>外部链接</strong><%End If%></span></td>
                    <td><span style="float:right; font-size:11px; font-family:Georgia, 'Times New Roman', Times, serif; font-style:oblique"><span rel="editOrder" style="line-height:30px;"><%=CateRow(1, i)%></span></span></td>
                </tr>
<%
		Next
	End If
%>
			</table>
	</div>
</div>
<%
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