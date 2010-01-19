<%
Public Sub c_link
%>
<script language="javascript" type="text/javascript">
$(function(){
	$("ul.slide").tabs("ul.toolbar", {
		current : "actived"
	});	
	conMain.links.NewClassAdd();
	//$(".linkItem3:odd .linkItem_Content").css("background", "#fcfcfc")
});
</script>
<style type="text/css">
.content h5{
	color:#1d6bb2;
	font-size:14px;
	margin:0;
}
h5.link{
	background: url(../../images/Control/Icon/link.gif) no-repeat;
	line-height:28px;
	height:28px;
	padding-left:35px;
}

.toolbar li{
	padding:0px 20px 10px 20px;
	width:650px;
}

.select{
	width:750px;
}

.select .Zcontent{
	margin-top:12px;
}

.tagname{color:#424242}

.setting table tr td{
	line-height:30px;
}

.label{
	margin-bottom:10px;
}

#linkContent{ margin-top:20px; line-height:30px;}
#toolsInfo{ line-height:30px; border:1px dotted #ccc; margin-top:10px; padding:10px;}

.cload{ float:left}
.mload{ float:right}
.corder{ width:50px;}
.cname{ width:300px;}
.maction{ width:100px;}
.maction a{ float:left; margin-right:3px}


.linkItem{ margin-bottom:10px; position:relative}
.linkItem .linkItem_Top{ background:url(../../images/Control/linedit_t.png) no-repeat; width:631px; height:4px;}
.linkItem .linkItem_Mid{ background:url(../../images/Control/linkedit_m.png) repeat-y; width:631px;}
.linkItem .linkItem_Bom{ background:url(../../images/Control/linedit_b.png) no-repeat; width:631px; height:4px;}
.linkItem .linkItem_Content{ padding:5px 15px;}
.linkItem .itemLeft{ float:left; width:85px;}
.linkItem .itemRight{ float:left; width:470px; margin-left:15px; border-left:1px solid #DDDBDB; padding-left:15px;}
.linkItem .itemLeft .itemLeftTop{ height:50px; line-height:50px; text-align:center; font-size:24px; font-family: Georgia, "Times New Roman", Times, serif}
.linkItem .itemLeft .itemLeftBom{ text-align:center}
.linkItem .itemRight .linkItemTitle{ line-height:30px; height:30px;}
.linkItem .itemRight .linkItemInfo{ margin:0 0 4px 0; color:#666;line-height:30px; text-indent:2em}
.linkItem .itemRight .linkItemAction{ line-height:30px; height:30px;}
.linkItem .itemRight .linkItemTitle .linkItemTitleleft{ float:left; color:#0F5FBB; font-weight:700}
.linkItem .itemRight .linkItemTitle .linkItemTitleright{ float:right; font-style:oblique; font-size:10px;}
.linkItem .top{ background:url(../../images/Control/link_top.png) no-repeat; width:11px; height:11px; position:absolute; top:4px; left:4px;}

.linkItem2{ margin-bottom:10px; position:relative}
.linkItem2 .linkItem_Top{ background:url(../../images/Control/linedit_t_hover.png) no-repeat; width:631px; height:4px;}
.linkItem2 .linkItem_Mid{ background:url(../../images/Control/linkedit_m_hover.png) repeat-y; width:631px;}
.linkItem2 .linkItem_Bom{ background:url(../../images/Control/linedit_b_hover.png) no-repeat; width:631px; height:4px;}
.linkItem2 .linkItem_Content{ padding:5px 15px;}
.linkItem2 .itemLeft{ float:left; width:85px;}
.linkItem2 .itemRight{ float:left; width:470px; margin-left:15px; border-left:1px solid #49A1F6; padding-left:15px;}
.linkItem2 .itemLeft .itemLeftTop{ height:50px; line-height:50px; text-align:center; font-size:24px; font-family: Georgia, "Times New Roman", Times, serif}
.linkItem2 .itemLeft .itemLeftBom{ text-align:center}
.linkItem2 .itemRight .linkItemTitle{ line-height:30px; height:30px;}
.linkItem2 .itemRight .linkItemInfo{ margin:0 0 4px 0; color:#666;line-height:30px; text-indent:2em}
.linkItem2 .itemRight .linkItemAction{ line-height:30px; height:30px;}
.linkItem2 .itemRight .linkItemTitle .linkItemTitleleft{ float:left; color:#0F5FBB; font-weight:700}
.linkItem2 .itemRight .linkItemTitle .linkItemTitleright{ float:right; font-style:oblique; font-size:10px;}
.linkItem2 .top{ background:url(../../images/Control/link_top.png) no-repeat; width:11px; height:11px; position:absolute; top:0px; left:0px;}

.linkItem3{}
.linkItem3 .linkItem_Top{ width:631px; height:4px;}
.linkItem3 .linkItem_Mid{ width:631px;}
.linkItem3 .linkItem_Bom{ width:631px; height:4px;}
.linkItem3 .linkItem_Content{ padding:5px 15px; border-bottom:1px dashed #efefef}
.linkItem3 .itemLeft{ float:left; width:85px;}
.linkItem3 .itemRight{ float:left; width:470px; margin-left:15px; border-left:1px solid #49A1F6; padding-left:15px;}
.linkItem3 .itemLeft .itemLeftTop{ height:50px; line-height:50px; text-align:center; font-size:24px; font-family: Georgia, "Times New Roman", Times, serif}
.linkItem3 .itemLeft .itemLeftBom{ text-align:center}
.linkItem3 .itemRight .linkItemTitle{ line-height:30px; height:30px;}
.linkItem3 .itemRight .linkItemInfo{ margin:0 0 4px 0; color:#666;line-height:30px; text-indent:2em}
.linkItem3 .itemRight .linkItemAction{ line-height:30px; height:30px;}
.linkItem3 .itemRight .linkItemTitle .linkItemTitleleft{ float:left; color:#0F5FBB; font-weight:700}
.linkItem3 .itemRight .linkItemTitle .linkItemTitleright{ float:right; font-style:oblique; font-size:10px;}

.linkItem3Title .cload{ color:#000}
.linkItem3Title{ margin:0px}
.linkItem3Title .linkItem_Content{border-bottom:1px solid #bbb}

.addNewlinkCate{ border:1px dashed #ccc; background:#efefef; padding:10px 20px; margin:10px; line-height:30px; display:none;}

</style>
<h5 class="link">友情链接管理</h5>
	<div class="select">
        <div class="tits">
            <ul class="slide">
                <li>基本链接</li>
                <li>链接分类</li>
            </ul>
        </div>
        
        <div class="Zcontent">
            <ul class="toolbar">
                <li><%Links%></li>
                <li><%LinkCates%></li>
            </ul>
        </div>
   </div>
<%
End Sub

' ***********************************
'	链接管理
' ***********************************
Public Sub Links
%>
	<div class="tools"><input type="button" value="添加链接" class="button" onclick="conMain.links.addNewLink(this, 11)" /><span style="margin-left:10px;"><input type="button" value="批量管理" class="button" onclick="$('#toolsInfo').slideToggle('fast')"></span></div>
    <div id="toolsInfo" style="display:none">aaa</div>
    <div id="linkContent">

<%
Dim Rs, id
Set Rs = Conn.Execute("Select * From blog_Links Where Link_Type=False Order By link_Order ASC")
If Rs.Bof Or Rs.Eof Then

Else
	Do While Not Rs.Eof
		id = Rs("link_ID").value
%>
    
    	<div id="linkItem_<%=id%>" class="linkItem">
        	<div class="linkItem_Top"></div>
            <div class="linkItem_Mid">
            	<div class="linkItem_Content">
                	<div class="itemLeft">
                    	<div class="itemLeftTop" name="link_order"><%=Rs("link_Order").value%></div>
                        <div class="itemLeftBom"><%If Rs("link_IsShow").value Then%>已通过<%Else%>未通过<%End If%></div>
                    </div>
                    <div class="itemRight">
                    	<div class="linkItemTitle"><div class="linkItemTitleleft" name="link_name"><%=Rs("link_Name").value%></div> <div class="linkItemTitleright" name="link_url"><%=Rs("link_URL").value%></div><div class="clear"></div></div>
                        <div class="linkItemInfo" name="link_info">
							<%
								If LCase(TypeName(Rs("link_info").value)) = "null" Then
									Response.Write("[..该链接暂未链接说明..]") 
								Else
									Response.Write(Rs("link_info").value)
								End If
							%>
                        </div>
                        <div class="linkItemAction">快捷操作：<span class="active">
                        	<a href="javascript:;" onclick="conMain.links.CanEdit($('#linkItem_<%=id%>'), <%=id%>)">编辑</a> | 
                            <a href="javascript:;" onclick="conMain.links.dellins(<%=id%>)">删除</a> | 
                            <%If Rs("link_IsMain").value Then%><a href="javascript:;" onclick="conMain.links.Action('unmain', <%=id%>)">已置顶</a> | <%Else%><a href="javascript:;" onclick="conMain.links.Action('tomain', <%=id%>)">未置顶</a> | <%End If%>
                            <%If Rs("link_IsShow").value Then%><a href="javascript:;" onclick="conMain.links.Action('unshow', <%=id%>)">已通过</a><%Else%><a href="javascript:;" onclick="conMain.links.Action('show', <%=id%>)">未通过</a><%End If%>
                            
                        </span>
                        <span class="unactive" style="display:none">
                        	<a href="javascript:;" class="saveLink">保存</a> 
                            <a href="javascript:;" onclick="conMain.links.UnCanEdit($('#linkItem_<%=id%>'))">取消</a>
                        </span>
                        </div>
                    </div>
                    <div class="clear"></div>
                </div>
            </div>
            <%If Rs("link_IsMain").value Then%><div class="top"></div><%End If%>
            <div class="linkItem_Bom"></div>
        </div>
<%
	Rs.MoveNext
	Loop
End If
%>       
    </div>
<%
End Sub

' ***********************************
'	分类管理
' ***********************************
Public Sub LinkCates
%>
	<div>快捷操作： <a href="javascript:;" onclick="$('.addNewlinkCate').toggle('fast')">添加分类</a></div>
    <div class="addNewlinkCate">
    <form action="../pjblog.logic/control/log_link.asp?action=addNewClass" method="post" id="NewClassAdded">
    	<div>分类名称 <input type="text" name="link_Name" class="text"></div>
        <div>分类排序 <input type="text" name="link_Order" class="text" value="0"></div>
        <div><input type="submit" value="提交" class="button"></div>
    </form>
    </div>
    <div style="margin-top:20px;">
    
    <div class="linkItem3 linkItem3Title">
        <div class="linkItem_Top"></div>
        <div class="linkItem_Mid">
            <div class="linkItem_Content">
                <div class="cload corder" name="link_order">排序</div>
                <div class="cload cname" name="link_name">分类名称</div>
                <div class="mload maction">
                <div class="clear"></div>
                </div><div class="clear"></div>
            </div>
        </div>
        <div class="linkItem_Bom"></div>
    </div>
<%
	Dim Rs 'link_Name, link_Order
	Set Rs = Conn.Execute("Select * From blog_Links Where Link_Type=True Order By link_Order Asc")
	If Not (Rs.Bof And Rs.Eof) Then
		Do While Not Rs.Eof
%>
    <div id="linkClassItem_<%=Rs("link_ID")%>" class="linkItem3">
        <div class="linkItem_Top"></div>
        <div class="linkItem_Mid">
            <div class="linkItem_Content">
                <div class="cload corder" name="link_order"><%=Rs("link_Order")%></div>
                <div class="cload cname" name="link_name"><%=Rs("link_Name")%></div>
                <div class="mload maction">
                <a href="javascript:;" onclick="conMain.links.NewClassEditBox(<%=Rs("link_ID")%>)" class="linkclasseditbutton">编辑</a> <a href="javascript:;" class="linkclassdelebutton" onclick="conMain.links.ClassDelete(<%=Rs("link_ID")%>)">删除</a> <a href="javascript:;" style="display:none;" class="linkclassupdatebutton">更新</a>
                <div class="clear"></div>
                </div><div class="clear"></div>
            </div>
        </div>
        <div class="linkItem_Bom"></div>
    </div>
<%
		Rs.MoveNext
		Loop
	End If
	Set Rs = Nothing
%>
	</div>
<%
End Sub
%>