<%
Public Sub Skin_Set
%>
<script language="javascript" type="text/javascript">
$(function(){
	$(".yellowbox").slideDown('fast');
});
</script>
<style type="text/css">
.content h5{
	color:#1d6bb2;
	font-size:14px;
	margin:0;
}
h5.skinset{
	background: url(../../images/Control/Icon/waiguan.gif) no-repeat;
	line-height:28px;
	height:28px;
	padding-left:35px;
}
.readinfo, .readStyle{ display:none;}
.ConZone{margin-top:10px;}
.settitle{line-height:30px;border-top: 1px solid #64B0F5;}
.settitle a{margin-right:10px;}
.setBody{}
.setBody ul{ list-style:none; margin-left:0; margin-right:0; padding-left:0px; padding-right:0px;}
.setBody ul li.item{ float:left; width:33%; margin-bottom:15px}
.setBody ul li.item .skinbg{ width:204px; margin:0 auto;}
.setBody ul li.item .skinbg .top{ height:6px; width:204px; background:url(../../images/Control/skinset_t.png) no-repeat;}
.setBody ul li.item .skinbg .mid{ width:204px; background:url(../../images/Control/skinset_m.png) repeat-y}
.setBody ul li.item .skinbg .bom{ height:9px; width:204px; background:url(../../images/Control/skinset_b.png) no-repeat;}
.setBody ul li.item .skinbg .mid .whole{ padding:0px 10px}
.setBody ul li.item .skinbg .mid .whole .tools{ height:20px;}


.setBody ul li.item .skinbg .mid .whole .tools .yulan{ width:13px; background:url(../../images/Control/yulan.gif) no-repeat; float:right; height:16px;margin-right:5px;}
.setBody ul li.item .skinbg .mid .whole .tools .yulan:hover{ width:13px; background: url(../../images/Control/yulan2.gif) no-repeat; float:right; height:16px;margin-right:5px; cursor:pointer;}

.setBody ul li.item .skinbg .mid .whole .tools .fabu{ float:right; height:16px; width:12px; background:url(../../images/Control/fabu.gif) no-repeat; margin-right:5px;}
.setBody ul li.item .skinbg .mid .whole .tools .fabu:hover{float:right; height:16px; width:12px; background: url(../../images/Control/fabu2.gif) no-repeat; margin-right:5px; cursor:pointer;}

.setBody ul li.item .skinbg .mid .whole .tools .daochu{ float:right; height:16px; width:17px; background:url(../../images/Control/daochu.gif) no-repeat;}
.setBody ul li.item .skinbg .mid .whole .tools .daochu:hover{ float:right; height:16px; width:17px; background: url(../../images/Control/daochu2.gif) no-repeat; cursor:pointer;}


.setBody ul li.item .skinbg .mid .whole .themepic{ border:1px solid #f1f0f0; width:180px; position:relative}
.setBody ul li.item .skinbg .mid .whole .themepic:hover{ border:1px solid #C4E0FB; width:180px; position:relative; cursor:pointer}
.setBody ul li.item .skinbg .mid .whole .themepic .checked{ width:21px; height:20px; background:url(../../images/Control/xuanzhong.gif) no-repeat; position:absolute; top:0px; left:0px;}
.setBody ul li.item .skinbg .mid .whole .themetitle{ height:30px;}
.setBody ul li.item .skinbg .mid .whole .themetitle .themename{ width:180px; overflow:hidden; text-align:center; line-height:30px; margin:0 auto; color:#1d6bb2; font-weight:700;}
.setBody ul li.item .skinbg .mid .whole .themepic .img{ padding:10px;}
.setBody ul li.item .skinbg .mid .whole .themepic .img img{ width:160px; height:137px;}
.style{
	margin-left:58px;
	line-height:22px;
	list-style:none;
	width:300px;
	overflow:hidden;
}
.style li{
	width:44px;
	height:60px;
	overflow:hidden;
	float:left;
	margin-right:10px;
}
.cimg{
	border:1px solid #fff;
	padding:1px;
}
.style li .cimg:hover{
	cursor:pointer;
}
.selectimg{
	border:1px solid #327EC9;
	padding:1px;
}
</style>
<h5 class="skinset">外观设置 - 主题设置</h5>
<div class="ConZone">
	<div class="settitle">
    	快速通道 : &nbsp;<a href="?m=skin&s=net">在线安装主题</a><a href="?m=skin&s=upd">主题导入</a>
    </div>
    <%=yellowBox("分类模块注意事项:<br />1. 添加分类请点击＂添加＂按钮, 新增成功后请选择 <strong>是否刷新页面</strong> 来显示该数据.<br />2. 分类的信息现实模式带有上下展开和收起功能, 分类数据量多的用户可以使用收起.<br />3. 简单更新分类的排序,只需要鼠标移动到右侧排序序号上便能自动更新,填写完成后根据提示进行确认.<br />4. 删除分类时须谨慎, 删除分类的同时也将删除该分类下的所有日志,切勿盲目操作!")%>
<%
	Dim cxml, fso, FolderItems, TempFolderArray, i, txml, txmls, j, zoom, temp
	Dim SkinName, SkinDesigner, pubDate, DesignerURL, DesignerMail, version, smallpic, bigpic, Intro
	Set cxml = New xml
	Set zoom = New xml
	Set fso = New cls_fso
		FolderItems = fso.FolderItem("../pjblog.template")
		TempFolderArray = Split(FolderItems, "|")
		If Int(TempFolderArray(0)) = 0 Then ' 是否有文件夹
			Response.Write("<p>找不到模板, 请上传模板!</p>")
		Else
%>
    <div class="setBody">
    	<ul>
<%
On Error Resume Next
For i = 1 To UBound(TempFolderArray)
	If fso.FileExists("../pjblog.template/" & TempFolderArray(i) & "/skin.xml") Then
		cxml.FilePath = "../pjblog.template/" & TempFolderArray(i) & "/skin.xml"
		If cxml.open Then
			' ---------------------------------------------
			' 	获取信息
			' ---------------------------------------------
			SkinName = cxml.GetNodeText(cxml.FindNode("//SkinSet/SkinName"))
				If Err Then Err.Clear : SkinName = ""
			SkinDesigner = cxml.GetNodeText(cxml.FindNode("//SkinSet/SkinDesigner"))
				If Err Then Err.Clear : SkinDesigner = ""
			pubDate = cxml.GetNodeText(cxml.FindNode("//SkinSet/pubDate"))
				If Err Then Err.Clear : pubDate = ""
			DesignerURL = cxml.GetNodeText(cxml.FindNode("//SkinSet/DesignerURL"))
				If Err Then Err.Clear : DesignerURL = ""
			DesignerMail = cxml.GetNodeText(cxml.FindNode("//SkinSet/DesignerMail"))
				If Err Then Err.Clear : DesignerMail = ""
			version = cxml.GetNodeText(cxml.FindNode("//SkinSet/version"))
				If Err Then Err.Clear : version = ""
			smallpic = cxml.GetNodeText(cxml.FindNode("//SkinSet/smallpic"))
				If Err Then Err.Clear : smallpic = ""
			bigpic = cxml.GetNodeText(cxml.FindNode("//SkinSet/bigpic"))
				If Err Then Err.Clear : bigpic = ""
			Intro = cxml.GetNodeText(cxml.FindNode("//SkinSet/Intro"))
				If Err Then Err.Clear : Intro = ""
%>
        	<li class="item">
            	<div class="skinbg">
                	<div class="top"></div>
                    <div class="mid">
                    	<div class="whole">
                        	<div class="tools">
                            	<div class="daochu" title="导出该主题"></div>
                                <div class="fabu" title="发布该主题"></div>
                                <div class="yulan" title="预览该主题"></div>
                                <div class="clear"></div>
                            </div>
                            <div class="themepic" onclick="conMain.skin.set.read.open(this, '<%=blog_DefaultSkin%>', '<%=blog_style%>', '<%=TempFolderArray(i)%>')">
                            	<div class="readinfo">
                                	<div><%=SkinName%></div>
                                    <div><%=SkinDesigner%></div>
                                    <div><%=pubDate%></div>
                                    <div><%=DesignerURL%></div>
                                    <div><%=DesignerMail%></div>
                                    <div><%=version%></div>
                                    <div><%=smallpic%></div>
                                    <div><%=bigpic%></div>
                                    <div><%=Intro%></div>
                                </div>
                                <div class="readStyle">
<%

'	遍历样式

	txml = fso.FolderItem("../pjblog.template/" & TempFolderArray(i) & "/style")
	txmls = Split(txml, "|")
	If txmls(0) > 0 Then
		For j = 1 To UBound(txmls)
			If fso.FileExists("../pjblog.template/" & TempFolderArray(i) & "/style/" & txmls(j) & "/style.xml") Then
				zoom.FilePath = "../pjblog.template/" & TempFolderArray(i) & "/style/" & txmls(j) & "/style.xml"
				If zoom.open Then
%>
									<div class="item">
                                    	<div>
                                        	<%
												temp = cxml.GetNodeText(cxml.FindNode("//SkinSet/SkinName"))
												If Err Then Err.Clear : temp = ""
												Response.Write(temp)
											%>
                                        </div>
                                        <div>
                                        	<%
												temp = cxml.GetNodeText(cxml.FindNode("//SkinSet/SkinDesigner"))
												If Err Then Err.Clear : temp = ""
												Response.Write(temp)
											%>
                                        </div>
                                        <div>
                                        	<%
												temp = cxml.GetNodeText(cxml.FindNode("//SkinSet/pubDate"))
												If Err Then Err.Clear : temp = ""
												Response.Write(temp)
											%>
                                        </div>
                                        <div>
                                        	<%
												temp = cxml.GetNodeText(cxml.FindNode("//SkinSet/DesignerURL"))
												If Err Then Err.Clear : temp = ""
												Response.Write(temp)
											%>
                                        </div>
                                        <div>
                                        	<%
												temp = cxml.GetNodeText(cxml.FindNode("//SkinSet/DesignerMail"))
												If Err Then Err.Clear : temp = ""
												Response.Write(temp)
											%>
                                        </div>
                                        <div>
                                        	<%
												temp = cxml.GetNodeText(cxml.FindNode("//SkinSet/version"))
												If Err Then Err.Clear : temp = ""
												Response.Write(temp)
											%>
                                        </div>
                                        <div>
                                        	<%
												temp = cxml.GetNodeText(cxml.FindNode("//SkinSet/smallpic"))
												If Err Then Err.Clear : temp = ""
												Response.Write(temp)
											%>
                                        </div>
                                        <div>
                                        	<%
												temp = cxml.GetNodeText(cxml.FindNode("//SkinSet/bigpic"))
												If Err Then Err.Clear : temp = ""
												Response.Write(temp)
											%>
                                        </div>
                                        <div>
                                        	<%
												temp = cxml.GetNodeText(cxml.FindNode("//SkinSet/Intro"))
												If Err Then Err.Clear : temp = ""
												Response.Write(temp)
											%>
                                        </div>
                                        <div><%=txmls(j)%></div>
                                    </div>
<%
				End If
			End If
		Next
	End If
%>
                                </div>
                            	<div class="img">
                                	<img src="../pjblog.template/<%=TempFolderArray(i)%>/<%=smallpic%>" onerror="this.src='../images/control/Preview.jpg'" />
                                </div>
                                <%If blog_DefaultSkin = TempFolderArray(i) Then%><div class="checked"></div><%End If%>
                            </div>
                            <div class="themetitle">
                            	<div class="themename"><%=SkinName%></div>
                            </div>
                        </div>
                    </div>
                    <div class="bom"></div>
                </div>
            </li>
<%
		End If
	End If
Next
%>
        </ul>
    </div>
<%
		End If
	Set zoom = Nothing
	Set cxml = Nothing
	Set fso = Nothing
%>
</div>
<%
End Sub
%>