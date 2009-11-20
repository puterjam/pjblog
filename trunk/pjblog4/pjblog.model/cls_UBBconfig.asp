<%
'|===========================|
'|   UBB编辑器 2.0           |
'|      作者:舜子(PuterJam)  |
'|   版权所有 2008          |
'|===========================|
Dim UBB_TextArea_Height, UBB_Tools_Items, UBB_Tools_Fonts, UBB_Tools_Size, UBB_Tools_Color, UBB_Msg_Value, UBB_AutoHidden

'-------默认工具条------
Dim UBB_Tools_default
UBB_Tools_default = ""
UBB_Tools_default = UBB_Tools_default&"bold,italic,underline,deleteline,justifyleft,justifycenter,justifyright"
UBB_Tools_default = UBB_Tools_default&"||"
UBB_Tools_default = UBB_Tools_default&"link,mail,image,insertunorderedlist,quote,hidden,code,html"
UBB_Tools_default = UBB_Tools_default&"||"
UBB_Tools_default = UBB_Tools_default&"smiley"
UBB_Tools_default = UBB_Tools_default&"/"
UBB_Tools_default = UBB_Tools_default&"flash,music,mediaplayer,realplayer,ed2k,mDown,htmlubb,highlightcode,AjaxLogSave"
UBB_Tools_default = UBB_Tools_default&"||"
UBB_Tools_default = UBB_Tools_default&"about"

UBB_Tools_default = UBB_Tools_default&"/"
UBB_Tools_default = UBB_Tools_default&"fonts,sizes,color"
UBB_Tools_default = UBB_Tools_default&"||"
UBB_Tools_default = UBB_Tools_default&"method"


'-------UBB编辑器基本配置-------
UBB_Tools_Items = UBB_Tools_default
UBB_TextArea_Height = ""
UBB_Msg_Value = ""
UBB_AutoHidden = True

UBB_Tools_Fonts = "宋体,黑体,Arial,Book Antiqua,Century Gothic,Courier New,Georgia,Impact,Tahoma,Times New Roman,Verdana"
UBB_Tools_Size = "8,9,10,11,12,13,14,15,16,18,20,24,36,48"
UBB_Tools_Color = "White,Black,Red,Yellow,Pink,Green,Orange,Purple,Blue,Beige,Brown,Teal,Navy,Maroon,LimeGreen"


Sub UBBeditor(TextName)
	Response.write (UBBeditorCore(TextName))
End Sub

Function UBBeditorCore(TextName)
    Dim TempStyle
    TempStyle = " style="""
    UBBeditorCore = ""
    
    If UBB_TextArea_Height<>"" Then 
   	 TempStyle = TempStyle&"height:"&UBB_TextArea_Height&";"
    End IF
    
    TempStyle = TempStyle&""""
    
    UBBeditorCore = UBBeditorCore&"<script language=""javascript"" type=""text/javascript"" src=""../pjblog.common/UBBCode.js""></script>"
    UBBeditorCore = UBBeditorCore&"<script language=""javascript"" type=""text/javascript"" src=""../pjblog.common/UBBCode_help.js""></script>"
    

    
    If UBB_AutoHidden Then
   		UBBeditorCore = UBBeditorCore&"<div id=""UBBSmiliesPanel"" class=""UBBSmiliesPanel""></div>"
        UBBeditorCore = UBBeditorCore&"<textarea id=""editMask"" class=""editTextarea"" style=""width:99%;height:100px"" onfocus=""loadUBB('"&TextName&"')"">"&UBB_Msg_Value&"</textarea><div id=""editorbody"" style=""display:none""><div id=""editorHead"">正在加载编辑器...</div>"
    Else
   		UBBeditorCore = UBBeditorCore&"<div id=""UBBSmiliesPanel"" class=""UBBSmiliesPanel"">"&showSmilie&"</div>"
        UBBeditorCore = UBBeditorCore&"<div id=""editorbody""><div id=""editorHead"">"&ToolsToCode&"</div>"
    End If
    
    UBBeditorCore = UBBeditorCore&"<div class=""editorContent""><textarea name="""&TextName&""" class=""editTextarea"""&TempStyle&" cols=""1"" rows=""1"" accesskey=""R"" id=""" & TextName & """>"&UBB_Msg_Value&"</textarea></div></div>"
    UBBeditorCore = UBBeditorCore&"<script language=""javascript"" type=""text/javascript"">setTimeout(""initUBB(\"""&TextName&"\"")"", 100)</script>"
End Function

Function showSmilie
	Dim Arr_Smilies, Arr_Smilie, SmilieItem, SmilieHtml, SmilieDr, SmilieCount
    SmilieHtml = ""
    SmilieDr = ""
    SmilieCount = 0
    
    Arr_Smilies = Cache.SmiliesCache(1)
    For Each Arr_Smilie in Arr_Smilies
        SmilieItem = Split(Arr_Smilie, "|")
        SmilieCount = SmilieCount + 1
        If SmilieCount = 1 Then SmilieHtml = SmilieHtml + "<tr>"
        SmilieHtml = SmilieHtml + "<td><a href=""javascript:AddSmiley('"&SmilieItem(2)&"')"" class=""Smilie"" title="""&SmilieItem(2)&"""><img border=""0"" src=""../images/smilies/"&SmilieItem(1)&""" alt=""""/></a></td>"
        If SmilieCount = 8 Then
            SmilieHtml = SmilieHtml + "</tr>"
            SmilieCount = 0
        End If
        SmilieDr = ""
    Next
    
    showSmilie = "<table cellspacing=""2"" cellpadding=""0"">" & SmilieHtml & "</table>"
End Function

Function ToolsToCode()
    Dim Toolsbar, barItems, barItem, ItemButtons, ItemButton, Items
    Dim UBB_fonts, UBB_sizes, UBB_colors, UBB_font, UBB_size, UBB_color
    ToolsToCode = ""
    Toolsbar = Split(UBB_Tools_Items, "/")
    For Each barItems in Toolsbar
        ToolsToCode = ToolsToCode&"<div class=""editorTools"">"
        barItem = Split(barItems, "||")
        For Each ItemButtons in barItem
            ToolsToCode = ToolsToCode&"<div class=""Toolsbar""><ul class=""ToolsUL"">"
            ItemButton = Split(ItemButtons, ",")
            For Each Items in ItemButton
                Select Case Items
                    Case "fonts"
                        UBB_fonts = Split(UBB_Tools_Fonts, ",")
                        ToolsToCode = ToolsToCode&"<li><select name=""UBBfonts"" onchange=""UBB_CFont(this)""><option value="""" selected=""selected"">选择字体</option>"
                        For Each UBB_font in UBB_fonts
                            ToolsToCode = ToolsToCode&"<option value="""&UBB_font&""">"&UBB_font&"</option>"
                        Next
                        ToolsToCode = ToolsToCode&"</select></li>"
                    Case "sizes"
                        UBB_sizes = Split(UBB_Tools_Size, ",")
                        ToolsToCode = ToolsToCode&"<li><select name=""UBBfonts"" onchange=""UBB_CFontSize(this)""><option value="""" selected=""selected"">字体大小</option>"
                        For Each UBB_size in UBB_sizes
                            ToolsToCode = ToolsToCode&"<option value="""&UBB_size&""">"&UBB_size&"</option>"
                        Next
                        ToolsToCode = ToolsToCode&"</select></li>"
                    Case "color"
                        UBB_colors = Split(UBB_Tools_Color, ",")
                        ToolsToCode = ToolsToCode&"<li><select name=""UBBfonts"" onchange=""UBB_CFontColor(this)""><option value="""" selected=""selected"">字体颜色</option>"
                        For Each UBB_color in UBB_colors
                            ToolsToCode = ToolsToCode&"<option value="""&UBB_color&""" style=""background:"&UBB_color&""">"&UBB_color&"</option>"
                        Next
                        ToolsToCode = ToolsToCode&"</select></li>"
                    Case "hidden"
                        ToolsToCode = ToolsToCode&"<li><a id=""A_hidden"" href=""javascript:UBB_hidden();void(0)"" title=""插入只允许会员查看的隐藏内容"" class=""Toolsbutton""><img src=""../images/UBB/pager.gif"" border=""0"" alt=""插入只允许会员查看的隐藏内容"" /></a></li>"
                    Case "method"
                        ToolsToCode = ToolsToCode&"<li>UBB编辑模式 <label for=""UBBmethod1""><input id=""UBBmethod1"" name=""UBBmethod"" type=""radio"" checked=""checked"" onclick=""EditMethod='normal'""/>通常</label> <label for=""UBBmethod2""><input id=""UBBmethod2"" name=""UBBmethod"" type=""radio"" onclick=""EditMethod='expert'""/>专家</label>"
                    Case "ed2k"
                        ToolsToCode = ToolsToCode&"<li><a id=""A_ed2k"" href=""javascript:UBB_ed2k();void(0)"" title=""插入eMule超链接"" class=""Toolsbutton""><img src=""../images/UBB/ed2k.gif"" border=""0"" alt=""插入eMule超链接"" /></a></li>"
                    Case "mDown"
                        ToolsToCode = ToolsToCode&"<li><a id=""A_mDown"" href=""javascript:UBB_mDown();void(0)"" title=""插入只允许会员下载的文件地址"" class=""Toolsbutton""><img src=""../images/UBB/disk.gif"" border=""0"" alt=""插入只允许会员下载的文件地址"" /></a></li>"
					Case "AjaxLogSave"
                        ToolsToCode = ToolsToCode&"<li><a id=""A_AjaxLogSave"" href=""javascript:UBB_AjaxLogSave();void(0)"" title=""点击开启或关闭Ajax草稿自动保存"" class=""Toolsbutton""><img src=""../images/ubb/AjaxLogSave.png"" border=""0"" alt=""点击开启或关闭Ajax草稿自动保存"" /></a></li>"
                    Case Else
                        ToolsToCode = ToolsToCode&"<li><a id=""A_"&Items&""" href=""javascript:UBB_"&Items&"();void(0)"" title=""" + Tip(Items) + """ class=""Toolsbutton""><img src=""../pjblog.template/" & blog_DefaultSkin & "/UBB/Icons/"&Items&".gif"" border=""0"" alt=""" + Tip(Items) + """ /></a></li>"
                End Select
            Next
            ToolsToCode = ToolsToCode&"</ul></div>"
        Next
        ToolsToCode = ToolsToCode&"<div style=""clear: both;display: block;height:1px;overflow:hidden""></div></div>"
    Next
	ToolsToCode = ToolsToCode&"<div id=""AjaxTimeSave"" style=""display:none; width:100%""></div>"
End Function

Function Tip(Str)
    Select Case Str
Case "about":
        Tip = "关于"
Case "bold":
        Tip = "粗体"
Case "deleteline":
        Tip="删除线"
Case "code":
        Tip = "插入代码"
Case "flash":
        Tip = "插入Flash动画"
Case "highlightcode":
        Tip = "幻码,加亮代码"
Case "html":
        Tip = "插入可执行html代码"
Case "htmlubb":
        Tip = "HTML转贴能手"
Case "image":
        Tip = "插入图像"
Case "insertunorderedlist":
        Tip = "项目符号和编号"
Case "italic":
        Tip = "斜体"
Case "justifycenter":
        Tip = "居中"
Case "justifyleft":
        Tip = "居左"
Case "justifyright":
        Tip = "居右"
Case "link":
        Tip = "超链接"
Case "mail":
        Tip = "电子邮件"
Case "mediaplayer":
        Tip = "插入视频"
Case "music":
        Tip = "插入音频"
Case "quote":
        Tip = "插入引用"
Case "realplayer":
        Tip = "插入Real媒体"
Case "underline":
        Tip = "下划线"
Case "smiley":
        Tip = "表情符号"
Case Else:
        Tip = ""
    End Select
End Function
%>
