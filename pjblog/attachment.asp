<!--#include file="BlogCommon.asp" -->
<!--#include file="common/function.asp" -->
<!--#include file="common/library.asp" -->
<!--#include file="common/upfile.asp" -->
<!--#include file="common/cache.asp" -->
<!--#include file="common/checkUser.asp" -->
<%
'==================================
'  附件上传页面
'    更新时间: 2005-10-20
'==================================
On Error Resume Next
%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<link rel="stylesheet" rev="stylesheet" href="skins/<%=Skins%>/layout.css" type="text/css" media="all" /><!--层次样式表-->
<link rel="stylesheet" rev="stylesheet" href="skins/<%=Skins%>/typography.css" type="text/css" media="all" /><!--局部样式表-->
<link rel="stylesheet" rev="stylesheet" href="skins/<%=Skins%>/link.css" type="text/css" media="all" /><!--超链接样式表-->
<script type="text/javascript" src="common/common.js"></script>
</head>
<body class="attachmentBody">
<%
Server.ScriptTimeOut = 999
If stat_FileUpLoad = True And memName<>Empty Then
    If Request.QueryString("action") = "upload" Then
        Dim upl, FSOIsOK
        FSOIsOK = 1
        Set upl = Server.CreateObject("Scripting.FileSystemObject")
        If Err<>0 Then
            Err.Clear
            FSOIsOK = 0
        End If
        Dim D_Name, F_Name
        If FSOIsOK = 1 Then
            D_Name = "month_"&DateToStr(Now(), "ym")
            If upl.FolderExists(Server.MapPath("attachments/"&D_Name)) = False Then
                upl.CreateFolder Server.MapPath("attachments/"&D_Name)
            End If
        Else
            D_Name = "All_Files"
        End If
        Set upl = Nothing
        Dim FileUP
        Set FileUP = New Upload_File
        FileUP.GetDate( -1)
        Dim F_File, F_Type, sy, FilePath, AntiDown
        Set F_File = FileUP.File("File")
		sy = FileUp.Form("sy")
		AntiDown = FileUp.Form("AntiDown")
'blog_UpLoadSet = "0|0|0|PJBlog|PJBlog|0|1|10|10|FFFFFF|0|10|10|0.5|images/wind.png|120|35|www.pjhome.net|FFFFFF|18|宋体|1|0|000000|0|0"
'防盗链|文件命名|文件命名2|前缀|后缀|水印位置|计数边距|离左边距|离顶边距|边框颜色|边框宽度|水平边距|垂直边距|透明度|图片水印|图宽|图高|文字|字体颜色|字体大小|字体类型|加粗|斜体|阴影颜色|阴影向右偏移量|阴影向下偏移量
'   0      1        2      3    4     5       6        7      8        9      10         11      12     13      14     15  16  17     18      19      20      21  22     23       24            25
        Select Case Split(blog_UpLoadSet,"|")(1)
            Case 0		'上传日期时间
	           F_Name = Year(now)&lenNum(Month(now))&lenNum(Day(now))&lenNum(Hour(now))& lenNum(Minute(now))&lenNum(Second(now))
            Case 1		'文件原名
	           F_Name = getF_Name(F_File.FileName)
            Case 2		'上传日期时间_文件原名
	           F_Name = Year(now)&lenNum(Month(now))&lenNum(Day(now))&lenNum(Hour(now))& lenNum(Minute(now))&lenNum(Second(now))&"_"&getF_Name(F_File.FileName)
            Case 3		'上传日期时间带-
               F_Name = Year(now)&"-"&lenNum(Month(now))&"-"&lenNum(Day(now))&"-"&lenNum(Hour(now))&"-"&lenNum(Minute(now))&"-"&lenNum(Second(Now))&"-"&randomStr(1)
            Case Else	'原系统默认
	           F_Name = randomStr(1)&Year(Now)&Month(Now)&Day(Now)&Hour(Now)&Minute(Now)&Second(Now)
        End Select

        Select Case Split(blog_UpLoadSet,"|")(2)
            Case 0	'类型
	           F_Name = F_Name
            Case 1	'前缀_类型
	           F_Name = Split(blog_UpLoadSet,"|")(3)&"_"&F_Name
            Case 2	'类型_后缀
	           F_Name = F_Name&"_"&Split(blog_UpLoadSet,"|")(4)
            Case 3	'前缀_类型_后缀
	           F_Name = Split(blog_UpLoadSet,"|")(3)&"_"&F_Name&"_"& Split(blog_UpLoadSet,"|")(4)
        End Select

        F_Name = F_Name&"."&F_File.FileExt
        F_Type = FixName(F_File.FileExt)
        If F_File.FileSize > Int(UP_FileSize) Then
            Response.Write("<div style=""padding:6px""><a href='attachment.asp'>文件大小超出，请返回重新上传</a></div>")
        ElseIf IsvalidFile(UCase(F_Type)) = False Then
            Response.Write("<div style=""padding:6px""><a href='attachment.asp'>文件格式非法，请返回重新上传</a></div>")
        Else
		    FilePath = "attachments/"&D_Name&"/"&F_Name
            F_File.SaveAs Server.MapPath(FilePath)
			If UCase(F_Type) = "JPG" Or UCase(F_Type) = "BMP" Or UCase(F_Type) = "GIF" Then
			  If sy <> "" And IsNumeric(sy) Then Call CreateView(FilePath,sy,blog_UpLoadSet)
		    End If
			Conn.Execute("INSERT INTO blog_Files (FilesPath) VALUES ("""&FilePath&""")")
            Dim UploadID : UploadID = Conn.Execute("SELECT ID FROM blog_Files order by ID desc")(0)
			If AntiDown = "1" then UploadID = UploadID&"&code="&right(md5(right(Ucase(FilePath),15)),10)
            response.Write "<script>addUploadItem('"&F_Type&"','download.asp?id="&UploadID&"',"&Request.QueryString("MSave")&")</script>"
            Response.Write("<div style=""padding:6px""><a href='attachment.asp'>文件上传成功，请返回继续上传</a></div>")
        End If
        Set F_File = Nothing
        Set FileUP = Nothing
        Response.Write("</td>")
    Else
%>
<script>
 function MSave(o){
  if (o.checked) {
    document.forms["frm"].action="attachment.asp?action=upload&MSave=1"
  }
  else
  {
    document.forms["frm"].action="attachment.asp?action=upload&MSave=0"
  }
 }
</script>
<form name="frm" enctype="multipart/form-data" method="post" action="attachment.asp?action=upload&MSave=0"><input name="File" type="File" size="28" style="font-size:12px;border-width:1px">&nbsp;<input type="Submit" name="Submit" value="确定上传" class="userbutton"><input type="checkbox" name="MemberDown" value="1" id="Md" onclick="MSave(this)" title="只对UBB编辑有效,媒体文件包括图片无效"/><label for="Md" title="只对UBB编辑有效,媒体文件包括图片无效">此文件只允许会员下载 </label> <select name="sy" style="font-size:12px;border-width:1px"<%If not CheckObjInstalled("Persits.Jpeg") Then Response.Write (" disabled=""disabled""")%>><option selected>水印类型</option><option value="1">文字水印</option><option value="2">图片水印</option></select><input type="checkbox" name="AntiDown" id="AntiDown" value="1"/><label for="AntiDown">允许盗链 </label></form>
<%
End If
Else
    Response.Write("<div style=""padding:6px;color:#f00"">对不起，你没有权限上传附件！</div>")
End If
%>
