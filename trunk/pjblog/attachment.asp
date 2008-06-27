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
        Dim F_File, F_Type
        Set F_File = FileUP.File("File")
        F_Name = randomStr(1)&Year(Now)&Month(Now)&Day(Now)&Hour(Now)&Minute(Now)&Second(Now)&"."&F_File.FileExt
        F_Type = FixName(F_File.FileExt)
        If F_File.FileSize > Int(UP_FileSize) Then
            Response.Write("<div style=""padding:6px""><a href='attachment.asp'>文件大小超出，请返回重新上传</a></div>")
        ElseIf IsvalidFile(UCase(F_Type)) = False Then
            Response.Write("<div style=""padding:6px""><a href='attachment.asp'>文件格式非法，请返回重新上传</a></div>")
        Else
            F_File.SaveAs Server.MapPath("attachments/"&D_Name&"/"&F_Name)
            response.Write "<script>addUploadItem('"&F_Type&"','attachments/"&D_Name&"/"&F_Name&"',"&Request.QueryString("MSave")&")</script>"
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
<%
Response.Write("<form name=""frm"" enctype=""multipart/form-data"" method=""post"" action=""attachment.asp?action=upload&MSave=0""><input name=""File"" type=""File"" size=""28"" style=""font-size:12px;border-width:1px"">&nbsp;<input type=""Submit"" name=""Submit"" value=""确定上传"" class=""userbutton""><input type=""checkbox"" name=""MemberDown"" value=""1"" id=""Md"" onclick=""MSave(this)"" title=""只对UBB编辑有效,媒体文件包括图片无效""/><label for=""Md"" title=""只对UBB编辑有效,媒体文件包括图片无效"">此文件只允许会员下载 </label></form>")
End If
Else
    Response.Write("<div style=""padding:6px;color:#f00"">对不起，你没有权限上传附件！</div>")
End If
%>
