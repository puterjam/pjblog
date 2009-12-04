<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
'----------------------------------------------------------
'*************** 风声 ASP 无组件上传类 2.11 ***************
'速度测试
'本程序将保存您上传的文件，并记录上传速度及相关信息<br />
'上传速度主要由网速决定，测试结果仅供参考<br />
'该例以 gb2312（类默认，无须显示设置）字符集测试
'下面是上传程序(upload.asp)的代码和注释
'**********************************************************
'---------------------------------------------------------- 
OPTION EXPLICIT
Server.ScriptTimeOut=5000
%>
<!--#include file="scripts/conn.asp"-->
<!--#include file="../UpLoadClass.asp"-->
<!--#include file="../help/Code2Info.sc.asp"-->
<%
dim nTime : nTime = Timer()
dim myrequest,lngUpSize
Set myrequest=new UpLoadClass
	myrequest.TotalSize= 10485760
	myrequest.MaxSize  = 5000*1024
	myrequest.FileType = "rar/zip/gif/jpg/swf"
	myrequest.Savepath = "UpLoadFile/"
lngUpSize = myrequest.Open()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>

<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>速度测试——风声 ASP 无组件上传类 V2.11 [Fonshen ASP UpLoadClass Version 2.11]</title>
<meta name="Keywords" content="ASP,无组件上传,上传组件,图片上传,风声,风声边界,Fonshen,Upload" />
<meta name="Description" content="风声 ASP 无组件上传类,Fonshen ASP UpLoadClass" />
<link href="../styles/works.css" rel="stylesheet" type="text/css" />
</head>
<body id="UpLoadClass_V2.11">
	<div id="Context">
		<div id="Topic">
			风声 ASP 无组件上传类 V2.11 [Fonshen ASP UpLoadClass Version 2.11]
		</div>
		<div id="Menu">
			<a href="../index.htm">首 页</a>
			<a href="../help/help1.htm">对象参考</a>
			<a href="../help/help2.htm">专家说明</a>
			<a href="../demo1/index.htm">范例一</a>
			<a href="../demo2/index.htm">范例二</a>
			<a href="../demo3/index.htm">范例三</a>
			<a>速度测试</a>
			<a href="uplist.asp">测试结果</a>
		</div>
		<p id="Title">
			UpLoadClass对象[速度测试]
		</p>
		<div id="Content">
			<%
			Response.Write("<p id=""Title"">程序版本："&myrequest.Version&"</p>")
			Response.Write("<p>字 符 集："&myrequest.Charset&"<br />")	
			Response.Write "<span class=""Highlight"">处理状态："&Error2Info(myrequest.Error)&"</span></p>"
		
			dim strFile1 : strFile1=myrequest.Form("Photo")
			dim intError : intError=myrequest.Form("photo_Err")
			dim lngSize1 : lngSize1=0
			Response.Write("<p><strong>上传文件 1 信息</strong>")
			if intError>=0 then
				Response.Write "<br />文件名：<a href='"&myrequest.Savepath&strFile1&"' target='_blank'>"&strFile1&"</a>"
				Response.Write "<br />类　型："&myrequest.Form("Photo_Type")
				Response.Write "<br />源文件："&myrequest.Form("Photo_Name")
				Response.Write "<br />宽×高："&myrequest.Form("photo_Width")&"×"&myrequest.Form("photo_Height")
				lngSize1=myrequest.Form("photo_Size")
				if intError>0 then strFile1="Err:"&intError&"/"&myrequest.Form("photo_Name")
				Response.Write "<br />大　小："&lngSize1&" Bytes"
				Response.Write "<br />源路径："&myrequest.Form("photo_Path")
				Response.Write "<br />扩展名："&myrequest.Form("photo_Ext")
			end if
			Response.Write "<br />"&Err2Info(myrequest.Form("photo_Err"))
			Response.Write("</p>")
		
			dim strFile2 : strFile2=myrequest.Form("Photo2")
			intError=myrequest.Form("photo2_Err")
			dim lngSize2 : lngSize2=0
			Response.Write("<p><strong>上传文件 2 信息</strong>")
			if intError>=0 then
				Response.Write "<br />文件名：<a href='"&myrequest.Savepath&strFile2&"' target='_blank'>"&strFile2&"</a>"
				Response.Write "<br />类　型："&myrequest.Form("Photo2_Type")
				Response.Write "<br />源文件："&myrequest.Form("Photo2_Name")
				Response.Write "<br />宽×高："&myrequest.Form("photo2_Width")&"×"&myrequest.Form("photo2_Height")
				lngSize2=myrequest.Form("photo2_Size")
				if intError>0 then strFile2="Err:"&intError&"/"&myrequest.Form("photo2_Name")
				Response.Write "<br />大　小："&lngSize2&" Bytes"
				Response.Write "<br />源路径："&myrequest.Form("photo2_Path")
				Response.Write "<br />扩展名："&myrequest.Form("photo2_Ext")
			end if
			Response.Write "<br />"&Err2Info(myrequest.Form("photo2_Err"))
			Response.Write("</p>")
		
			dim strUser : strUser=replace(replace(left(myrequest.Form("intro"),8),"<","["),">","]")
			Response.Write "<p>网友大名："&strUser&"</p>"
			dim strMessage : strMessage=replace(replace(left(myrequest.Form("intro2"),30),"<","["),">","]")
			Response.Write "<p>网友留言："&strMessage&"</p>"
		
			Set myrequest=nothing
		
			nTime = Timer()-nTime
			dim dblTime : dblTime=FormatNumber(nTime,3)
			Response.Write "<p>上传用时："&dblTime&" s</p>"
			Response.Write("<p><a href='uplist.asp'>[查看上传历史纪录]</a></p>")
		
			' 上传处理完成
			' 开始数据库处理
		
			dim lngFileSize : lngFileSize = lngSize1+lngSize2
			if lngFileSize>0 then
				Response.Write "<p>平均上传速率："
				dim dblSpeed : dblSpeed = -1
				if nTime>0 then dblSpeed=FormatNumber(lngFileSize/(1024*nTime),3)
				'这里用lngUpSize更合理，lngUpSize由Open返回
				'if nTime>0 then dblSpeed=FormatNumber(lngUpSize/(1024*nTime),3)
				Response.Write dblSpeed&" k/s</p>"
				if strUser   ="" then strUser="匿名"
				if strMessage="" then strMessage=" "
				if strFile1  ="" then strFile1=" "
				if strFile2  ="" then strFile2=" "
				dim strIp : strIp=GetIp()
				strMessage=Replace(strMessage,chr(10),"")
			
				dim sql,conn
				set conn=Server.CreateObject("ADODB.Connection")
				conn.Open strDconn
				sql="Insert Into book"&_
					"(datePost,strUser,strMessage,strFile1,strSize1,strFile2,strSize2,strTime,strSpeed,strIp) "&_
					"values"&_
					"(#"&Now()&"#,'"&strUser&"','"&strMessage&"','"&strFile1&"',"&lngSize1&",'"&strFile2&"',"&lngSize2&",'"&dblTime&"','"&dblSpeed&"','"&strIp&"')"
				conn.Execute(sql)
				conn.Close()
				set conn=nothing
			end if
			%>
		</div>
	</div>
</body>
</html>