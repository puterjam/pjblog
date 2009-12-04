<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'----------------------------------------------------------
'************** 风声 ASP 无组件上传类 V2.11 ***************
'用法举例：快速应用[添加产品一]
'该例主要说明默认模式下的运用
'以常见的产品更新为例<br>
'该例以UTF-8字符集测试
'下面是上传程序(upload.asp)的代码和注释
'**********************************************************
'---------------------------------------------------------- 
OPTION EXPLICIT
Server.ScriptTimeOut=5000
%>
<!--#include file="../UpLoadClass.asp"-->
<%
dim request2

'建立上传对象
set request2=New UpLoadClass

	'设置字符集
	request2.Charset="UTF-8"

	'打开对象
	request2.Open()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>

<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>快速应用——风声 ASP 无组件上传类 V2.11 [Fonshen ASP UpLoadClass Version 2.11]</title>
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
			<a>范例一</a>
			<a href="../demo2/index.htm">范例二</a>
			<a href="../demo3/index.htm">范例三</a>
			<a href="../speed/index.htm">速度测试</a>
			<a href="../speed/uplist.asp">测试结果</a>
		</div>
		<p id="Title">
			范例一：快速应用[添加产品一]
		</p>
		<div id="Content">
			<%
			'显示类版本
			Response.Write("<p>程序版本："&request2.Version&"</p>")
			
			'显示字符集
			Response.Write("<p id=""Title"">字 符 集："&request2.Charset&"</p>")
			
			'显示产品名称
			Response.Write("<p>产品名称："&request2.Form("strName")&"</p>")
			
			'显示源文件路径与名称
			Response.Write("<p>产品图片："&request2.Form("strPhoto_Path")&request2.Form("strPhoto_Name"))
			Response.Write("=>")
			
			'显示目标文件路径与名称
			Response.Write(request2.SavePath&request2.Form("strPhoto")&"</p>")
			
			'显示产品介绍
			Response.Write("<p>产品介绍："&request2.Form("strRemark")&"</p>")
			
			'-------说明开始------
			'可以看出上面的显示是淋漓尽致的
			'文件是否需要保存由类自动判断，这已经符合大多数情况下的应用
			'如果您需要更灵活的处理,参见[进阶应用]
			'-------说明结束------
			
			Response.Write "<p>[<a href=""javascript:history.back();"">返回</a>]</p>"
			%>
		</div>
		</div>
</body>
</html>
<%
'释放上传对象
set request2=nothing
%>