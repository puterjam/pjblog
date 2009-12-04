<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
'----------------------------------------------------------
'*************** 风声 ASP 无组件上传类 V2.11 **************
'用法举例：进阶应用[添加产品二]
'该例主要说明手动保存模式更灵活的运用
'以常见的产品更新为例<br>
'该例以 gb2312（类默认，无须显示设置）字符集测试
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

	'设置为手动保存模式
	request2.AutoSave=2

	'打开对象，默认为 gb2312 字符集，故没有显示设置
	request2.Open()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>

<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>进阶应用——风声 ASP 无组件上传类 V2.11 [Fonshen ASP UpLoadClass Version 2.11]</title>
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
			<a>范例二</a>
			<a href="../demo3/index.htm">范例三</a>
			<a href="../speed/index.htm">速度测试</a>
			<a href="../speed/uplist.asp">测试结果</a>
		</div>
		<p id="Title">
			范例二：进阶应用[添加产品二]
		</p>
		<div id="Content">
			<%
			'显示类版本
			Response.Write("<p>程序版本："&request2.Version&"</p>")
			
			'显示字符集
			Response.Write("<p id=""Title"">字 符 集："&request2.Charset&"</p>")
			
			'显示产品名称
			Response.Write("<p>产品名称："&request2.Form("strName")&"</p>")
			
			'设置产品小图最大为10K
			'任何时候都可以重设参数，重设参数对后面的save/open有效
			'不过第二次调用open什么都不做
			request2.MaxSize=10240
			
			'如果保存小图成功，系统生成目标文件名
			if request2.Save("strPhoto1",0) then
		
				'显示保存位置
				Response.Write("<p>产品小图=&gt;"&request2.SavePath&request2.Form("strPhoto1")&"</p>")
			end if
		
			'设置产品大图最大为150K
			'当然这里还可以设置不同的保存路径，限制格式
			request2.MaxSize=153600
			
			'如果保存大图成功
			Response.Write("")
			if request2.Save("strPhoto2",0) then
		
				'显示保存位置
				Response.Write("<p>产品大图=&gt;"&request2.SavePath&request2.Form("strPhoto2")&"</p>")
			end if
			
			'显示产品介绍
			Response.Write("<p>产品介绍："&request2.Form("strRemark")&"</p>")
			
			'-------说明开始------
			'如果是正确的jpg/gig/png/bmp图片格式文件，还可以获取图片宽高
			'if request2.form("strPhoto2_Width")<>"" then
			'Response.Write("图片宽："&request2.form("strPhoto2_Width"))
			'Response.Write("图片高："&request2.form("strPhoto2_Height"))
			'end if
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