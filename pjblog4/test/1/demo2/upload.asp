<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
'----------------------------------------------------------
'*************** ���� ASP ������ϴ��� V2.11 **************
'�÷�����������Ӧ��[��Ӳ�Ʒ��]
'������Ҫ˵���ֶ�����ģʽ����������
'�Գ����Ĳ�Ʒ����Ϊ��<br>
'������ gb2312����Ĭ�ϣ�������ʾ���ã��ַ�������
'�������ϴ�����(upload.asp)�Ĵ����ע��
'**********************************************************
'---------------------------------------------------------- 
OPTION EXPLICIT
Server.ScriptTimeOut=5000
%>
<!--#include file="../UpLoadClass.asp"-->
<%
dim request2 
'�����ϴ�����
set request2=New UpLoadClass

	'����Ϊ�ֶ�����ģʽ
	request2.AutoSave=2

	'�򿪶���Ĭ��Ϊ gb2312 �ַ�������û����ʾ����
	request2.Open()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>

<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>����Ӧ�á������� ASP ������ϴ��� V2.11 [Fonshen ASP UpLoadClass Version 2.11]</title>
<meta name="Keywords" content="ASP,������ϴ�,�ϴ����,ͼƬ�ϴ�,����,�����߽�,Fonshen,Upload" />
<meta name="Description" content="���� ASP ������ϴ���,Fonshen ASP UpLoadClass" />
<link href="../styles/works.css" rel="stylesheet" type="text/css" />
</head>
<body id="UpLoadClass_V2.11">
	<div id="Context">
		<div id="Topic">
			���� ASP ������ϴ��� V2.11 [Fonshen ASP UpLoadClass Version 2.11]
		</div>
		<div id="Menu">
			<a href="../index.htm">�� ҳ</a>
			<a href="../help/help1.htm">����ο�</a>
			<a href="../help/help2.htm">ר��˵��</a>
			<a href="../demo1/index.htm">����һ</a>
			<a>������</a>
			<a href="../demo3/index.htm">������</a>
			<a href="../speed/index.htm">�ٶȲ���</a>
			<a href="../speed/uplist.asp">���Խ��</a>
		</div>
		<p id="Title">
			������������Ӧ��[��Ӳ�Ʒ��]
		</p>
		<div id="Content">
			<%
			'��ʾ��汾
			Response.Write("<p>����汾��"&request2.Version&"</p>")
			
			'��ʾ�ַ���
			Response.Write("<p id=""Title"">�� �� ����"&request2.Charset&"</p>")
			
			'��ʾ��Ʒ����
			Response.Write("<p>��Ʒ���ƣ�"&request2.Form("strName")&"</p>")
			
			'���ò�ƷСͼ���Ϊ10K
			'�κ�ʱ�򶼿��������������������Ժ����save/open��Ч
			'�����ڶ��ε���openʲô������
			request2.MaxSize=10240
			
			'�������Сͼ�ɹ���ϵͳ����Ŀ���ļ���
			if request2.Save("strPhoto1",0) then
		
				'��ʾ����λ��
				Response.Write("<p>��ƷСͼ=&gt;"&request2.SavePath&request2.Form("strPhoto1")&"</p>")
			end if
		
			'���ò�Ʒ��ͼ���Ϊ150K
			'��Ȼ���ﻹ�������ò�ͬ�ı���·�������Ƹ�ʽ
			request2.MaxSize=153600
			
			'��������ͼ�ɹ�
			Response.Write("")
			if request2.Save("strPhoto2",0) then
		
				'��ʾ����λ��
				Response.Write("<p>��Ʒ��ͼ=&gt;"&request2.SavePath&request2.Form("strPhoto2")&"</p>")
			end if
			
			'��ʾ��Ʒ����
			Response.Write("<p>��Ʒ���ܣ�"&request2.Form("strRemark")&"</p>")
			
			'-------˵����ʼ------
			'�������ȷ��jpg/gig/png/bmpͼƬ��ʽ�ļ��������Ի�ȡͼƬ���
			'if request2.form("strPhoto2_Width")<>"" then
			'Response.Write("ͼƬ��"&request2.form("strPhoto2_Width"))
			'Response.Write("ͼƬ�ߣ�"&request2.form("strPhoto2_Height"))
			'end if
			'-------˵������------
			
			Response.Write "<p>[<a href=""javascript:history.back();"">����</a>]</p>"
			%>
		</div>
		
	</div>


</body>
</html>
<%
'�ͷ��ϴ�����
set request2=nothing
%> 