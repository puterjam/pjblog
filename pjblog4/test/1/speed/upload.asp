<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
'----------------------------------------------------------
'*************** ���� ASP ������ϴ��� 2.11 ***************
'�ٶȲ���
'�����򽫱������ϴ����ļ�������¼�ϴ��ٶȼ������Ϣ<br />
'�ϴ��ٶ���Ҫ�����پ��������Խ�������ο�<br />
'������ gb2312����Ĭ�ϣ�������ʾ���ã��ַ�������
'�������ϴ�����(upload.asp)�Ĵ����ע��
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
<title>�ٶȲ��ԡ������� ASP ������ϴ��� V2.11 [Fonshen ASP UpLoadClass Version 2.11]</title>
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
			<a href="../demo2/index.htm">������</a>
			<a href="../demo3/index.htm">������</a>
			<a>�ٶȲ���</a>
			<a href="uplist.asp">���Խ��</a>
		</div>
		<p id="Title">
			UpLoadClass����[�ٶȲ���]
		</p>
		<div id="Content">
			<%
			Response.Write("<p id=""Title"">����汾��"&myrequest.Version&"</p>")
			Response.Write("<p>�� �� ����"&myrequest.Charset&"<br />")	
			Response.Write "<span class=""Highlight"">����״̬��"&Error2Info(myrequest.Error)&"</span></p>"
		
			dim strFile1 : strFile1=myrequest.Form("Photo")
			dim intError : intError=myrequest.Form("photo_Err")
			dim lngSize1 : lngSize1=0
			Response.Write("<p><strong>�ϴ��ļ� 1 ��Ϣ</strong>")
			if intError>=0 then
				Response.Write "<br />�ļ�����<a href='"&myrequest.Savepath&strFile1&"' target='_blank'>"&strFile1&"</a>"
				Response.Write "<br />�ࡡ�ͣ�"&myrequest.Form("Photo_Type")
				Response.Write "<br />Դ�ļ���"&myrequest.Form("Photo_Name")
				Response.Write "<br />����ߣ�"&myrequest.Form("photo_Width")&"��"&myrequest.Form("photo_Height")
				lngSize1=myrequest.Form("photo_Size")
				if intError>0 then strFile1="Err:"&intError&"/"&myrequest.Form("photo_Name")
				Response.Write "<br />��С��"&lngSize1&" Bytes"
				Response.Write "<br />Դ·����"&myrequest.Form("photo_Path")
				Response.Write "<br />��չ����"&myrequest.Form("photo_Ext")
			end if
			Response.Write "<br />"&Err2Info(myrequest.Form("photo_Err"))
			Response.Write("</p>")
		
			dim strFile2 : strFile2=myrequest.Form("Photo2")
			intError=myrequest.Form("photo2_Err")
			dim lngSize2 : lngSize2=0
			Response.Write("<p><strong>�ϴ��ļ� 2 ��Ϣ</strong>")
			if intError>=0 then
				Response.Write "<br />�ļ�����<a href='"&myrequest.Savepath&strFile2&"' target='_blank'>"&strFile2&"</a>"
				Response.Write "<br />�ࡡ�ͣ�"&myrequest.Form("Photo2_Type")
				Response.Write "<br />Դ�ļ���"&myrequest.Form("Photo2_Name")
				Response.Write "<br />����ߣ�"&myrequest.Form("photo2_Width")&"��"&myrequest.Form("photo2_Height")
				lngSize2=myrequest.Form("photo2_Size")
				if intError>0 then strFile2="Err:"&intError&"/"&myrequest.Form("photo2_Name")
				Response.Write "<br />��С��"&lngSize2&" Bytes"
				Response.Write "<br />Դ·����"&myrequest.Form("photo2_Path")
				Response.Write "<br />��չ����"&myrequest.Form("photo2_Ext")
			end if
			Response.Write "<br />"&Err2Info(myrequest.Form("photo2_Err"))
			Response.Write("</p>")
		
			dim strUser : strUser=replace(replace(left(myrequest.Form("intro"),8),"<","["),">","]")
			Response.Write "<p>���Ѵ�����"&strUser&"</p>"
			dim strMessage : strMessage=replace(replace(left(myrequest.Form("intro2"),30),"<","["),">","]")
			Response.Write "<p>�������ԣ�"&strMessage&"</p>"
		
			Set myrequest=nothing
		
			nTime = Timer()-nTime
			dim dblTime : dblTime=FormatNumber(nTime,3)
			Response.Write "<p>�ϴ���ʱ��"&dblTime&" s</p>"
			Response.Write("<p><a href='uplist.asp'>[�鿴�ϴ���ʷ��¼]</a></p>")
		
			' �ϴ��������
			' ��ʼ���ݿ⴦��
		
			dim lngFileSize : lngFileSize = lngSize1+lngSize2
			if lngFileSize>0 then
				Response.Write "<p>ƽ���ϴ����ʣ�"
				dim dblSpeed : dblSpeed = -1
				if nTime>0 then dblSpeed=FormatNumber(lngFileSize/(1024*nTime),3)
				'������lngUpSize������lngUpSize��Open����
				'if nTime>0 then dblSpeed=FormatNumber(lngUpSize/(1024*nTime),3)
				Response.Write dblSpeed&" k/s</p>"
				if strUser   ="" then strUser="����"
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