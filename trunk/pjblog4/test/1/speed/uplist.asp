<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%OPTION EXPLICIT%>
<!--#include file="scripts/conn.asp"-->
<%
dim lngSize : lngSize=Request.QueryString("lngSize")
dim strFolder : strFolder = "UpLoadFile"
if isNumeric(lngSize) then
	lngSize=clng(lngSize)
else
	lngSize=0
end if
dim conn,sql,rs
set conn=Server.CreateObject("ADODB.Connection")
conn.Open strDconn
sql="select * from book where strSize1+strSize2>"&lngSize&" order by ID desc"
set rs=Server.CreateObject("ADODB.RecordSet")
rs.Open sql,conn,1,1
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>

<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>测试结果——风声 ASP 无组件上传类 V2.11 [Fonshen ASP UpLoadClass Version 2.11]</title>
<meta name="Keywords" content="ASP,无组件上传,上传组件,图片上传,风声,风声边界,Fonshen,Upload" />
<meta name="Description" content="风声 ASP 无组件上传类,Fonshen ASP UpLoadClass" />
<link href="../styles/works.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="scripts/function.js" charset="gb2312"></script>
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
			<a href="index.htm">速度测试</a>
			<a>测试结果</a>
		</div>
		<p id="Title">
			UpLoadClass对象[测试结果]
		</p>
		<div id="Content">
			<table width="100%" border="0" cellpadding="2" cellspacing="1" style="font-size:12px;text-align:center">
				<tr>
					<th width="80">时间</th>
					<th width="60">网友</th>
					<th>留言</th>
					<th width="140">上传文件/大小(Byte)</th>
					<th width="65">总用时(s)</th>
					<th width="65">速率(K/s)</th>
					<th width="100">上传IP</th>
				</tr>
				<%
				dim intRsCount,intPageSize,intPageCount,intPageNo
				intPageSize=20
				if pageDivide(intRsCount,intPageSize,intPageCount,intPageNo) then
					dim intTemp
					for intTemp=1 to intPageSize
						if rs.EOF then exit for
				%>
				<tr>
					<td width="80" title="<%=rs("datePost")%>"><%=datevalue(rs("datePost"))%></td>
					<td width="60"><%=rs("strUser")%></td>
					<td style="text-align:left;"><%=rs("strMessage")%></td>
					<td width="140">
					<%
						if rs("strSize1")<>0 then Response.Write("<a href='"&strFolder&"/"&rs("strFile1")&"' target='_blank' title='文件大小："&rs("strSize1")&" Bytes'>"&rs("strFile1")&"</a><br />")
						if rs("strSize2")<>0 then Response.Write("<a href='"&strFolder&"/"&rs("strFile2")&"' target='_blank' title='文件大小："&rs("strSize2")&" Bytes'>"&rs("strFile2")&"</a>")
					%>					</td>
					<td width="65"><%=rs("strTime")%></td>
					<td width="65"><%=rs("strSpeed")%></td>
					<td width="100"><%=rs("strIp")%></td>
				</tr>
				<%
						rs.MoveNext ()
					next
				end if
				%>
			</table>
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td><%call pageTurn(intRsCount,intPageSize,intPageCount,intPageNo)%></td>
				</tr>
			</table>
		</div>
	</div>

</body>
</html>
<%
rs.Close() : set rs=nothing
conn.Close() : set conn=nothing
%>