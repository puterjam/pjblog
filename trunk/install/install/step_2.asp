<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>安装程序 - PJblog - www.pjhome.net</title>
<link href="style.css" rel="stylesheet" type="text/css" />
</head>
<!--#include file = "misc.asp" -->
<body>
<%
	Session("con_db") = ""
	Response.Cookies("InstallType") = ""
	Response.Cookies("InstallType").Expires = DateAdd("d", 1, now())
	Response.Cookies("insName") = ""
	Response.Cookies("insName").Expires = DateAdd("d", 1, now())
	Response.Cookies("InstallCookie") = ""
	Response.Cookies("InstallCookie").Expires = DateAdd("d", 1, now())
	Response.Cookies("InstallAccess") = ""
	Response.Cookies("InstallAccess").Expires = DateAdd("d", 1, now())
%>
<div class="main">
	<div class="pleft">
		<dl class="setpbox t1">
			<dt>安装步骤</dt>
			<dd>
				<ul>
					<li class="succeed">许可协议</li>
					<li class="now">环境检测</li>
					<li>参数配置</li>
					<li>正在安装</li>
					<li>安装完成</li>
				</ul>
			</dd>
		</dl>
	</div>
	<div class="pright">
		<div class="pr-title"><h3>服务器信息</h3></div>
		<table width="726" border="0" align="center" cellpadding="0" cellspacing="0" class="twbox">
			<tr>
				<th width="300" align="center"><strong>参数</strong></th>
				<th width="424"><strong>值</strong></th>
			</tr>
			<tr align="center">
					<td><strong>服务器域名</strong></td>
					<td><%=Request.ServerVariables("LOCAL_ADDR")%></td>
			</tr>
			<tr align="center">
					<td><strong>服务器操作系统</strong></td>
					<td><%=Request.ServerVariables("OS")%></td>
			</tr>
			<tr align="center">
					<td><strong>服务器解译引擎</strong></td>
					<td><%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
			</tr>
			<tr align="center">
					<td><strong>IIS版本</strong></td>
					<td><%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
			</tr>
			<tr align="center">
					<td><strong>系统安装目录</strong></td>
					<td><%=Request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
			</tr>
		</table>
		<div class="pr-title"><h3>系统环境检测</h3></div>
		<div style="padding:2px 8px 0px; line-height:33px; height:23px; overflow:hidden; color:#666;">
			系统环境要求必须满足下列所有条件/组件，否则系统或系统部份功能将无法使用。
		</div>
		<table width="726" border="0" align="center" cellpadding="0" cellspacing="0" class="twbox">
			<tr>
				<th width="200" align="center"><strong>需要支持的服务器组件</strong></th>
				<th width="80"><strong>要求</strong></th>
				<th width="400"><strong>实际状态及建议</strong></th>
			</tr>
			<tr>
					<td>Scripting.FileSystemObject</td>
					<td align="center">On </td>
					<td><%=DisPlay(CheckObjInstalled("Scripting.FileSystemObject"))%> <small>(不符合要求将导致无法实现静态化等文件和文件夹操作)</small></td>
			</tr>
			<tr>
					<td>MSXML2.ServerXMLHTTP</td>
					<td align="center">On</td>
					<td><%=DisPlay(CheckObjInstalled("MSXML2.ServerXMLHTTP"))%> <small>(不符合要求将无法使用某些功能)</small></td>
			</tr>
			
			<tr>
					<td>Microsoft.XMLDOM</td>
					<td align="center">On</td>
					<td><%=DisPlay(CheckObjInstalled("Microsoft.XMLDOM"))%> <small>(不支持将导致无法使用ASP操作XML文件)</small></td>
			</tr>
			<tr>
					<td>ADODB.Stream</td>
					<td align="center">On</td>
					<td><%=DisPlay(CheckObjInstalled("ADODB.Stream"))%> <small>(不支持无法实现远程保存和本地创建文件的功能)</small></td>
			</tr>
            <tr>
					<td>Scripting.Dictionary</td>
					<td align="center">On</td>
					<td><%=DisPlay(CheckObjInstalled("Scripting.Dictionary"))%> <small>(不支持无法使用本系统)</small></td>
			</tr>
		</table>
		
		
		<div class="btn-box">
			<input type="button" class="btn-back" value="后退" onclick="window.location.href='step_1.asp';" />
			<input type="button" class="btn-next" value="继续" onclick="window.location.href='step_3.asp';" />
		</div>
	</div>
</div>

<div class="foot">

</div>

</body>
</html>
