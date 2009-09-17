<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="style.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="install.js"></script>
<script language="javascript" src="Ajax.js"></script>
<title>PJBlog 安装程序</title>
</head>
<script language="jscript" runat="server">
	if (!Request.Form){Response.Redirect("step_3.asp")}
</script>
<script language="javascript">
function scollDiv(){
	$("Ajax").scrollTop = parseInt($("Ajax").scrollHeight);
	window.reload = false;
	FollowBox("Ajax", "processBar");
}
</script>
<body onload="scollDiv()">
<!--#include file = "misc.asp" -->
<!--#include file = "pack.asp" -->
<!--#include file = "Stream.asp" -->
<%
	If Len(Session("con_db")) > 0 Then Response.Redirect("step_3.asp")
	If Request.Cookies("insName") <> "new" And Request.Cookies("insName") <> "update" Then Response.Redirect("step_3.asp")
	Dim InstallType, InstallCookie, InstallAccess, InsStream, ConstStr, BaseStr
	InstallType = Request.Form("installType")
	If InstallType = "new" Then
		If Len(Request.Form("dbfolder")) = 0 Or Len(Request.Form("dbname")) = 0 Then Response.Redirect("step_3.asp")
		InstallAccess = Request.Form("dbfolder") & "/" & Request.Form("dbname")
		InstallCookie = Request.Form("cookieName")
		BaseStr = "{adminuser : '" & Request.Form("adminuser") & "', adminpwd : '" & Request.Form("adminpwd") & "', webname : '" & Request.Form("webname") & "', adminmail : '" & Request.Form("adminmail") & "', baseurl : '" & Request.Form("baseurl") & "', author : '" & Request.Form("author") & "'}"
	ElseIf InstallType = "update" Then
		BaseStr = ""
		Set InsStream = New Stream
			ConstStr = InsStream.LoadFile("../const.asp")
			Dim MatchCookie, MatchAccess, outCookie, outAccess
			Set MatchCookie = GetMatch(ConstStr, "Const(\s+)CookieName(\s+)\=(\s+)\""(.*)\""")
			If MatchCookie.count > 0 Then
				outCookie = MatchCookie(0).SubMatches(3)
			End If
			Set MatchCookie = Nothing
			Set MatchAccess = GetMatch(ConstStr, "Const(\s+)AccessFile(\s+)\=(\s+)\""(.*)\""")
			If MatchAccess.count > 0 Then
				outAccess = MatchAccess(0).SubMatches(3)
			End If
			Set MatchAccess = Nothing
		Set InsStream = Nothing
		InstallCookie = outCookie
		InstallAccess = outAccess
		If Len(InstallCookie) = 0 Or Len(InstallAccess) = 0 Then Response.Redirect("step_3.asp")
	Else
		Response.Redirect("step_3.asp")
	End If
	Session("con_db") = BaseStr
	Response.Cookies("InstallType") = InstallType
	Response.Cookies("InstallType").Expires = DateAdd("d", 1, now())
	Response.Cookies("InstallCookie") = InstallCookie
	Response.Cookies("InstallCookie").Expires = DateAdd("d", 1, now())
	Response.Cookies("InstallAccess") = InstallAccess
	Response.Cookies("InstallAccess").Expires = DateAdd("d", 1, now())
%>
<div class="main">
	<div class="pleft">
		<dl class="setpbox t1">
			<dt>安装步骤</dt>
			<dd>
				<ul>
					<li class="succeed">许可协议</li>
					<li class="succeed">环境检测</li>
					<li class="succeed">参数配置</li>
					<li class="now">安装模块</li>
					<li>安装完成</li>                  
				</ul>
			</dd>
		</dl>
	</div>
	<div class="pright">
		<div class="pr-title"><h3>正在安装模块</h3></div>
		<div class="install-msg">
			<!--<iframe src="index.php?step=5" id="mainfra" name="mainfra" frameborder="0" width='100%' height='350px'></iframe>-->
            <div class=""Area"">
            <div id="Ajax">
            <%
			Dim insP
			Set insP = new PJSetup
			If InstallType = "new" Then
				insP.xmlPath = "package.pbd"
			ElseIf InstallType = "update" Then
				insP.xmlPath = "update.pbd"
			End If
			insP.open
			insP.install
			Set insP = Nothing
			If InstallType = "new" Then
				Set InsStream = New Stream
					ConstStr = InsStream.LoadFile("const.asp")
					ConstStr = Replace(ConstStr, "<@CookieName#>", InstallCookie)
					ConstStr = Replace(ConstStr, "<@AccessFile#>", InstallAccess)
					InsStream.SaveToFile ConstStr, "../const.asp"
				Set InsStream = Nothing
			End If
			%>
            </div>
            	<div id="processBar"><div id="process"></div></div>
             	<p class="post"><a href="javascript:;" onclick="this.disabled=true;install.Start('Ajax');this.onclick=function(){return false;}" id="button">开始在线创建数据库或者升级数据库</a><sup id="number" style="color:#f00; font-size:10px;"></sup></p>
             </div>
		</div>
	</div>
</div>

<div class="foot">
</div>

</body>
</html>

</body>
</html>