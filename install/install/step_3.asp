<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>安装程序 - 织梦内容管理系统</title>
<script src="install.js" language="javascript" type="text/javascript"></script>
<script language="javascript" type="text/javascript">
	var new_open = {id : ["open_1", "open_2", "open_3"], type : "open"};
	var update_open = {id : ["open_1", "open_2", "open_3"], type : "close"};
</script>
<link href="style.css" rel="stylesheet" type="text/css" />
</head>
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
<body>
<form action="step_4.asp" method="post" onsubmit="return doCheck(this)">
<div class="main">
	<div class="pleft">
		<dl class="setpbox t1">
			<dt>安装步骤</dt>
			<dd>
				<ul>
					<li class="succeed">许可协议</li>
					<li class="succeed">环境检测</li>
					<li class="now">参数配置</li>
					<li>正在安装</li>
					<li>安装完成</li>
				</ul>
			</dd>
		</dl>
	</div>
	<div class="pright">
		<div class="pr-title"><h3>安装类型选择</h3></div>
		<table width="726" border="0" align="center" cellpadding="0" cellspacing="0" class="twbox">
			<tr>
				<td style='line-height:180%'>
					<b>选择安装类型(全新安装或者程序升级)：</b><br />
					<label for="T1"><input type="radio" value="new" id="T1" name="installType" onclick="Idisplay(new_open)" checked="checked" /><em>全新安装(您将安装全新的PJBlog程序.)</em></label><br />
                    <label for="T2"><input type="radio" value="update" id="T2" name="installType" onclick="Idisplay(update_open)" /><em>程序升级(您将升级到300版本.)</em></label>
				</td>
			</tr>
		</table>
        
      <div id="open_1" style="display:block">	
		<div class="pr-title"><h3>数据库相关设定</h3></div>
		<table width="726" border="0" align="center" cellpadding="0" cellspacing="0" class="twbox">
			<tr>
				<td class="onetd"><strong>数据库文件夹：</strong></td>
				<td><input name="dbfolder" type="text" value="BlogDB" class="input-txt" onblur="ReplaceInput(this,window.event)" onkeyup="ReplaceInput(this,window.event)" />
				<small>数据库所在文件夹</small></td>
			</tr>
			<tr>
				<td class="onetd"><strong>数据库名称：</strong></td>
				<td>
					<input name="dbname" type="text" value="PBlog3.asp" class="input-txt" />
				</td>
			</tr>
		</table>
      </div>
		
      <div id="open_2" style="display:block">
		<div class="pr-title"><h3>管理员初始密码</h3></div>
		<table width="726" border="0" align="center" cellpadding="0" cellspacing="0" class="twbox">
			<tr>
				<td class="onetd"><strong>用户名：</strong></td>
				<td>
					<input name="adminuser" type="text" value="admin" class="input-txt" />
					<p><small>只能用'0-9'、'a-z'、'A-Z'、'.'、'@'、'_'、'-'、'!'以内范围的字符</small></p>
				</td>
			</tr>
			<tr>
				<td class="onetd"><strong>密　码：</strong></td>
				<td><input name="adminpwd" type="password" value="" class="input-txt" /></td>
			</tr>
			<tr>
				<td class="onetd"><strong>重复密码：</strong></td>
				<td><input name="cookieencode" type="password" value="" class="input-txt" /></td>
			</tr>
		</table>
      </div>

	  <div id="open_3" style="display:block">
		<div class="pr-title"><h3>网站设置</h3></div>
		<table width="726" border="0" align="center" cellpadding="0" cellspacing="0" class="twbox">
			<tr>
				<td class="onetd"><strong>网站名称：</strong></td>
				<td>
					<input name="webname" type="text" value="我的网站" class="input-txt" />
				</td>
			</tr>
			<tr>
				<td class="onetd"><strong>管理员邮箱：</strong></td>
				<td><input name="adminmail" type="text" value="admin@website.com" class="input-txt" /></td>
			</tr>
			<tr>
				<td class="onetd"><strong> 网站网址：</strong></td>
				<td><input name="baseurl" type="text" value="" class="input-txt" id="c_webUrl" /><small>很重要的地址.</small></td>
			</tr>
			<tr>
				<td class="onetd"><strong> 博主昵称：</strong></td>
				<td><input name="author" type="text" class="input-txt" value="PuterJam" /></td>
			</tr>
            <tr>
				<td class="onetd"><strong> 设置Cookie名称：</strong></td>
				<td><input name="cookieName" type="text" class="input-txt" value="PJBlog3" /></td>
			</tr>
		</table>
      </div>
	<script language="javascript" type="text/javascript">
		document.getElementById("c_webUrl").value = location.href.replace(/(.*)\/step_3.asp/,"$1\/");
	</script>
		<div class="btn-box">
			<input type="button" class="btn-back" value="后退" onclick="window.location.href='step_2.asp';" />
			<input type="submit" class="btn-next" value="开始安装" />
		</div>
	</div>
</div>
<div class="foot">
</div>
</form>
</body>
</html>
