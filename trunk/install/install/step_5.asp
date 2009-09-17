<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="style.css" rel="stylesheet" type="text/css" />
<title>无标题文档</title>
<script language="javascript">
function goUrl(str){
	document.getElementById("delForm").gourl.value = str;
	document.getElementById("delForm").submit();
}
</script>
<!--#include file = "misc.asp" -->
<!--#include file = "md5.asp" -->
<!--#include file = "sha1.asp" -->
</head>
<%
Response.Cookies("insName") = ""
Response.Cookies("insName").Expires = DateAdd("d", 1, now())
If Len(Request.Cookies("InstallType")) = 0 Then Response.Redirect("step_1.asp")
If Lcase(Request.Cookies("InstallType")) = "new" Then
	If Len(Session("con_db")) > 0 Then
		Dim con_Get, con_Connect, con_True, con_SQL
		Dim adminuser, adminpwd, webname, adminmail, baseurl, author
		Set con_Get = toJson(Session("con_db"))
			adminuser = con_Get.adminuser
			adminpwd = con_Get.adminpwd
			webname = con_Get.webname
			adminmail = con_Get.adminmail
			baseurl = con_Get.baseurl
			author = con_Get.author
		Set con_Get = Nothing
'		Response.Write(Session("con_db") & "<br />")
'		Response.Write(adminuser & "|" & adminpwd & "|" & webname & "|" & adminmail & "|" & baseurl & "|" & author)
		con_Connect = "../" & Request.Cookies("InstallAccess")
		If createConnection(con_Connect) Then
			con_SQL = "Update blog_Member Set mem_Name='" & adminuser & "', mem_Password='" & SHA1(adminpwd & "123456") & "', mem_salt='123456' Where mem_ID=1"
			con_True = UpdateSQL(con_SQL)
			con_SQL = "Update blog_Info Set blog_Name='" & webname & "', blog_URL='" & baseurl & "', blog_email='" & adminmail & "', blog_master='" & author & "'"
			con_True = UpdateSQL(con_SQL)
		End If
	End If
End If
%>
<body>
<div class="main">
	<div class="pleft">
		<dl class="setpbox t1">
			<dt>安装步骤</dt>
			<dd>
				<ul>
					<li class="succeed">许可协议</li>
					<li class="succeed">环境检测</li>
					<li class="succeed">参数配置</li>
					<li class="succeed">正在安装</li>
					<li class="now">安装完成</li>
				</ul>
			</dd>
		</dl>
	</div>
    <form action="Action.asp" method="post" id="delForm">
	<div class="pright">
		<div class="pr-title"><h3>安装完成</h3></div>
		<div class="install-msg">
        
			恭喜您!已成功安装/升级了PJBlog系统.您现在可以:
			<br />
            <p>
            	
                	<input type="hidden" name="step" value="10" />
                    <input type="hidden" name="gourl" value="" />
                	<ul style="margin-top:20px;">
                    	<li style="line-height:22px;"><label for="T1"><input type="checkbox" value="1" name="DelInstallFiles" id="T1" checked="checked" />&nbsp;删除安装文件(如果不删除, 请手动删除网站目录下的install.html 和 install文件夹!)</label></li>
                        <li style="line-height:22px;">至此,所有的安装过程都已经结束.请登入后台设置你的分类或者进行必要的操作!</li>
                    </ul>
                
            </p>
		</div>
		<div class="over-link fs-14">
			<a href="javascript:;" onclick="goUrl('../')">访问网站首页</a>
			<a href="javascript:;" onclick="goUrl('../control.asp')">登陆网站后台</a>
		</div>
		<div class="install-msg">
			或者访问织梦网站:<br />
		</div>
		<div class="over-link">
			<a href="javascript:;" onclick="goUrl('http://www.pjhome.net')">PJBlog 官方网站</a>
			<a href="javascript:;" onclick="goUrl('http://bbs.pjhome.net')">PJBlog 官方论坛</a>
		</div>
	</div>
    </form>
</div>

<div class="foot">
</div>

</body>
</html>