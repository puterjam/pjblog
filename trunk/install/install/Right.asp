<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
	Response.Charset = "UTF-8"
	Session.CodePage = 65001
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" rev="stylesheet" href="Style.Css" type="text/css" media="all" />
<script type="text/javascript" language="javascript" src="Ajax.js"></script>
<script type="text/javascript" language="javascript" src="install.js"></script>
<title>PJBlog  安装升级程序</title>
</head>
<!--#include file = "Pack.asp" -->
<!--#include file = "misc.asp" -->
<body class="ehome" marginwidth="0" marginheight="0">
	<div id="main">
    	<div class="title">PJBlog 安装升级程序</div>
       	<%
			If Request.QueryString("misc") = "install" Then
				Response.Write(installBox(0))
			ElseIf Request.QueryString("misc") = "file" Then
				Response.Write(installBox(1))
			ElseIf Request.QueryString("misc") = "complate" Then
		%>
        	<p>
            	<form action="Action.asp" method="post">
                	<input type="hidden" name="step" value="10" />
                	<ul style="margin-top:20px;">
                    	<li style="line-height:22px;"><label for="T1"><input type="checkbox" value="1" name="DelInstallFiles" id="T1" checked="checked" />&nbsp;删除安装文件(如果不删除, 请手动删除网站目录下的install.html 和 install文件夹!)</label></li>
                        <li style="line-height:22px;">至此,所有的安装过程都已经结束.请登入后台设置你的分类或者进行必要的操作!</li>
                    </ul>
                    <p style=" margin-top:10px; margin-left:10px;"><input type="submit" value="完成安装" class="Textbutton" /></p>
                </form>
            </p>
        <%
			Else
				Response.Cookies("InstallCookie") = ""
				Response.Cookies("InstallCookie").Expires = DateAdd("d", 1, now())
				Response.Cookies("InstallAccess") = ""
				Response.Cookies("InstallAccess").Expires = DateAdd("d", 1, now())
		%>
        <p class="p2em" style="font-size:14px;"></p>
        	  <h4> <em>★ 程序安装前必须注意的事项</em></h4>
              <ul style="margin-top:10px;">
              	<li>1. 打开 <font color="red">install</font> 文件夹, 找到 <font color="red">install.js</font>, 在第二行处修改 <font color="red">cookie</font> 和 <font color="red">AccessPath</font> 的值, 改成你自己的设计, 即和你博客上的设置相同.(你博客上此处设置在网站根目录的 <font color="red">const.asp</font> 文件中, 你可以找到 <font color="red">CookieName</font> 和 <font color="red">AccessFile</font>, 如果你还是找不到或者不会修改, 请参看本程序详细的图文修改教程!)</li>
              	<li>2. 打开 <font color="red">install</font> 文件夹, 找到 <font color="red">Action.asp</font>, 在 12 行找到 <font color="red">var MsPath = "blogDB/PBLog3.asp";</font> 修改此处的数据库地址.你可以参照以上的修改来修改!</li>
                <li>如果出现<font color="blue"><em>"参数类型不正确, 或者不可接受的范围之内, 或与其他参数冲突"</em></font>的错误,请别慌张,属于正确的现象.</li>
              </ul>
			  <h4 style=" margin-top:10px;"> <em>★ PJBlog v 3.2.7.300 更新的内容</em></h4>
              <ul style="margin-top:10px;">
              	<li>1. dasf</li>
              	<li>2. adsf</li>
              </ul>
              
              <form action="Action.asp?step=11" method="post" id="PostCookie">
              <input type="hidden" name="step" value="11" />
              <h4 style=" margin-top:10px;"> <em>★ 安装前基本设置 (设置你的Cookie和AccessPath)</em></h4>
              <ul style="margin-top:10px;">
              	<li>CookieName <input type="text" value="" name="Install_Cookie"></li>
              	<li>AccessPath <input type="text" value="blogDB/PBLog3.asp" name="Install_Access"></li>
              </ul>
              </form>
        <p style="text-align:right; padding:3px 50px 3px 50px;">evio<br />2009-08-29</p>
         <p style="text-align:right; padding:3px 50px 3px 50px;"><input type="checkbox" value="1" onclick="install.Agree(this)">我同意&nbsp;&nbsp;&nbsp;<a href="?misc=file" class="Textbuttons" disabled="disabled" id="next" onclick="return false">下一步</a></p>
        <%
			End If
		%>
    </div>
</body>
</html>
