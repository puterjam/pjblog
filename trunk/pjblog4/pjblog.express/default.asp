<!--#include file = "../include.asp" --><!--#include file = "../pjblog.model/cls_ubbcode.asp" --><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link type="text/css" rel="stylesheet" href="../SyntaxHighlighter/Styles/SyntaxHighlighter.css"></link>
<link type="text/css" rel="stylesheet" href="../SyntaxHighlighter/styles/shThemeDefault.css"/>
<style>
	body{ font-size:11px;}
</style>
<title>Untitled Document</title>
<script language="javascript" type="text/javascript" src="../pjblog.common/common.js"></script>
<script language="javascript" type="text/javascript" src="../pjblog.common/Ajax.js"></script>
<script language="javascript" type="text/javascript" src="../pjblog.common/language.js"></script>
<script language="javascript" type="text/javascript" src="../pjblog.common/checkform.js"></script>
	<script type="text/javascript" src="../SyntaxHighlighter/scripts/shCore.js"></script>
	<script type="text/javascript" src="../SyntaxHighlighter/scripts/shBrushBash.js"></script>
	<script type="text/javascript" src="../SyntaxHighlighter/scripts/shBrushCpp.js"></script>
	<script type="text/javascript" src="../SyntaxHighlighter/scripts/shBrushCSharp.js"></script>
	<script type="text/javascript" src="../SyntaxHighlighter/scripts/shBrushCss.js"></script>
	<script type="text/javascript" src="../SyntaxHighlighter/scripts/shBrushDelphi.js"></script>
	<script type="text/javascript" src="../SyntaxHighlighter/scripts/shBrushDiff.js"></script>
	<script type="text/javascript" src="../SyntaxHighlighter/scripts/shBrushGroovy.js"></script>
	<script type="text/javascript" src="../SyntaxHighlighter/scripts/shBrushJava.js"></script>
	<script type="text/javascript" src="../SyntaxHighlighter/scripts/shBrushJScript.js"></script>
	<script type="text/javascript" src="../SyntaxHighlighter/scripts/shBrushPhp.js"></script>
	<script type="text/javascript" src="../SyntaxHighlighter/scripts/shBrushPlain.js"></script>
	<script type="text/javascript" src="../SyntaxHighlighter/scripts/shBrushPython.js"></script>
	<script type="text/javascript" src="../SyntaxHighlighter/scripts/shBrushRuby.js"></script>
	<script type="text/javascript" src="../SyntaxHighlighter/scripts/shBrushScala.js"></script>
	<script type="text/javascript" src="../SyntaxHighlighter/scripts/shBrushSql.js"></script>
	<script type="text/javascript" src="../SyntaxHighlighter/scripts/shBrushVb.js"></script>
	<script type="text/javascript" src="../SyntaxHighlighter/scripts/shBrushXml.js"></script>
	<link type="text/css" rel="stylesheet" href="../SyntaxHighlighter/styles/shCore.css"/>
	<link type="text/css" rel="stylesheet" href="../SyntaxHighlighter/styles/shThemeDefault.css"/>
	<script type="text/javascript">
		SyntaxHighlighter.config.clipboardSwf = '../SyntaxHighlighter/scripts/clipboard.swf';
		SyntaxHighlighter.all();
	</script>
</head>

<body>
<form id="checkUser" action="../pjblog.logic/log_User.asp?action=login" method="post">
    <div id="MsgContent">
      <div id="MsgHead"><%=lang.Set.Asp(17)%></div>
      <div id="MsgBody">
	   <label><input name="username" type="text" size="18" class="userpass" maxlength="24"/> <%=lang.Set.Asp(18)%> </label><br/>
	   <label><input name="password" type="password" size="18" class="userpass" onfocus="this.select()"/> <%=lang.Set.Asp(19)%> </label><br/>
	   <label><input name="validate" type="text" size="4" class="userpass" maxlength="4" id="validate" onfocus="CheckForm.User.GetCode('../pjblog.common/GetCode.asp', 'checkcode')" onkeyup="CheckForm.User.CheckCode('validate', 'isok_checkcode')"/> <span id="checkcode"><label style="cursor:pointer;" onClick="CheckForm.User.GetCode('../pjblog.common/GetCode.asp', 'checkcode')"><%=lang.Set.Asp(4)%></label></span> <span id="isok_checkcode"></span></label><br/>
	   　　<label><input name="KeepLogin" type="checkbox" value="1"/><%=lang.Set.Asp(20)%></label><br/>
	   <input type="submit" value="<%=lang.Set.Asp(21)%>" class="userbutton"/>　<input type="button" value="<%=lang.Set.Asp(22)%>" class="userbutton" onclick="location='register.asp'"/>
	   </div>
	</div>
  </form>
  <a href="blogpost.asp">发表日志</a>
  <span class="chk_edit">dfasf</span>
  <%
  Dim v
  v = UBBCode("[code=asp]function a() when[/code]", 1, 0, 1, 1, 1, True)
  Response.Write(v)
  Response.Write(stat_EditAll)
  %>
  <pre class="brush: vb">
  	function test() : String  

2 {  

3     return 10;  

4 } 

  </pre>

  <pre class="brush: c-sharp;">
function test() : String
{
	return 10;
}
</pre>
<script language="javascript" type="text/javascript" src="../pjblog.logic/code.asp?action=default"></script>
</body>
</html>

<%
Sys.Close
%>
