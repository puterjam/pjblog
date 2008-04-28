<%
	Response.Charset = "UTF-8"
	Response.ContentType = "text/vnd.wap.wml" 
	Response.Expires = -1
	Response.AddHeader "Pragma", "no-cache"
	Response.AddHeader "Cache-Control", "no-cache, must-revalidate"
	
	function toUnicode(str) 'To Unicode
	    dim i,a
		For i = 1 to Len (str)
			a=Mid(str, i, 1)
			toUnicode=toUnicode & "&#x" & Hex(Ascw(a)) & ";"
		next
	end function
%>
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
	<head><meta forua="true" http-equiv="Cache-Control" content="max-age=0" /></head>
	<card id="MainCard" title="&#x6B22;&#x8FCE;&#x5149;&#x4E34;" ontimer="../wap.asp">
		<timer value="1"/>
		<p><%=toUnicode("页面加载中... 自动跳转到首页")%></p>
		<p>[<a href="../wap.asp"><%=toUnicode("点击进入首页")%></a>]</p>
	</card>
</wml>


