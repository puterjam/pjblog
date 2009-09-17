<script language="jscript" type="text/jscript" runat="server">
/*
	作者 evio
	日期 2009-09-05
*/

//*************************************
//	检测系统组件是否安装
//*************************************
	function CheckObjInstalled(strClassString){
		try{
			var TmpObj = Server.CreateObject(strClassString);
			return true;
		}catch(e){
			return false;
		}
	}

//*************************************
//	安装与否
//*************************************
	function DisPlay(str){
		var Bool;
		if (str){
			Bool = "<font color='green'>√</font>";
		}else{
			Bool = "<font color='red'>×</font>";
		}
		return Bool;
	}
	
//*************************************
//	创建数据库的方法
//*************************************
	function createObj(strPath){
		try{
			Cate = Server.CreateObject("ADOX.Catalog");
			Cate.create("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Server.MapPath(strPath));
			return true;
		}catch(e){return false;}
	}
	
//*************************************
//	创建数据库连接的方法
//*************************************
	function createConnection(strPath){
		try{
			Conn = Server.CreateObject("ADODB.Connection");
			Conn.connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Server.MapPath(strPath);
			Conn.open();
			return true;
		}catch(e){
		    return false;
		}
	}
	
//*************************************
//	操作数据库方法
//*************************************
	function UpdateSQL(str){
		try{
			Conn.Execute(str);
			return Array(true, ("执行操作 " + str + " 成功..."));
		}catch(e){
			if (e.description.length > 0){
				return Array(false, (e.description));
			}
		}
	}
	
//*************************************
//	转换特殊字符
//*************************************
function transHtml(html){
	html = html.replace(/[\n\r]/g,"");
	html = html.replace(/\\/g,"\\\\");
	html = html.replace(/\'/g,"\\'");
	return html;
}

function ticheck(str){
	str = str.replace(/\</g, "&lt;");
	str = str.replace(/\>/g, "&gt;");
	str = str.replace(/\"/g, "&#34;");
	str = str.replace(/\'/g, "&#39;");
	str = str.replace(/\&/g, "&amp;");
	return str
}

//*************************************
// Json格式ASP化
//*************************************
	function toJson(Str){
		try{
			eval("var jsonStr = (" + Str + ")");
		}catch(ex){
			var jsonStr = null;
		}
		return jsonStr;
	}
	
	function Escape(str){return escape(str);}
	function Unescape(str){return unescape(str);}
	
</script>

<%
	Public Function installBox(ByVal i)
		installBox = ""
		If Int(i) = 1 Then
			Dim xmlFileCount, insP
			Set insP = New PJSetup
				insP.xmlPath = "package.pbd"
				insP.open
				xmlFileCount = (Int(insP.fileTotal) - 1)
			Set insP = Nothing
		End If
		installBox = installBox & "<script language=""javascript"" type=""text/javascript"">" & vbcrlf
		installBox = installBox & "window.onload = function(){window.reload = false;FollowBox(""Ajax"", ""processBar"")}" & vbcrlf
		installBox = installBox & "document.onkeydown=function(){" & vbcrlf
		installBox = installBox & "with(window.event){" & vbcrlf
		installBox = installBox & "if(keyCode == 116){" & vbcrlf
		installBox = installBox & "keyCode = 0;" & vbcrlf
		installBox = installBox & "cancelBubble = true;" & vbcrlf
		installBox = installBox & "return false;" & vbcrlf
		installBox = installBox & "}" & vbcrlf
		installBox = installBox & "}" & vbcrlf
		installBox = installBox & "}" & vbcrlf
		If Int(i) = 1 Then installBox = installBox & "var xml = " & xmlFileCount & ";"
		installBox = installBox & "</script>" & vbcrlf
		installBox = installBox & "<div class=""Area"">" & vbcrlf
        installBox = installBox & "<div id=""Ajax""></div>" & vbcrlf
        installBox = installBox & "<div id=""processBar""><div id=""process""></div></div>" & vbcrlf
		If Int(i) = 1 Then
			installBox = installBox & "<p class=""post""><a href=""javascript:;"" onclick=""this.disabled=true;File.Start();this.onclick=function(){return false;}"" id=""button"">开始解压程序文件</a><sup id=""number"" style=""color:#f00; font-size:10px;""></sup></p>" & vbcrlf
		Else
			installBox = installBox & "<p class=""post""><a href=""javascript:;"" onclick=""this.disabled=true;install.Start('Ajax');this.onclick=function(){return false;}"" id=""button"">开始在线创建数据库或者升级数据库</a><sup id=""number"" style=""color:#f00; font-size:10px;""></sup></p>" & vbcrlf
		End If
        installBox = installBox & "<p class=""post"" style="" text-align:left""></p>" & vbcrlf
        installBox = installBox & "</div>" & vbcrlf
	End Function
	
	'*************************************
	'	CheckStr
	'*************************************
	Public Function CheckStr(byVal ChkStr)
		Dim Str
		Str = ChkStr
		If IsNull(Str) Then
			CheckStr = ""
			Exit Function
		End If
		Str = Replace(Str, "&", "&amp;")
		Str = Replace(Str, "'", "&#39;")
		Str = Replace(Str, """", "&#34;")
		Str = Replace(Str, "<", "&lt;")
		Str = Replace(Str, ">", "&gt;")
		Dim re
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = "(w)(here)"
		Str = re.Replace(Str, "$1h&#101;re")
		re.Pattern = "(s)(elect)"
		Str = re.Replace(Str, "$1el&#101;ct")
		re.Pattern = "(i)(nsert)"
		Str = re.Replace(Str, "$1ns&#101;rt")
		re.Pattern = "(c)(reate)"
		Str = re.Replace(Str, "$1r&#101;ate")
		re.Pattern = "(d)(rop)"
		Str = re.Replace(Str, "$1ro&#112;")
		re.Pattern = "(a)(lter)"
		Str = re.Replace(Str, "$1lt&#101;r")
		re.Pattern = "(d)(elete)"
		Str = re.Replace(Str, "$1el&#101;te")
		re.Pattern = "(u)(pdate)"
		Str = re.Replace(Str, "$1p&#100;ate")
		re.Pattern = "(\s)(or)"
		Str = re.Replace(Str, "$1o&#114;")
		re.Pattern = "(java)(script)"
		Str = re.Replace(Str, "$1scri&#112;t")
		re.Pattern = "(j)(script)"
		Str = re.Replace(Str, "$1scri&#112;t")
		re.Pattern = "(vb)(script)"
		Str = re.Replace(Str, "$1scri&#112;t")
		If Instr(Str, "expression") <> 0 Then
			Str = Replace(Str, "expression", "e&#173;xpression", 1, -1, 0) '防止xss注入
		End If
		Set re = Nothing
		CheckStr = Str
	End Function
	
	Public Function GetMatch(ByVal Str, ByVal Rex)
		Dim Reg, Mag
		Set Reg = New RegExp
		With Reg
			.IgnoreCase = True
			.Global = True
			.Pattern = Rex
			Set Mag = .Execute(str)
			If Mag.Count > 0 Then
				Set GetMatch = Mag
			Else
				Set GetMatch = Server.CreateObject("Scripting.Dictionary")
			End If
		End With
		Set Reg = nothing
	End Function
%>