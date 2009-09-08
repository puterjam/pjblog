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
%>