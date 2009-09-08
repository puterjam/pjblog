<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% Session.CodePage = 65001 %>
<% Response.Charset="UTF-8" %>
<!--#include file="misc.asp" -->
<!--#include file="xml.asp" -->
<!--#include file="Stream.asp" -->
<!--#include file="fso.asp" -->
<!--#include file="Pack.asp" -->
<script language="jscript" type="text/jscript" runat="server">
/******************************************************/
// 请修改此处的数据库地址
	var MsPath = "blogDB/PBLog3.asp";
	try{MsPath = Request.Cookies("InstallAccess");}catch(e){}
	//var MsPath = "blogDB/PBLog3.asp";
/******************************************************/
	var step = Request.Form("step");
	var info = unescape(Request.Form("info"));
	var check = false, str, Conn;
	
	//检测组件
	if (step == "1" || step == 1){
		check = CheckObjInstalled(info);
		if (check){
			str = "{suc : true, info : '" + transHtml("<font color=green>检测组件 " + info + " 已安装...</font>") + "'}";
		}else{str = "{suc : false, info : '" + transHtml("<font color='red'>检测组件 " + info + " 未安装...</font>") + "'}"};
		Response.Write(str)
	}
	
	//创建数据库
	if (step == "2" || step == 2){
		if (MsPath.length == 0){
			Response.Write("{suc : false, info : '数据库地址不能为空!'}")
		}else{
			var CmsPath = "../" + MsPath;
			if (createObj(CmsPath)){
				str = "{suc : true, info : '" + transHtml("<font color=green>创建数据库 [" + info + "] 成功...</font>") + "'}";
			}else{str = "{suc : false, info : '" + transHtml("<font color='#cccccc'>创建数据库 [" + info + "] 失败...</font>") + "'}"}
			Response.Write(str)
		}
	}
	
	//连接数据库
	if (step == "3" || step == 3){
		if (MsPath.length == 0){
			Response.Write("{suc : false, info : '数据库地址不能为空!'}")
		}else{
			var CmsPath = "../" + MsPath;
			if (createConnection(CmsPath)){
				var Re = UpdateSQL(info);
				if (Re[0]){
					str = "{suc : true, info : '" + transHtml("<font color=green>" + Re[1] + "</font>") + "'}";
				}else{
					str = "{suc : false, info : '" + transHtml("<font color='#cccccc'>" + Re[1] + "</font>") + "'}"
				}
			}else{str = "{suc : false, info : '" + transHtml("<font color='#cccccc'>连接数据库 [" + MsPath + "] 失败...</font>") + "'}"}
			Response.Write(str)
		}
	}
	
	//清除全局缓存
	if (step == "4" || step == 4){
		try{
			Application.Lock();
			Application(info) = "";
			Application.UnLock();
			str = "{suc : true, info : '" + transHtml("<font color=green>清除全局缓存 [" + info + "] 成功...</font>") + "'}";
		}catch(e){str = "{suc : false, info : '" + transHtml("<font color='#cccccc'>清除全局缓存 [" + info + "] 失败...</font>") + "'}"}
		Response.Write(str)
	}
	
</script>
<%
	If step = "11" Or step = 11 Then
		Dim Install_Cookie, Install_Access
		Install_Cookie = Request.Form("Install_Cookie")
		Install_Access = Request.Form("Install_Access")
		Response.Cookies("InstallCookie") = Install_Cookie
		Response.Cookies("InstallCookie").Expires = DateAdd("d", 1, now())
		Response.Cookies("InstallAccess") = Install_Access
		Response.Cookies("InstallAccess").Expires = DateAdd("d", 1, now())
		Response.Redirect("Right.asp?misc=file")
	End If
	
	If step = "0" Or step = 0 Then
		On Error Resume Next
		Dim id, appendStr, PJclass, cFn, Cfso, SplitFile, cFolder, cStream, Scontent, baseM, Str
		Dim ElementName, ElementContent
		id = Trim(Request.Form("id"))
		appendStr = "../"
		Set PJclass = New PJSetup
			PJclass.xmlPath = "package.pbd"
			PJclass.open
			ElementName = PJclass.element_fn(id)
			ElementContent = PJclass.element_fb(id)
		Set PJclass = Nothing
		Set baseM = New base64
			If Len(ElementContent) = 0 Then
				Scontent = ""
			Else
				Scontent = baseM.decode(ElementContent)
			End If
		Set baseM = Nothing
		cFn = Replace(ElementName, "\", "/")
		cFn = appendStr & cFn
		Set Cfso = New cls_FSO
			If Instr(cFn, "/") <> 0 Then
				SplitFile = Split(cFn, "/")
				cFolder = Replace(cFn, SplitFile(UBound(SplitFile)), "")
				Cfso.CreateLoopFolder cFolder
			End If
		Set Cfso = Nothing
		Set cStream = Server.CreateObject("adodb.stream")
			If Err <> 0 Then Err.clear
			cStream.Type = 1
			cStream.Mode = 3
			cStream.Open
			cStream.Position = 0
			cStream.write Scontent
			cStream.SaveToFile Server.MapPath(cFn), 2
			If Err <> 0 Then
				Str = "{suc : false, info : '" & transHtml("<font color=red>" & id & ".文件名" & ElementName & "," & Err.Description & "</font>") & "'}"
			Else
				Str = "{suc : true,info : '<font color=green>" & id & "." & transHtml(ElementName & " 保存成功...") & "</font>'}"	
			End If
		Set cStream = Nothing
		Response.Write(Str)		
	End If
	
	If step = "10" Or step = 10 Then
		Dim DelInstallFiles
		DelInstallFiles = Request.Form("DelInstallFiles")
		If DelInstallFiles = 1 Or DelInstallFiles = "1" Then
			Dim Jfso
			Set Jfso = New cls_FSO
				With Jfso
					.DeleteFile("../install.html")
					.DeleteFolder "../install/", True
				End With
			Set Jfso = Nothing
		End If
		Response.Cookies("install_step") = 0
		Response.Write("<a href=""../login.asp"" style=""font-size:11px;"" target=""_blank"">临时文件清理成功, 请登入博客</a>")
	End If
%>