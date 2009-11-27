<%
'*************** PJblog4 基本设置 *******************
' PJblog4 Copyright 2009
' Update: 2009-11-03
' Author: evio
' Owner : PuterJam
' Mode  : Class
' MoName: User Model
'***************************************************
Dim Control : Set Control = New Sys_Control

Class Sys_Control

	' ***********************************************
	'	类初始化
	' ***********************************************
	Private Sub Class_Initialize
    End Sub 
     
	' ***********************************************
	'	类终结化
	' ***********************************************
    Private Sub Class_Terminate
    End Sub
	
	'----------- 显示操作信息 ----------------------------
	Public Sub getMsg
		If Session(Sys.CookieName & "_ShowMsg") = True Then
			Response.Write ("<script>$(document).ready(function(){effect.WarnTip.open({title : '反馈信息', html : '" & Session(Sys.CookieName & "_MsgText") & "'})});</script>")
			Session(Sys.CookieName & "_ShowMsg") = False
			Session(Sys.CookieName & "_ShowMsg") = ""
		End If
	End Sub
	
	' ***********************************************
	'	后台页面标题选择
	' ***********************************************
	Public Function categoryTitle
		Dim Fmenu, Smenu, cTitle
		Fmenu = Request.QueryString("Fmenu")
		Smenu = Request.QueryString("Smenu")
		Set cTitle = Server.CreateObject("Scripting.Dictionary")
		cTitle.Add "General.clear" , "站点基本设置 - 清理服务器缓存"
		cTitle.Add "General.Misc" , "站点基本设置 - 初始化数据"
		cTitle.Add "General.visitors" , "站点基本设置 - 查看访客记录"
		cTitle.Add "General.UpLoadSet" , "站点基本设置 - 附件基本设置"
		cTitle.Add "General.Ping" , "站点基本设置 - Ping服务设置"
		cTitle.Add "General." , "站点基本设置 - 设置基本信息"
	
		cTitle.Add "Categories.move" , "日志分类管理 - 批量移动日志"
		cTitle.Add "Categories.tag" , "日志分类管理 - Tag管理"
		cTitle.Add "Categories.del" , "日志分类管理 - 批量删除日志"
		cTitle.Add "Categories." , "日志分类管理 - 设置日志分类"
	
		cTitle.Add "Comment.spam" , "评论留言管理 - 初级过滤设置 (垃圾关键字过滤黑名单)"
		cTitle.Add "Comment.reg" , "评论留言管理 - 高级过滤设置 (利用正则表达式过滤)"
		cTitle.Add "Comment.trackback" , "评论留言管理 - 引用管理"
		cTitle.Add "Comment.msg" , "评论留言管理 - 留言管理"
		cTitle.Add "Comment." , "评论留言管理 - 评论管理"
	
		cTitle.Add "Skins.module" , "界面设置 - 设置模块"
		cTitle.Add "Skins.Plugins" , "界面设置 - 已装插件管理"
		cTitle.Add "Skins.PluginsInstall" , "界面设置 - 安装插件"
		cTitle.Add "Skins.PluginsOnline" , "界面设置 - 在线安装插件"
		cTitle.Add "Skins.editModule" , "界面设置 - 可视化编辑HTML代码"
		cTitle.Add "Skins.editModuleNormal" , "界面设置 - 编辑HTML源代码"
		cTitle.Add "Skins.PluginsOptions" , "界面设置 - 插件配置"
		cTitle.Add "Skins.BaseSetting" , "界面设置 - 插件基本设置"
		cTitle.Add "Skins." , "界面设置 - 设置外观"
	
		cTitle.Add "SQLFile.Attachment" , "数据库与附件 - 附件信息"
		cTitle.Add "SQLFile.Attachments" , "数据库与附件 - 附件管理"
		cTitle.Add "SQLFile." , "数据库与附件 - 数据库管理"
	
		cTitle.Add "Members.Users" , "帐户与权限管理 - 帐户管理"
		cTitle.Add "Members.EditRight" , "帐户与权限管理 - 编辑权限细节"
		cTitle.Add "Members." , "帐户与权限管理 - 权限管理"
	
		cTitle.Add "Link." , "友情链接管理 - 链接列表"
		cTitle.Add "Link.LinkClass" , "友情链接管理 - 分类管理"
	
		cTitle.Add "smilies.KeyWord" , "表情与关键字 - 关键字管理"
		cTitle.Add "smilies." , "表情与关键字 - 表情管理"
	
		cTitle.Add "Status." , "服务器配置信息 - PJBlog4"
	
		cTitle.Add "welcome." , "欢迎使用PJBlog4"
		
		cTitle.Add "CodeEditor." , "在线代码编辑器 beta"
		
		categoryTitle = cTitle(Fmenu & "." & Smenu)
		Set cTitle = Nothing
	End Function
	
	Public Function FreeApplicationMemory
		On Error Resume Next
		Response.Write "释放网站缓存数据列表：<div style='padding:5px 5px 5px 10px;'>"
		Dim Thing,i
		i=0
		For Each Thing IN Application.Contents
			
			If Left(Thing, Len(CookieName)) = CookieName Then
				i=i+1
				if i<30 then
					Response.Write "<span style='color:#666'>" & thing & "</span><br/>"
				elseif i<31 then
					Response.Write "<span style='color:#666'>...</span><br/>"
				end if
				
				If IsObject(Application.Contents(Thing)) Then
					Application.Contents(Thing).Close
					Set Application.Contents(Thing) = Nothing
					Application.Contents.Remove(Thing)
				ElseIf IsArray(Application.Contents(Thing)) Then
					Set Application.Contents(Thing) = Nothing
					Application.Contents.Remove(Thing)
				Else
					Application.Contents.Remove(Thing)
				End If
			End If
			
		Next
		Response.Write "<br/><span style='color:#666'>共清理了 " & i & " 个缓存数据</span><br/>"
		Response.Write "</div>"
	End Function
End Class
%>