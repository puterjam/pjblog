<script Language="JScript" runat="server">
//***************PJblog2 插件附件专用连接数据库脚本*******************
// PJblog2 Copyright 2005
// Update:2005-10-28
//******************************************************************

	function createConnection(strPath){
		try{
			Conn = Server.CreateObject("ADODB.Connection");
			Conn.connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Server.MapPath(strPath);
			Conn.open();
		}catch(e){
		Response.Write (transferID)
		    Response.Write('<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" /><div style="font-size:12px;font-weight:bold;border:1px solid #006;padding:6px;background:#fcc">数据库连接出错，请检查连接字串!</div>');
		    Response.End
		}
	}
</script>
<%
Dim DBPath
DBPath = "../../" & AccessFile
Call createConnection(DBPath)

getInfo(1)

'使用界面
Skins = blog_DefaultSkin
If Len(Skins)<1 Then Skins = "default"

'验证用户登录信息
checkCookies

'读取用户权限
UserRight(1)

'写入标签
Tags(1)

'写入表情符号
Smilies(1)

'写入关键字列表
Keywords(1)

'写入自定义模块缓存
log_module(1)

'禁止IP访问
If MatchIP(getIP) Then
    response.Write "<span style=""font-size:12px;font-weight:bold;margin:4px"">本站不欢迎你的访问！！</span>"
    response.End
End If
%>
