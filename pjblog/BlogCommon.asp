<!--#include file="const.asp" -->
<!--#include file="conn.asp" -->
<%
'==================================
'  Blog参数调用页面
'    更新时间: 2005-10-28
'==================================
'读取Blog设置信息
getInfo(1)

'使用界面
Skins = blog_DefaultSkin
'客户端自选界面Cookie
If Len(Request.Cookies(CookieNameSetting)("BlogSkin"))>0 Then Skins = Request.Cookies(CookieNameSetting)("BlogSkin")
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
