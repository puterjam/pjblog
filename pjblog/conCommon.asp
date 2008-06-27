<!--#include file="const.asp" -->
<!--#include file="conn.asp" -->
<%
'==================================
'  Blog后台参数调用页面
'    更新时间: 2005-10-26
'==================================

'读取Blog设置信息
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

'写入首页链接列表
Bloglinks(1)

'写入自定义模块缓存
log_module(1)
%>
