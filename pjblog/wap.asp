<!--#include file="const.asp" -->
<!--#include file="conn.asp" -->
<!--#include file="common/cache.asp" -->
<!--#include file="common/function.asp" -->
<!--#include file="common/ubbcode.asp" -->
<!--#include file="common/checkUser.asp" -->
<!--#include file="class/cls_wap.asp" -->
<!--#include file="class/cls_logAction.asp" -->
<!--#include file="class/cls_article.asp" -->
<!--#include file="class/cls_fso.asp" -->
<!--#include file="common/library.asp" -->
<%
'==================================
'  wap For PJBlog2
'    更新时间: 2006-6-12
'==================================
Response.Charset = "UTF-8"
Response.ContentType = "text/vnd.wap.wml"

Response.Expires = -1
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "Cache-Control", "no-cache, must-revalidate"
'读取Blog设置信息
getInfo(1)

'验证用户登录信息
checkCookies

'读取用户权限
UserRight(1)

'写入标签
Tags(1)

'写入关键字列表
Keywords(1)

'写入表情符号
Smilies(1)
%>
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
	<%
Select Case request.QueryString("do")
Case "showTag":
    outTagList
Case "Login":
    outLogin
Case "Logout":
    outLogout
Case "CheckUser":
    outLoginUser
Case "showLog":
    outLog
Case "showLogDetail":
    outLogDetail
Case "showComment":
    outComment
Case "postComment":
    outPostComment
Case "postArtile":
    outwap_AritclePost
Case "editLog":
    outEditLog
Case "Articleadd":
    Articleadd
Case Else: 'Wap Default Page
    outDefault
End Select

Call CloseDB
%>
</wml>
