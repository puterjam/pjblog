<!--#include file="BlogCommon.asp" -->
<!--#include file="common/function.asp" -->
<!--#include file="common/library.asp" -->
<!--#include file="common/upfile.asp" -->
<!--#include file="common/cache.asp" -->
<!--#include file="common/checkUser.asp" -->
<%
If request("action")="checkAlias" then
dim strcname,checkcdb
strcname=request("cname")
set checkcdb=conn.execute("select * from blog_Content where log_cname="""&strcname&"""")
if checkcdb.bof or checkcdb.eof then
response.write"<img src=""images/check_right.gif"">"
else
response.write"<img src=""images/check_error.gif"">"
end if
set checkcdb=nothing

elseif request("action")="type1" then
response.expires=-1 
response.expiresabsolute=now()-1 
response.cachecontrol="no-cache"
dim mainurl,main,mainstr
mainurl=request("mainurl")
main=trim(checkstr(request("main")))
response.clear()
mainstr=""
If Len(memName)>0 Then
mainstr=mainstr&"<img src=""images/download.gif"" alt=""下载文件"" style=""margin:0px 2px -4px 0px""/> <a href="""&mainurl&""" target=""_blank"">"&main&"</a>"
Else
mainstr=mainstr&"<img src=""images/download.gif"" alt=""只允许会员下载"" style=""margin:0px 2px -4px 0px""/> 该文件只允许会员下载! <a href=""login.asp"">登录</a> | <a href=""register.asp"">注册</a>"
End If
response.write mainstr

elseif request("action")="type2" then
response.expires=-1 
response.expiresabsolute=now()-1 
response.cachecontrol="no-cache"
dim main2,mstr
main2=request("main")
response.clear()
mstr=""
If Len(memName)>0 Then
mstr=mstr&"<img src=""images/download.gif"" alt=""下载文件"" style=""margin:0px 2px -4px 0px""/> <a href="""&main2&""" target=""_blank"">下载此文件</a>"
Else
mstr=mstr&"<img src=""images/download.gif"" alt=""只允许会员下载"" style=""margin:0px 2px -4px 0px""/> 该文件只允许会员下载! <a href=""login.asp"">登录</a> | <a href=""register.asp"">注册</a>"
End If
response.write mstr

else
response.write "非法操作!"
End If
%>
