<!--#include file="const.asp" -->
<!--#include file="conn.asp" -->
<!--#include file="common/function.asp" -->
<!--#include file="common/cache.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<style>
body{
 font-size:12px;
 font-family:verdana;
}
</style>
<h3>PJBlog3 V<%=blog_version%>留言本升级程序</h3><br/>
<b>开始执行升级 SQL 语句</b><br/>
 - 说明:留言本数据库升级程序；运行一次即可。
<br/>
<br/>
<div style="border:1px solid #4b372e;background:#fefded;padding:6px;line-height:160%;">
<%
'	If Not IsInteger(blog_version) Then
'		response.write "无法从此版本升级！"
'		response.end
'	End If
'	If blog_version <> 157 and blog_version <> 170 and blog_version <> 227 Then
'		response.write "无法从此版本升级！"
'		response.end
'	End If
	
 On Error Resume Next
	Conn.execute("UPDATE blog_book SET book_Url = siteurl,book_Mail = email")
	
	SQL="ALTER TABLE `blog_book` DROP COLUMN `email`"
	UpdateSQL SQL
	
	SQL="ALTER TABLE `blog_book` DROP COLUMN `siteurl`"
	UpdateSQL SQL

%>
  </div>
  <% 
Conn.Close
Set Conn=Nothing

function UpdateSQL(SQLString)
 On Error Resume Next
  Conn.execute SQLString
 if err then 
   response.write "<span style=""color:#004000""><b></b> "&err.description&"</span><br/>"
  else
   response.write "<span style=""color:#0000a0""><b>执行:</b> "&SQLString&"</span><br/>"
 end if
end function
%>
<div style="border:1px solid #4b372e;background:#fefded;padding:6px;line-height:160%;margin-top:2px">
 升级完成为了保证你的系统安全，请删除升级文件。<br/>升级后到后台 <b>"站点基本设置-初试化数据"</b> 执行 <b>"重建数据缓存"</b> 
</div>  
