<!--#include file="BlogCommon.asp" -->
<!--#include file="common/function.asp" -->
<!--#include file="common/cache.asp" -->
<!--#include file="common/checkUser.asp" -->
<%

Dim action, id, typee
action = Checkxss(Request("action"))
id = Checkxss(Request("id"))
typee = Checkxss(Request("typee"))
If IsBlank(id) Then Response.End
if not IsInteger(id) Then Response.End
Dim digg,showdigg,good,bad
Set digg = conn.execute("Select log_digggood,log_diggbad From blog_Content Where log_id="&id)
  If digg.EOF And digg.BOF Then
    Response.Write "1"
    digg.Close
    Set digg = Nothing
  Else
    good = digg(0)
    bad = digg(1)
    If IsBlank(good) Then conn.execute("Update blog_Content set log_digggood=0 where log_id="&id) : good = 0
    If IsBlank(bad) Then conn.execute("Update blog_Content set log_diggbad=0 where log_id="&id) : bad = 0
    showdigg = good&","&bad
    digg.Close
    Set digg = Nothing
    If action="show" then
      Response.Write showdigg
    ElseIf action = "postDigg" Then
      If Request.Cookies(CookieName)("digg"&id) <> id Then
        If typee = "good" Then
          good = good + 1
        ElseIf typee = "bad" Then
          bad = bad +1
        Else
          Response.Write "0"
          Response.End
        End If
        conn.execute("Update blog_Content set log_digggood="&good&" where log_id="&id)
        conn.execute("Update blog_Content set log_diggbad="&bad&" where log_id="&id)
        Response.Cookies(CookieName)("digg"&id) = id
	    Response.Write good&","&bad
      Else
	    Response.Write "1"
      End If
    End If
  End If
 %>