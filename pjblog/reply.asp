<!--#include file="BlogCommon.asp" -->
﻿<!--#include file="common/function.asp" -->
<!--#include file="common/library.asp" -->
<!--#include file="common/cache.asp" -->
<!--#include file="common/checkUser.asp" -->
<!--#include file="common/ModSet.asp" -->
<!--#include file="class/cls_article.asp" -->
<!--#include file="class/cls_logAction.asp" -->
<!--#include file="common/ubbcode.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="UTF-8">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />

<%
Sub replyComm
	dim cID,replay,quest,result,a1,a2,a3,logId
	logId = CheckStr(Request.Form("logId"))
	cID = CheckStr(Request.Form("id"))
	a1 = Int(Request.Form("a1"))
	a2 = Int(Request.Form("a2"))
	a3 = Int(Request.Form("a3"))
	
	replay = CheckStr(Request.Form("replay"))
	
	if isEmpty(memName) Or stat_Admin <> True then
		response.write ("<script>parent.removeReplyMsg("&cID&");alert('您没有权限回复评论。')</script>")
		exit Sub
	end if 
	

	
	if isEmpty(logId) or isEmpty(cID) or isEmpty(replay) then
		response.write ("<script>parent.removeReplyMsg("&cID&");alert('您没有权限回复评论。')</script>")
		exit sub
	end if
	
	if isEmpty(replay) then
		response.write ("<script>parent.removeReplyMsg("&cID&");alert('您没有输入评论内容')</script>")
		exit sub
	end if
	
'	set quest = Conn.Execute("select top 1 comm_Content from blog_Comment where comm_ID=" & cID)
    set quest = Conn.Execute("select top 1 a.comm_Content,a.comm_Author,b.log_title,b.log_ID,a.comm_Email from blog_Comment a inner join blog_Content b on a.blog_ID=b.log_ID where a.comm_ID=" & cID)

 	If not quest.EOF Then
		result = quest(0) & vbcrlf & "[reply=" + memName + "," & DateToStr(now(),"Y-m-d H:I A") & "]" & replay & "[/reply]"
  		conn.Execute("UPDATE blog_Comment SET comm_Content='"&result&"' WHERE comm_ID="&cID)
  		
  		dim ubbResult
  		 ubbResult  = UBBCode(HtmlEncode(result),cBool(a1),blog_commUBB,blog_commIMG,cBool(a2),cBool(a3))
		response.write ("<script>parent.$(""commcontent_"&cID&""").innerHTML = '"&output(ubbResult)&"';</script>")
		

       If blog_reply_Isjmail and trim(quest(4))<>"" Then
            dim emailcontent, emailtitle, CommUrl
            emailtitle = "您在 "&siteName&" 上发表的评论已有了新回复！"
            if blog_postFile = 2 then
                CommUrl = "<a href="""&siteurl&caload(quest(3))&"#comm_"&cID&"""  target=""_blank"">详情请点击查看</a>"
            else 
                CommUrl = "<a href="""&siteurl&"default.asp?id="&quest(3)&"#comm_"&cID&"""  target=""_blank"">详情请点击查看</a>"
            end if
            emailcontent = "尊敬的 <strong>"&quest(1)&"</strong> ，您好！<br/>您在博客 <strong>"&siteName&"</strong> 上对日志<strong>《"&quest(2)&"》</strong>发表的评论<br/>博主 <strong>"&memName&"</strong> 已经有了新的回复，回复内容为： <div style=""margin:10px 0px;padding:10px;width:400px;border:1px solid #999;border-bottom-left-radius:4px 4px;border-bottom-right-radius:4px 4px;border-top-left-radius:4px 4px;border-top-right-radius:4px 4px;color:#574D31;background:#F0ECD0;""><strong>您 </strong>说："&DelQuote(quest(0))&"<br/><strong>"&memName&" </strong>回复："&replay&"<br/><br/>"&CommUrl&"</div>谢谢您参与评论，欢迎您再次光临本博客！<br/>本邮件系统自动发送，请勿直接回复。"
            call sendmail(quest(4),emailtitle,emailcontent,quest(1))
	end if

		 PostArticle logId, False
    End If

	quest.close
	set quest = nothing
End Sub

call replyComm
%>
<script runat="server" Language="javascript">
	function output(html){
		html = html.replace(/[\n\r]/g,"");
		html = html.replace(/\\/g,"\\\\");
		html = html.replace(/\'/g,"\\'");
		return html
	}
</script>
</head>
<body></body></html>
