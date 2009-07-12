<!--#include file="../../const.asp" -->
<!--#include file="../../p_conn.asp" -->
<!--#include file="../../common/function.asp" -->
<!--#include file="../../common/library.asp" -->
<!--#include file="../../common/cache.asp" -->
<!--#include file="../../common/checkUser.asp" -->
<!--#include file="../../common/ModSet.asp" -->
<%
'=====================================
'  留言插件信息处理页面
'    更新时间: 2005-10-24
'=====================================
%>

<%
dim GBSet
dim charCount,delay,OpenState
Set GBSet=New ModSet
GBSet.open("GuestBookForPJBlog")
if not GBSet.PasreError<>-18903 then
   showmsg "错误信息","留言本插件没有安装<br/><a href=""javascript:history.go(-1)"">单击返回</a>","MessageIcon","plugins"
end if
charCount=GBSet.getKeyValue("charCount")
delay=GBSet.getKeyValue("delay")
OpenState=GBSet.getKeyValue("OpenState")

if not cBool(OpenState) then
   showmsg "错误信息","留言本暂时关闭！<br/><a href=""default.asp"">单击返回首页</a>","WarningIcon","plugins"
end if

if request.form("action")="post" then
   postMsg '发表留言
 elseif Request.QueryString("action")="del" then 
   delMsg  '删除留言
 elseif Request.form("action")="reply" then 
   replyMsg '回复留言
 else
   showmsg "错误信息","非法操作！<br/><a href=""javascript:history.go(-1)"">单击返回</a>","ErrorIcon","plugins"
end if

'============================= 发表留言 ========================================
function postMsg
 dim username,post_Message,validate,hiddenreply,face
 dim password,LastMSG,FlowControl
  username=trim(CheckStr(request.form("username")))
  password=trim(CheckStr(request.form("password")))
  face=CheckStr(request.form("book_face"))
  validate=trim(request.form("validate"))
  hiddenreply=request.form("hiddenMsg")
  post_Message=CheckStr(request.form("Message"))
dim email,tsiteURL
  email=trim(CheckStr(request.form("myblogemail")))
  tsiteURL=trim(CheckStr(request.form("myblogsiteurl")))

  FlowControl=false
  
  set LastMSG=conn.execute("select top 1 book_Content from blog_book order by book_ID desc")
  if LastMSG.eof then 
     FlowControl=false
   else
    if LastMSG("book_Content")=post_Message then FlowControl=true
  end if

    if memName=empty and email<>"" and IsValidEmail(email)=false then
        showmsg "留言发表错误信息","<b>邮箱格式错误</b><br/><a href=""javascript:history.go(-1);"">返回</a>","WarningIcon","plugins"
        exit function
    end if

    if memName=empty and tsiteURL<>"" and tsiteURL<>"http://" and IsRightUrl(tsiteURL)=false then
          showmsg "留言发表错误信息","<b>网址格式错误</b><br/><a href=""javascript:history.go(-1);"">返回</a>","WarningIcon","plugins"
          exit function 
    end if


  if filterSpam(post_Message,"../../spam.xml") and stat_Admin=false then
      showmsg "留言发表错误信息","<b>留言中包含被屏蔽的字符</b><br/><a href=""javascript:history.go(-1);"">返回</a>","WarningIcon","plugins"
      exit function 
  end if

	if regFilterSpam(post_Message,"../../reg.xml") and stat_Admin=false then
	    showmsg "留言发表错误信息","<b>留言中包含被屏蔽的字符</b><br/><a href=""javascript:history.go(-1);"">返回</a>","WarningIcon","plugins"
	    exit function 
	end if


  if FlowControl then 
      showmsg "留言发表错误信息","<b>禁止恶意灌水！</b><br/><a href=""LoadMod.asp?plugins=GuestBookForPJBlog"">返回</a>","WarningIcon","plugins"
      exit function 
  end if 
  
  if DateDiff("s",Request.Cookies(CookieName)("bookLastPost"),DateToStr(now(),"Y-m-d H:I:S"))<int(delay) then 
      showmsg "留言发表错误信息","<b>发言太快,请 "&delay&" 秒后再发表留言</b><br/><a href=""LoadMod.asp?plugins=GuestBookForPJBlog"">返回</a>","WarningIcon","plugins"
      exit function  
  end if
  
  if len(username)<1 then
      showmsg "留言发表错误信息","<b>请输入你的昵称.</b><br/><a href=""javascript:history.go(-1);"">请返回重新输入</a>","ErrorIcon","plugins"
      exit function  
  end if
  
  IF cstr(lcase(Session("GetCode")))<>cstr(lcase(validate)) Or IsEmpty(Session("GetCode")) then
      showmsg "留言发表错误信息","<b>验证码有误，请返回重新输入</b><br/><a href=""LoadMod.asp?plugins=GuestBookForPJBlog"">请返回重新输入</a>","ErrorIcon","plugins"
      exit function
  end if
  
  
  dim checkMem
  if memName=empty then
    if len(password)>0 then
        Dim loginUser
        loginUser=login(Request.Form("username"),Request.Form("password"))
         if not request.Cookies(CookieName)("memName")=username then
            showmsg "留言发表错误信息","<b>登陆失败，请检查用户名和密码</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>","WarningIcon","plugins"
            exit function
         end if
    else
       set checkMem=Conn.ExeCute("select * from blog_Member where mem_Name='"&username&"'")
       if not checkMem.eof then
         showmsg "留言发表错误信息","<b>该用户已经存在，无法发表留言</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>","WarningIcon","plugins"
    	 exit function
       end if
    end if
  end if 

  if len(post_Message)<1 then
     showmsg "留言发表错误信息","<b>不允许发表空留言信息</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>","WarningIcon","plugins"
	 exit function
  end if
  if len(post_Message)>int(charCount) then
     showmsg "留言发表错误信息","留言超过最大字数限制<br/><a href=""javascript:history.go(-1);"">单击返回</a>","ErrorIcon","plugins"
	 exit function
  end if
  if hiddenreply=1 then hiddenreply=true else hiddenreply=false
  if memName=empty and hiddenreply then
     showmsg "留言发表错误信息","必须登陆才可以发表隐藏留言<br/><a href=""javascript:history.go(-1);"">单击返回</a>","ErrorIcon","plugins"
	 exit function
  end if
 '插入数据
Conn.ExeCute("Insert INTO blog_book(book_Messager,book_face,book_IP,book_Content,book_HiddenReply,email,siteurl) VALUES ('"&username&"','"&face&"','"&getIP()&"','"&post_Message&"',"&hiddenreply&",'"&email&"','"&tsiteURL&"')")

 Conn.ExeCute("update blog_Info set blog_MessageNums=blog_MessageNums+1")
 if memName<>empty then
   conn.execute("update blog_Member set mem_PostMessageNums=mem_PostMessageNums+1 where mem_Name='"&memName&"'")
 end if
 Response.Cookies(CookieName)("bookLastPost")=DateToStr(now(),"Y-m-d H:I:S")
 
	'留言邮件通知
	If blog_Isjmail Then
		Dim SQLcomm, log_commcomm
		SQLcomm="Select TOP 1 * FROM blog_book Where book_Messager='"&username&"' order By book_ID Desc "
		dim email_bookid
		Set log_commcomm=conn.execute(SQLcomm) 
			email_bookid=log_commcomm("book_ID")
		log_commcomm.Close
		Set log_commcomm=Nothing
			dim emailcontent,emailtitle
			emailtitle = "您的博客已有客人留言"
			emailcontent = "["&username&"]在您的博客中发表了留言,请点击查看"&siteURL&"LoadMod.asp?plugins=GuestBookForPJBlog#book_"&email_bookid&"。留言内容如下："&post_Message&""
			call sendmail(blog_email,emailtitle,emailcontent)
	'		call sendmail(username,"",email_bookid,"",0,post_Message)
	End If
	'留言邮件通知结束


 getInfo(2)
 EmptyEtag
 SQLQueryNums=SQLQueryNums+3
 reloadMsg
 showmsg "留言发表信息","<b>你成功地对该日志发表了留言</b><br/><a href=""LoadMod.asp?plugins=GuestBookForPJBlog"">单击返回留言本</a>","MessageIcon","plugins" 
end function

'==================================== 删除留言 ===============================================
function delMsg
 dim book_ID,bookDB,PostMessager
  book_ID=CheckStr(request.QueryString("id"))
  set bookDB=Conn.ExeCute("select * from blog_book where book_ID="&book_ID)
  if bookDB.eof or bookDB.bof then
     showmsg "错误信息","<b>不存在此留言,或该评论已经被删除!</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>","ErrorIcon","plugins"
	 exit function
  end if
  PostMessager=bookDB("book_Messager")
  if (memName<>empty and stat_Admin) or (cbool(GBSet.getKeyValue("canDel")) and Lcase(PostMessager)=Lcase(memName)) then
     Conn.ExeCute("DELETE * FROM blog_book WHERE book_ID="&book_ID)
     Conn.ExeCute("update blog_Info set blog_MessageNums=blog_MessageNums-1")
     getInfo(2)
 	 reloadMsg
     showmsg "留言删除成功","<b>留言已经被删除成功!</b><br/><a href=""LoadMod.asp?plugins=GuestBookForPJBlog"">单击返回</a>","MessageIcon","plugins"
  else
     showmsg "错误信息","<b>你没有权限删除该留言</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>","ErrorIcon","plugins"
  end if
	EmptyEtag
end function 

'==================================== 回复留言留言 ===============================================
function replyMsg
  dim MsgReplyContent,MsgID
  MsgID = CheckStr(Request.form("MsgID"))
  MsgReplyContent=CheckStr(request.form("Message"))
	   if not (memName<>empty and stat_Admin) then
                showmsg "错误信息","你没有权限回复留言<br/><a href=""javascript:history.go(-1)"">单击返回</a>","WarningIcon","plugins"
	   end if
	   If MsgID=Empty then 
	            showmsg "错误信息","非法操作<br/><a href=""javascript:history.go(-1)"">单击返回</a>","WarningIcon","plugins" 
	   end if
	   If IsInteger(MsgID)=False then 
	            showmsg "错误信息","非法操作<br/><a href=""javascript:history.go(-1)"">单击返回</a>","WarningIcon","plugins" 
	   end if
   Conn.ExeCute("update blog_book set book_reply='"&MsgReplyContent&"',book_replyAuthor='"&memName&"',book_replyTime=#"&DateToStr(now(),"Y-m-d H:I:S")&"# where book_ID=" & MsgID)

    '留言邮件通知
       If blog_reply_Isjmail Then
			Dim SQLcomm, log_commcomm
			SQLcomm="Select TOP 1 * FROM blog_book Where book_ID="&MsgID
			Set log_commcomm=conn.execute(SQLcomm) 
			if trim(log_commcomm("email"))<>"" then
				dim emailcontent,emailtitle
				emailtitle = "您在"&siteName&"上发表的留言已被回复"
				emailcontent = "尊敬的｛"&log_commcomm("book_Messager")&"｝，您好，您在["&siteName&"]上发表的留言，现已被["&memName&"]回复，回复内容为：["&MsgReplyContent&"]，请点击查看"&siteURL&"LoadMod.asp?plugins=GuestBookForPJBlog#book_"&MsgID&"。谢谢您的留言，欢迎再次光临！系统自动发送，请勿直接回复。"
				call sendmail(log_commcomm("email"),emailtitle,emailcontent)
		'		call sendmail(username,"",email_bookid,"",0,post_Message)
			end if
			log_commcomm.Close
			Set log_commcomm=Nothing
		End If
    '留言邮件通知结束

   showmsg "回复信息","回复留言成功!<br/><a href=""LoadMod.asp?plugins=GuestBookForPJBlog"">单击返回留言本</a>","MessageIcon","plugins" 
end function 

function reloadMsg
           	Dim book_Messages,book_Message,blog_Message
             	Set book_Messages=Conn.Execute("SELECT top 10 book_ID,book_Messager,book_PostTime,book_Content,book_HiddenReply FROM blog_book order by book_PostTime Desc")
             	TempVar=""
             	Do While Not book_Messages.EOF
             	    if book_Messages("book_HiddenReply") then
                        book_Message=book_Message&TempVar&book_Messages("book_ID")&"|,|"&book_Messages("book_Messager")&"|,|"&book_Messages("book_PostTime")&"|,|"&"[隐藏留言]"
             	     else
                        book_Message=book_Message&TempVar&book_Messages("book_ID")&"|,|"&book_Messages("book_Messager")&"|,|"&book_Messages("book_PostTime")&"|,|"&book_Messages("book_Content")
             	    end if
             		TempVar="|$|"
             		book_Messages.MoveNext
             	Loop
             	Set book_Messages=Nothing
             	blog_Message=Split(book_Message,"|$|")
             	Application.Lock
             	Application(CookieName&"_blog_Message")=blog_Message
             	Application.UnLock
				EmptyEtag
end function
%>

