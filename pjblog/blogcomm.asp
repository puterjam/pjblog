<!--#include file="commond.asp" -->
<!--#include file="header.asp" -->
<!--#include file="class/cls_logAction.asp" -->
<!--#include file="class/cls_article.asp" -->
<div id="Tbody">
    <div style="text-align:center;"><br/>
<%
'=====================================
'  评论处理页面
'    更新时间: 2006-1-12
'=====================================
If Not ChkPost() Then
    response.Write ("非法操作!!")
ElseIf Request.Form("action") = "post" Then
    '评论发表代码
    Dim PostBComm
    PostBComm = postcomm
%>
   <div style="text-align:center;">
    <div id="MsgContent" style="width:300px">
      <div id="MsgHead"><%=PostBComm(0)%></div>
      <div id="MsgBody">
		 <div class="<%=PostBComm(2)%>"></div>
         <div class="MessageText"><%=PostBComm(1)%></div>
	  </div>
	</div>
  </div>
<%
ElseIf Request.QueryString("action") = "del" Then
    Dim DelBComm
    DelBComm = delcomm
%>
   <div style="text-align:center;">
    <div id="MsgContent" style="width:300px">
      <div id="MsgHead"><%=DelBComm(0)%></div>
      <div id="MsgBody">
		 <div class="<%=DelBComm(2)%>"></div>
         <div class="MessageText"><%=DelBComm(1)%></div>
	  </div>
	</div>
  </div>
<%
Else
    response.Write ("非法操作!!")
End If

'============================ 删除评论函数 =================================================

Function delcomm
    Dim post_commID, blog_Comm, blog_CommAuthor, logid
    Dim ReInfo
    ReInfo = Array("错误信息", "", "MessageIcon")
    post_commID = CLng(CheckStr(request.QueryString("commID")))
    Set blog_Comm = Conn.Execute("select top 1 comm_ID,blog_ID,comm_Author from blog_Comment where comm_ID="&post_commID)
    If blog_Comm.EOF Or blog_Comm.bof Then
        ReInfo(0) = "错误信息"
        ReInfo(1) = "<b>不存在此评论,或该评论已经被删除!</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
        ReInfo(2) = "WarningIcon"
        delcomm = ReInfo
        Exit Function
    End If
    blog_CommAuthor = blog_Comm("comm_Author")
    If stat_Admin = True Or (stat_CommentDel = True And memName = blog_CommAuthor) Then
        ReInfo(0) = "评论删除成功"
        ReInfo(1) = "<b>评论已经被删除成功!</b><br/><a href=""default.asp?id="&blog_Comm("blog_ID")&""">单击返回</a>"
        ReInfo(2) = "MessageIcon"
        logid = Conn.Execute("select blog_ID from blog_Comment where comm_ID="&post_commID)(0)
        Conn.Execute("update blog_Content set log_CommNums=log_CommNums-1 where log_ID="&blog_Comm("blog_ID"))
        Conn.Execute("update blog_Member set mem_PostComms=mem_PostComms-1 where mem_Name='"&blog_CommAuthor&"'")
        Conn.Execute("DELETE * FROM blog_Comment WHERE comm_ID="&post_commID)
        Conn.Execute("update blog_Info set blog_CommNums=blog_CommNums-1")
        PostArticle logid, False
        getInfo(2)
        NewComment(2)
        delcomm = ReInfo
        Session(CookieName&"_LastDo") = "DelComment"
    Else
        ReInfo(0) = "错误信息"
        ReInfo(1) = "<b>你没有权限删除评论</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
        ReInfo(2) = "WarningIcon"
        delcomm = ReInfo
    End If
    call newEtag
End Function

'====================== 评论发表函数 ===========================================================

Function postcomm
    Dim username, post_logID, post_From, post_FromURL, post_disImg, post_DisSM, post_DisURL, post_DisKEY, post_DisUBB, post_Message, validate, GuestCanRemeberComment
    Dim password
    Dim ReInfo, LastMSG, FlowControl
	' New Add
	Dim Post_Email, Post_WebSite
    ReInfo = Array("错误信息", "", "MessageIcon")
    username = Trim(CheckStr(request.Form("username")))
    password = Trim(CheckStr(request.Form("password")))
    post_logID = CLng(CheckStr(request.Form("logID")))
    validate = Trim(request.Form("validate"))
    post_Message = CheckStr(request.Form("Message"))
	' --------------------------
	Post_Email = Trim(CheckStr(request.Form("Email")))
	Post_WebSite = Trim(CheckStr(request.Form("WebSite")))
	' --------------------------
	GuestCanRemeberComment = CheckStr(Request.Form("log_GuestCanRemeberComment"))
	if GuestCanRemeberComment = "1" then GuestCanRemeberComment = 1 else GuestCanRemeberComment = 0
    FlowControl = False

  If stat_Admin = False Then
    If (memName=empty or blog_validate=true) and (cstr(lcase(Session("GetCode")))<>cstr(lcase(validate)) or IsEmpty(Session("GetCode"))) Then
        ReInfo(0) = "评论发表错误信息"
        ReInfo(1) = "<b>验证码有误，请返回重新输入</b><br/><a href=""javascript:history.go(-1);"">返回重新输入</a>"
        ReInfo(2) = "ErrorIcon"
        postcomm = ReInfo
        Session("GetCode") = Empty
        Exit Function
    End If

    Set LastMSG = conn.Execute("select top 1 comm_Content from blog_Comment order by comm_ID desc")
    If LastMSG.EOF Then
        FlowControl = False
    Else
        If Trim(LastMSG("comm_Content")) = Trim(post_Message) Then FlowControl = True
    End If

    If Left(post_Message,1)= Chr(32) then
            ReInfo(0)="评论发表错误信息"
            ReInfo(1)="<b>评论内容中不允许首字空格</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
            ReInfo(2)="WarningIcon"
            postcomm = ReInfo
	    Exit Function 
	  End If
  
    If Left(post_Message,1)= Chr(13) then
            ReInfo(0)="评论发表错误信息"
            ReInfo(1)="<b>评论内容中不允许首字换行</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
            ReInfo(2)="WarningIcon"
            postcomm = ReInfo
	    Exit Function 
	  End If
	  
        '高级过滤规则
    If regFilterSpam(post_Message, "reg.xml") Then
            ReInfo(0) = "评论发表错误信息"
            ReInfo(1) = "<b>评论中包含被屏蔽的字符</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
            ReInfo(2) = "WarningIcon"
            postcomm = ReInfo
      Exit Function
    End If

        '基本过滤规则
    If filterSpam(post_Message, "spam.xml") Then
            ReInfo(0) = "评论发表错误信息"
            ReInfo(1) = "<b>评论中包含被屏蔽的字符</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
            ReInfo(2) = "WarningIcon"
            postcomm = ReInfo
      Exit Function
    End If

    If FlowControl Then
        ReInfo(0) = "评论发表错误信息"
        ReInfo(1) = "<b>禁止恶意灌水！</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
        ReInfo(2) = "WarningIcon"
        postcomm = ReInfo
        Exit Function
    End If

    If DateDiff("s", Request.Cookies(CookieName)("memLastPost"), DateToStr(now(),"Y-m-d H:I:S"))<blog_commTimerout Then
        ReInfo(0) = "评论发表错误信息"
        ReInfo(1) = "<b>发言太快,请 "&blog_commTimerout&" 秒后再发表评论</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
        ReInfo(2) = "WarningIcon"
        postcomm = ReInfo
      Exit Function
    End If
  End If
    
    If Len(username)<1 Then
        ReInfo(0) = "评论发表错误信息"
        ReInfo(1) = "<b>请输入你的昵称.</b><br/><a href=""javascript:history.go(-1);"">返回重新输入</a>"
        ReInfo(2) = "ErrorIcon"
        postcomm = ReInfo
        Exit Function
    End If

    If IsValidUserName(username) = False Then
        ReInfo(0) = "错误信息"
        ReInfo(1) = "<b>非法用户名！<br/>请尝试使用其他用户名！</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
        ReInfo(2) = "ErrorIcon"
        postcomm = ReInfo
        Exit Function
    End If

    Dim checkMem
    If memName = Empty Then'匿名评论
        If Len(password)>0 Then
            Dim loginUser
            loginUser = login(Request.Form("username"), Request.Form("password"))
            If Not request.Cookies(CookieName)("memName") = username Then
                ReInfo(0) = "评论发表错误信息"
                ReInfo(1) = "<b>登录失败，请检查用户名和密码</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
                ReInfo(2) = "WarningIcon"
                postcomm = ReInfo
                Exit Function
            End If
        Else
            Set checkMem = Conn.Execute("select top 1 mem_id from blog_Member where mem_Name='"&username&"'")
            If Not checkMem.EOF Then
                ReInfo(0) = "评论发表错误信息"
                ReInfo(1) = "<b>该用户已经存在，无法发表评论</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
                ReInfo(2) = "WarningIcon"
                postcomm = ReInfo
                Exit Function
            End If
        End If
    Else
 			If Not request.Cookies(CookieName)("memName") = username Then
                ReInfo(0) = "评论发表错误信息"
                ReInfo(1) = "<b>请输入正确的用户名</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
                ReInfo(2) = "WarningIcon"
                postcomm = ReInfo
                Exit Function
            End If
    End If
    
    If Not stat_CommentAdd Then
        ReInfo(0) = "评论发表错误信息"
        ReInfo(1) = "<b>你没有权限发表评论</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
        ReInfo(2) = "ErrorIcon"
        postcomm = ReInfo
        Exit Function
    End If
    If Conn.Execute("select log_DisComment from blog_Content where log_ID="&post_logID)(0) Then
        ReInfo(0) = "评论发表错误信息"
        ReInfo(1) = "<b>该日志不允许发表任何评论</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
        ReInfo(2) = "WarningIcon"
        postcomm = ReInfo
        Exit Function
    End If
    post_DisSM = request.Form("log_DisSM")
    post_DisURL = request.Form("log_DisURL")
    post_DisKEY = request.Form("log_DisKey")

    If Len(post_Message)<1 Then
        ReInfo(0) = "评论发表错误信息"
        ReInfo(1) = "<b>不允许发表空评论</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
        ReInfo(2) = "ErrorIcon"
        postcomm = ReInfo
        Exit Function
    End If
    If Len(post_Message)>blog_commLength Then
        ReInfo(0) = "评论发表错误信息"
        ReInfo(1) = "<b>评论超过最大字数限制</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
        ReInfo(2) = "ErrorIcon"
        postcomm = ReInfo
        Exit Function
    End If
	
	If Len(Post_Email) > 0 Then
		If Not RegMatch(Post_Email, "\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*") Then
			ReInfo(0) = "评论发表错误信息"
        	ReInfo(1) = "<b>验证邮箱格式错误</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
        	ReInfo(2) = "ErrorIcon"
        	postcomm = ReInfo
        	Exit Function
		End If
	End If
	
	If Len(Post_WebSite) > 0 Then
		If Not RegMatch(Post_WebSite, "(http):\/\/[\w\-_]+(\.[\w\-_]+)+([\w\-\.,@?^=%&:/~\+#]*[\w\-\@?^=%&/~\+#])?") Then
			ReInfo(0) = "评论发表错误信息"
        	ReInfo(1) = "<b>网址格式错误,请用http://开头</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
        	ReInfo(2) = "ErrorIcon"
        	postcomm = ReInfo
        	Exit Function
		End If
	End If
    'UBB 特别属性
    post_disImg = 1
    post_DisUBB = 0
    If post_DisSM = 1 Then post_DisSM = 1 Else post_DisSM = 0
    If post_DisURL = 1 Then post_DisURL = 0 Else post_DisURL = 1
    If post_DisKEY = 1 Then post_DisKEY = 0 Else post_DisKEY = 1
    '插入数据
    Dim AddComm,re
    
    '去掉回复标签，不允许直接输入
    Set re = New RegExp
    re.IgnoreCase = True
    re.Global = True
    re.Pattern = "\[reply=(.[^\]]*)\](.*?)\[\/reply\]"
    post_Message = re.Replace(post_Message, "$2")

	'记住信息
	If GuestCanRemeberComment = 1 Then
		Response.Cookies(CookieName)("Guest") = ("true|-|" & username & "|$|" & Post_Email & "|$|" & Post_WebSite & "|+|")
		Response.Cookies(CookieName).Expires = Date+365
    	Response.Cookies(CookieName)("exp") = DateAdd("d", 365, date())
	Else
		Response.Cookies(CookieName)("Guest") = ("false|-|null|+|")
		Response.Cookies(CookieName).Expires = Date+365
    	Response.Cookies(CookieName)("exp") = DateAdd("d", 365, date())
	End If
	
    AddComm = Array(Array("blog_ID", post_logID), Array("comm_Content", post_Message), Array("comm_Author", username), Array("comm_DisSM", post_DisSM), Array("comm_DisUBB", post_DisUBB), Array("comm_DisIMG", post_disImg), Array("comm_AutoURL", post_DisURL), Array("comm_PostIP", getIP), Array("comm_AutoKEY", post_DisKEY), Array("comm_Email", Post_Email), Array("comm_WebSite", Post_WebSite))
    DBQuest "blog_Comment", AddComm, "insert"
    'Conn.ExeCute("INSERT INTO blog_Comment(blog_ID,comm_Content,comm_Author,comm_DisSM,comm_DisUBB,comm_DisIMG,comm_AutoURL,comm_PostIP,comm_AutoKEY) VALUES ("&post_logID&",'"&post_Message&"','"&username&"',"&post_DisSM&","&post_DisUBB&","&post_disImg&","&post_DisURL&",'"&getIP()&"',"&post_DisKEY&")")
    Conn.Execute("update blog_Content set log_CommNums=log_CommNums+1 where log_ID="&post_logID)
    Conn.Execute("update blog_Info set blog_CommNums=blog_CommNums+1")
    Response.Cookies(CookieName)("memLastpost") = DateToStr(now(),"Y-m-d H:I:S")
    getInfo(2)
    NewComment(2)
    If memName<>Empty Then
        conn.Execute("update blog_Member set mem_PostComms=mem_PostComms+1 where mem_Name='"&memName&"'")
    End If
    SQLQueryNums = SQLQueryNums + 3
	
    '评论邮件通知
    If blog_Isjmail Then
		Dim email_commid, SQLcomm, log_commcomm
		SQLcomm="Select TOP 1 * FROM blog_Comment Where comm_Author='"&username&"' order By comm_ID Desc "
		Set log_commcomm=conn.execute(SQLcomm) 
			email_commid=log_commcomm("comm_ID")
		log_commcomm.Close
		Set log_commcomm=Nothing
		dim email_log_title
		SQLcomm="Select * FROM blog_Content Where log_ID="&post_logID&""
		Set log_commcomm=conn.execute(SQLcomm) 
			email_log_title=log_commcomm("log_Title")
		log_commcomm.Close
		Set log_commcomm=Nothing
        dim emailcontent,emailtitle
        emailtitle = "您发表的文章《"&email_log_title&"》已有客人发表了评论"
        if blog_postFile = 2 then
            emailcontent = "["&username&"]在您的博客中发表了评论,请点击查"&siteURL&caload(post_logID)&"#comm_"&email_commid&"。评论内容如下："&post_Message&""
        else 
            emailcontent = "["&username&"]在您的博客中发表了评论,请点击查"&siteURL&"default.asp?id="&post_logID&"#comm_"&email_commid&"。评论内容如下："&post_Message&""
        end if
        call sendmail(blog_email,emailtitle,emailcontent,sitename)
    End If
    '评论邮件通知结束

	
    ReInfo(0) = "评论发表成功"
    ReInfo(1) = "<b>你成功地对该日志发表了评论</b><br/><a href=""default.asp?id="&post_logID&""">单击返回该日志</a>"
    ReInfo(2) = "MessageIcon"
    Session("GetCode") = Empty
    Session(CookieName&"_LastDo") = "AddComment"
    postcomm = ReInfo
    PostArticle post_logID, False
    call newEtag
End Function
%>
  <br/></div> 
 </div>
<!--#include file="footer.asp" -->
