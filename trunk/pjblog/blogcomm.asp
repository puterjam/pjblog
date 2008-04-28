<!--#include file="commond.asp" -->
<!--#include file="header.asp" -->
<!--#include file="class/cls_logAction.asp" -->
<div id="Tbody">
    <div style="text-align:center;"><br/>
<%
'=====================================
'  评论处理页面
'    更新时间: 2006-1-12
'=====================================
IF not ChkPost() Then
 response.write ("非法操作!!")
elseif Request.Form("action")="post" Then
 '评论发表代码
Dim PostBComm
PostBComm=postcomm
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
elseif Request.QueryString("action")="del" then
Dim DelBComm
DelBComm=delcomm
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
else
 response.write ("非法操作!!")
end if

'============================ 删除评论函数 =================================================
function delcomm
 dim post_commID,blog_Comm,blog_CommAuthor,logid
 dim ReInfo
  ReInfo=Array("错误信息","","MessageIcon")
  post_commID=clng(CheckStr(request.QueryString("commID")))
  set blog_Comm=Conn.ExeCute("select top 1 comm_ID,blog_ID,comm_Author from blog_Comment where comm_ID="&post_commID)
  if blog_Comm.eof or blog_Comm.bof then
	 ReInfo(0)="错误信息"
	 ReInfo(1)="<b>不存在此评论,或该评论已经被删除!</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
	 ReInfo(2)="WarningIcon"
	 delcomm=ReInfo
	 exit function
  end if
  blog_CommAuthor=blog_Comm("comm_Author")
  if stat_Admin=true or (stat_CommentDel=true and memName=blog_CommAuthor) then
	 ReInfo(0)="评论删除成功"
	 ReInfo(1)="<b>评论已经被删除成功!</b><br/><a href=""default.asp?id="&blog_Comm("blog_ID")&""">单击返回</a>"
	 ReInfo(2)="MessageIcon"
     logid=Conn.ExeCute("select blog_ID from blog_Comment where comm_ID="&post_commID)(0)
     Conn.ExeCute("update blog_Content set log_CommNums=log_CommNums-1 where log_ID="&blog_Comm("blog_ID"))
     Conn.ExeCute("DELETE * FROM blog_Comment WHERE comm_ID="&post_commID)
     Conn.ExeCute("update blog_Info set blog_CommNums=blog_CommNums-1")
     PostArticle logid
     getInfo(2)
     NewComment(2)
 	 delcomm=ReInfo
 	 Session(CookieName&"_LastDo")="DelComment"
  else
	 ReInfo(0)="错误信息"
	 ReInfo(1)="<b>你没有权限删除评论</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
	 ReInfo(2)="WarningIcon"
	 delcomm=ReInfo
  end if
end function

'====================== 评论发表函数 ===========================================================
function postcomm
 dim username,post_logID,post_From,post_FromURL,post_disImg,post_DisSM,post_DisURL,post_DisKEY,post_DisUBB,post_Message,validate
 dim password
 dim ReInfo,LastMSG,FlowControl
  ReInfo=Array("错误信息","","MessageIcon")
  username=trim(CheckStr(request.form("username")))
  password=trim(CheckStr(request.form("password")))
  post_logID=clng(CheckStr(request.form("logID")))
  validate=trim(request.form("validate"))
  post_Message=CheckStr(request.form("Message"))
  FlowControl=false
  
  IF (memName=empty or blog_validate=true) and cstr(lcase(Session("GetCode")))<>cstr(lcase(validate)) Or IsEmpty(Session("GetCode")) then
	  ReInfo(0)="评论发表错误信息"
	  ReInfo(1)="<b>验证码有误，请返回重新输入</b><br/><a href=""javascript:history.go(-1);"">请返回重新输入</a>"
	  ReInfo(2)="ErrorIcon"
	  postcomm=ReInfo
	  Session("GetCode") = empty
      exit function
  end if
  
  set LastMSG=conn.execute("select top 1 comm_Content from blog_Comment order by comm_ID desc")
  if LastMSG.eof then 
     FlowControl=false
   else
    if trim(LastMSG("comm_Content")) = trim(post_Message) then FlowControl=true
  end if
  
  if stat_Admin = false then
  	  '高级过滤规则
	  if regFilterSpam(post_Message,"reg.xml") then
	  	  ReInfo(0)="评论发表错误信息"
		  ReInfo(1)="<b>评论中包含被屏蔽的字符</b><br/><a href=""javascript:history.go(-1);"">返回</a>"
		  ReInfo(2)="ErrorIcon"
		  postcomm=ReInfo
	      exit function 
	  end if
	  
  	  '基本过滤规则
	  if filterSpam(post_Message,"spam.xml") then
	  	  ReInfo(0)="评论发表错误信息"
		  ReInfo(1)="<b>评论中包含被屏蔽的字符</b><br/><a href=""javascript:history.go(-1);"">返回</a>"
		  ReInfo(2)="WarningIcon"
		  postcomm=ReInfo
	      exit function 
	  end if
  end if
  
  if FlowControl then 
  	  ReInfo(0)="评论发表错误信息"
	  ReInfo(1)="<b>禁止恶意灌水！</b><br/><a href=""javascript:history.go(-1);"">返回</a>"
	  ReInfo(2)="WarningIcon"
	  postcomm=ReInfo
      exit function 
  end if 

  if DateDiff("s",Request.Cookies(CookieName)("memLastPost"),Now())<blog_commTimerout then 
	  ReInfo(0)="评论发表错误信息"
	  ReInfo(1)="<b>发言太快,请 "&blog_commTimerout&" 秒后再发表评论</b><br/><a href=""javascript:history.go(-1);"">返回</a>"
	  ReInfo(2)="WarningIcon"
	  postcomm=ReInfo
      exit function  
  end if
  if len(username)<1 then
	  ReInfo(0)="评论发表错误信息"
	  ReInfo(1)="<b>请输入你的昵称.</b><br/><a href=""javascript:history.go(-1);"">请返回重新输入</a>"
	  ReInfo(2)="ErrorIcon"
	  postcomm=ReInfo
      exit function  
  end if
  
  if IsValidUserName(username)=false then
	 ReInfo(0)="错误信息"
	 ReInfo(1)="<b>非法用户名！<br/>请尝试使用其他用户名！</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
	 ReInfo(2)="ErrorIcon"
	 postcomm=ReInfo
	 exit function
 end if
  
  dim checkMem
  if memName=empty then
    if len(password)>0 then
        Dim loginUser
        loginUser=login(Request.Form("username"),Request.Form("password"))
         if not request.Cookies(CookieName)("memName")=username then
            	 ReInfo(0)="评论发表错误信息"
            	 ReInfo(1)="<b>登录失败，请检查用户名和密码</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
            	 ReInfo(2)="WarningIcon"
            	 postcomm=ReInfo
            	 exit function
         end if
    else
       set checkMem=Conn.ExeCute("select top 1 mem_id from blog_Member where mem_Name='"&username&"'")
       if not checkMem.eof then
    	 ReInfo(0)="评论发表错误信息"
    	 ReInfo(1)="<b>该用户已经存在，无法发表评论</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
    	 ReInfo(2)="WarningIcon"
    	 postcomm=ReInfo
    	 exit function
       end if
    end if
  end if 
  if not stat_CommentAdd then
	 ReInfo(0)="评论发表错误信息"
	 ReInfo(1)="<b>你没有权限发表评论</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
	 ReInfo(2)="ErrorIcon"
	 postcomm=ReInfo
	 exit function
  end if  
  if Conn.ExeCute("select log_DisComment from blog_Content where log_ID="&post_logID)(0) then 
	 ReInfo(0)="评论发表错误信息"
	 ReInfo(1)="<b>该日志不允许发表任何评论</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
	 ReInfo(2)="WarningIcon"
	 postcomm=ReInfo
	 exit function
  end if 
  post_DisSM=request.form("log_DisSM")
  post_DisURL=request.form("log_DisURL")
  post_DisKEY=request.form("log_DisKey")

  if len(post_Message)<1 then
	 ReInfo(0)="评论发表错误信息"
	 ReInfo(1)="<b>不允许发表空评论</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
	 ReInfo(2)="ErrorIcon"
	 postcomm=ReInfo
	 exit function
  end if
  if len(post_Message)>blog_commLength then
	 ReInfo(0)="评论发表错误信息"
	 ReInfo(1)="评论超过最大字数限制<br/><a href=""javascript:history.go(-1);"">单击返回</a>"
	 ReInfo(2)="ErrorIcon"
	 postcomm=ReInfo
	 exit function
  end if   
  'UBB 特别属性
  post_disImg=1
  post_DisUBB=0
  if post_DisSM=1 then post_DisSM=1 else post_DisSM=0
  if post_DisURL=1 then post_DisURL=0 else post_DisURL=1
  if post_DisKEY=1 then post_DisKEY=0 else post_DisKEY=1
 '插入数据
 Dim AddComm
 AddComm=array(array("blog_ID",post_logID),array("comm_Content",post_Message),array("comm_Author",username),array("comm_DisSM",post_DisSM),array("comm_DisUBB",post_DisUBB),array("comm_DisIMG",post_disImg),array("comm_AutoURL",post_DisURL),Array("comm_PostIP",getIP),Array("comm_AutoKEY",post_DisKEY))
 DBQuest "blog_Comment",AddComm,"insert"
 'Conn.ExeCute("INSERT INTO blog_Comment(blog_ID,comm_Content,comm_Author,comm_DisSM,comm_DisUBB,comm_DisIMG,comm_AutoURL,comm_PostIP,comm_AutoKEY) VALUES ("&post_logID&",'"&post_Message&"','"&username&"',"&post_DisSM&","&post_DisUBB&","&post_disImg&","&post_DisURL&",'"&getIP()&"',"&post_DisKEY&")")
 Conn.ExeCute("update blog_Content set log_CommNums=log_CommNums+1 where log_ID="&post_logID)
 Conn.ExeCute("update blog_Info set blog_CommNums=blog_CommNums+1")
 Response.Cookies(CookieName)("memLastpost")=Now()
 getInfo(2)
 NewComment(2)
 if memName<>empty then
  	conn.execute("update blog_Member set mem_PostComms=mem_PostComms+1 where mem_Name='"&memName&"'")
 end if
 SQLQueryNums=SQLQueryNums+3
 ReInfo(0)="评论发表成功"
 ReInfo(1)="<b>你成功地对该日志发表了评论</b><br/><a href=""default.asp?id="&post_logID&""">单击返回该日志</a>"
 ReInfo(2)="MessageIcon"
 Session("GetCode") = empty
 Session(CookieName&"_LastDo")="AddComment"
 postcomm=ReInfo
 PostArticle post_logID
end function
%>
  <br/></div> 
 </div>
<!--#include file="footer.asp" -->