<!--#include file="BlogCommon.asp" -->
<!--#include file="common/function.asp" -->
<!--#include file="common/library.asp" -->
<!--#include file="common/upfile.asp" -->
<!--#include file="common/cache.asp" -->
<!--#include file="common/checkUser.asp" -->
<!--#include file="class/cls_logAction.asp" -->
<%
response.expires=-1 
response.expiresabsolute=now()-1 
response.cachecontrol="no-cache"
Dim title, cname, Message, lArticle, postLog, SaveId
Dim cCateID, e_tags, ctype, logWeather, logLevel, logcomorder, logDisComment, logIsTop, logIsHidden, logMeta, logFrom, logFromURL, logdisImg, logDisSM, logDisURL, logDisKey, logQuote
'--------------Alias-----------------
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
'-------------[mdown]---------------
elseif request("action")="type1" then
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
'---------------[hidden]-----------------
elseif request("action")="Hidden" then
    If Len(memName)>0 Then
	   response.write "1"
	Else
	   response.Write "0"
	End If
'---------------[用户名检测]-----------------
elseif request("action")="checkname" then
	dim strname,checkdb
		strname = CheckStr(request("usename"))
		set checkdb = conn.execute("select * from blog_Member where mem_Name='"&strname&"'")
	if checkdb.bof or checkdb.eof then
		response.write"<font color=""#0000ff"">用户名未注册！</font>|$|True"
	else
		response.write"<font color=""#ff0000"">用户名已注册！</font>|$|False"
	end if
		set checkdb = nothing
'---------------------------[Ajax草稿保存 -- 发表时保存]--------------------------    
elseif request("action")="PostSave" then    
    If ChkPost() Then   
        title = CheckStr(Request.QueryString("title"))    
        cname = CheckStr(Request.QueryString("cname"))   
        Message = CheckStr(Request.QueryString("Message")) 
		cCateID = CheckStr(Request.QueryString("cateid"))
		e_tags =  CheckStr(Request.QueryString("tags"))
		ctype = CheckStr(Request.QueryString("ctype"))
		logWeather = CheckStr(Request.QueryString("logweather"))
		logLevel = CheckStr(Request.QueryString("logLevel"))
		logcomorder = CheckStr(Request.QueryString("logcomorder"))
		logDisComment = Request.QueryString("logDisComment")
		logIsTop = Request.QueryString("logIsTop")
		'logIsHidden = Request.QueryString("logIsHidden")
		logMeta = Request.QueryString("logMeta")
		logFrom = Request.QueryString("logFrom")
		logFromURL = Request.QueryString("logFromURL")
		logdisImg = Request.QueryString("logdisImg")
		logDisSM = Request.QueryString("logDisSM")
		logDisURL = Request.QueryString("logDisURL")
		logDisKey = Request.QueryString("logDisKey")
		logQuote = Request.QueryString("logQuote")
            
        Set lArticle = New logArticle    
        lArticle.logTitle = title    
        lArticle.logcname = cname 
		lArticle.logCtype = ctype   
        lArticle.logMessage = Message 
		lArticle.categoryID = cCateID   
        lArticle.logAuthor = memName ' 关键是这个
		lArticle.logTags = e_tags    
        lArticle.logIsDraft = CBool(true)
		lArticle.isajax = true
		lArticle.logWeather = logWeather 
		lArticle.logLevel = logLevel
		lArticle.logCommentOrder = logcomorder  
		lArticle.logDisableComment = logDisComment
		lArticle.logIsTop = logIsTop
		lArticle.logMeta = logMeta
		lArticle.logFrom = logFrom
    	lArticle.logFromURL = logFromURL
		lArticle.logDisableImage = logdisImg
		lArticle.logDisableSmile = logDisSM
		lArticle.logDisableURL = logDisURL
		lArticle.logDisableKeyWord = logDisKey
		lArticle.logTrackback = logQuote
        postLog = lArticle.postLog    
        Set lArticle = Nothing   
            
        response.write "草稿保存成功！|$|1|$|"&postLog(2)    
    else    
        response.write "非法提交数据"   
    end if    
'---------------------------[Ajax草稿保存 -- 编辑时保存]--------------------------    
elseif request("action")="UpdateSave" then    
    If ChkPost() Then   
        title = CheckStr(Request.QueryString("title"))    
        cname = CheckStr(Request.QueryString("cname"))   
        Message = CheckStr(Request.QueryString("Message")) 
		e_tags =  CheckStr(Request.QueryString("tags")) 
		ctype = CheckStr(Request.QueryString("ctype"))
		logWeather = CheckStr(Request.QueryString("logweather")) 
		logLevel = CheckStr(Request.QueryString("logLevel"))
		logcomorder = CheckStr(Request.QueryString("logcomorder"))
		logDisComment = Request.QueryString("logDisComment")
		logIsTop = Request.QueryString("logIsTop")
		'logIsHidden = Request.QueryString("logIsHidden")
		logMeta = Request.QueryString("logMeta")
		logFrom = Request.QueryString("logFrom")
		logFromURL = Request.QueryString("logFromURL")
		logdisImg = Request.QueryString("logdisImg")
		logDisSM = Request.QueryString("logDisSM")
		logDisURL = Request.QueryString("logDisURL")
		logDisKey = Request.QueryString("logDisKey")
		logQuote = Request.QueryString("logQuote")
            
        SaveId = Request.QueryString("postId")    
            
        Set lArticle = New logArticle    
        lArticle.logTitle = title    
        lArticle.logcname = cname 
		lArticle.logCtype = ctype   
        lArticle.logMessage = Message    
        lArticle.logAuthor = memName ' 关键是这个
		lArticle.logTags = e_tags
		lArticle.logIsDraft = CBool(true)
		lArticle.isajax = true   
		lArticle.logWeather = logWeather 
		lArticle.logLevel = logLevel
		lArticle.logCommentOrder = logcomorder
		lArticle.logDisableComment = logDisComment
		lArticle.logIsTop = logIsTop
		lArticle.logMeta = logMeta
		lArticle.logFrom = logFrom
    	lArticle.logFromURL = logFromURL
		lArticle.logDisableImage = logdisImg
		lArticle.logDisableSmile = logDisSM
		lArticle.logDisableURL = logDisURL
		lArticle.logDisableKeyWord = logDisKey
		lArticle.logTrackback = logQuote
        postLog = lArticle.editLog(SaveId)    
        Set lArticle = Nothing   
            
        response.write "草稿更新成功！|$|0|$|"&SaveId    
    else    
        response.write "非法提交数据"   
    end if
else
response.write "非法操作!"
End If
%>
