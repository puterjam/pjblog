<!--#include file="BlogCommon.asp" -->
<!--#include file="common/function.asp" -->
<!--#include file="common/library.asp" -->
<!--#include file="common/upfile.asp" -->
<!--#include file="common/cache.asp" -->
<!--#include file="common/checkUser.asp" -->
<!--#include file="class/cls_logAction.asp" -->
<!--#include file="class/cls_article.asp" -->
<!--#include file="common/ubbcode.asp" -->
<!--#include file="control/f_control.asp" -->
<%
response.expires=-1
response.expiresabsolute=now()-1
response.cachecontrol="no-cache"
'*************************************************
' Ajax 类 ASP 代码 // AjaxRequest.js 框架支持
' evio edit 
'*************************************************
Dim title, cname, Message, lArticle, postLog, SaveId, ReadForm, SplitReadForm
Dim cCateID, e_tags, ctype, logWeather, logLevel, logcomorder, logDisComment, logIsTop, logIsHidden, logMeta, logFrom, logFromURL, logdisImg, logDisSM, logDisURL, logDisKey, logQuote
Dim NewCateName, AType, BlogId, eSql
'-------------- [Alias] -----------------
If request.QueryString("action") = "checkAlias" then
	If ChkPost() Then
   		dim strcname, checkcdb
   		strcname = Checkxss(request.QueryString("cname"))
   		set checkcdb = conn.execute("select * from [blog_Content] where [log_cname]="""&strcname&"""")
       		if checkcdb.bof or checkcdb.eof then
          		response.write "<img src=""images/check_right.gif"">"
       		else
          		response.write "<img src=""images/check_error.gif"">"
       		end if
    	set checkcdb=nothing
	End If
'------------- [mdown] ---------------
elseif request.QueryString("action") = "type1" then
    If ChkPost() Then
		dim mainurl, main, mainstr
        	mainurl = Checkxss(request.QueryString("mainurl"))
        	main = trim(Checkxss(request.QueryString("main")))
        	response.clear()
        	mainstr = ""
    	If Len(memName)>0 Then
        	mainstr = mainstr & "<img src=""images/download.gif"" alt=""下载文件"" style=""margin:0px 2px -4px 0px""/> <a href="""&mainurl&""" target=""_blank"">"&main&"</a>"
    	Else
        	mainstr = mainstr & "<img src=""images/download.gif"" alt=""只允许会员下载"" style=""margin:0px 2px -4px 0px""/> 该文件只允许会员下载! <a href=""login.asp"">登录</a> | <a href=""register.asp"">注册</a>"
    	End If
    	response.write mainstr
	End If
elseif request.QueryString("action") = "type2" then
    If ChkPost() Then
		dim main2, mstr
        	main2 = Checkxss(request.QueryString("main"))
        	response.clear()
        	mstr = ""
    	If Len(memName) > 0 Then
       		mstr=mstr&"<img src=""images/download.gif"" alt=""下载文件"" style=""margin:0px 2px -4px 0px""/> <a href="""&main2&""" target=""_blank"">下载此文件</a>"
    	Else
       		mstr=mstr&"<img src=""images/download.gif"" alt=""只允许会员下载"" style=""margin:0px 2px -4px 0px""/> 该文件只允许会员下载! <a href=""login.asp"">登录</a> | <a href=""register.asp"">注册</a>"
    	End If
    	response.write mstr
	End If
'--------------- [hidden] -----------------
elseif request.QueryString("action") = "Hidden" then
    If Len(memName) > 0 Then
	   response.write "1"
	Else
	   response.Write "0"
	End If
'--------------- [用户名检测] -----------------
elseif request.QueryString("action") = "checkname" then
	If ChkPost() Then
		dim strname, checkdb
			strname = Checkxss(request.QueryString("usename"))
			set checkdb = conn.execute("select * from blog_Member where mem_Name='"&strname&"'")
				if checkdb.bof or checkdb.eof then
					response.write"<font color=""#0000ff"">用户名未注册！</font>|$|True"
				else
					response.write"<font color=""#ff0000"">用户名已注册！</font>|$|False"
				end if
			set checkdb = nothing
	End If
'--------------------------- [Ajax草稿保存 -- 发表时保存] --------------------------    
elseif request.QueryString("action") = "PostSave" then    
    If ChkPost() Then
		ReadForm = Request.Form
		SplitReadForm = Split(ReadForm, "|$|")
'		response.Write(Ubound(SplitReadForm))
'		response.End()
		'str = escape(title) + "|$|" + escape(cname) + "|$|" + escape(ctype) + "|$|" + escape(logweather) + "|$|" + escape(logLevel) + "|$|" + escape(logcomorder) + "|$|" + escape(logDisComment) + "|$|" + escape(logIsTop) + "|$|" + escape(logMeta) + "|$|" + escape(logFrom) + "|$|" + escape(logFromURL) + "|$|" + escape(logdisImg) + "|$|" + escape(logDisSM) + "|$|" + escape(logDisURL) + "|$|" + escape(logDisKey) + "|$|" + escape(logQuote) + "|$|" + escape(Tags) + "|$|" + escape(mCateID) + "|$|" + escape(Message)
		'            0                      1                        2                         3                         4                              5                6                             7                         8                         9                             10                        11              12                        13                        14                              15                       16                     17        18
        title = UnEscape(SplitReadForm(0)) 
        cname = UnEscape(SplitReadForm(1))
		ctype = UnEscape(SplitReadForm(2))
		logWeather = UnEscape(SplitReadForm(3))
		logLevel = UnEscape(SplitReadForm(4))
		logcomorder = UnEscape(SplitReadForm(5))
		logDisComment = UnEscape(SplitReadForm(6))
		logIsTop = UnEscape(SplitReadForm(7))
		'logIsHidden = Request.QueryString("logIsHidden")
		logMeta = UnEscape(SplitReadForm(8))
		logFrom = UnEscape(SplitReadForm(9))
		logFromURL = UnEscape(SplitReadForm(10))
		logdisImg = UnEscape(SplitReadForm(11))
		logDisSM = UnEscape(SplitReadForm(12))
		logDisURL = UnEscape(SplitReadForm(13))
		logDisKey = UnEscape(SplitReadForm(14))
		logQuote = UnEscape(SplitReadForm(15))
		e_tags =  UnEscape(SplitReadForm(16))
		cCateID = UnEscape(SplitReadForm(17))
		Message = UnEscape(SplitReadForm(18))

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
            
        response.write "草稿于 " & DateToStr(now(), "Y-m-d H:I:S") & " <font color='#9C0024'>保存</font>成功,请不要刷新本页,以免丢失信息!|$|1|$|"&postLog(2)    
    else    
        response.write "非法提交数据"   
    end if    
'--------------------------- [Ajax草稿保存 -- 编辑时保存] --------------------------    
elseif request.QueryString("action") = "UpdateSave" then    
    If ChkPost() Then 
		ReadForm = Request.Form
		SplitReadForm = Split(ReadForm, "|$|")
'		response.Write(Ubound(SplitReadForm))
'		response.End()
		'str = escape(title) + "|$|" + escape(cname) + "|$|" + escape(ctype) + "|$|" + escape(logweather) + "|$|" + escape(logLevel) + "|$|" + escape(logcomorder) + "|$|" + escape(logDisComment) + "|$|" + escape(logIsTop) + "|$|" + escape(logMeta) + "|$|" + escape(logFrom) + "|$|" + escape(logFromURL) + "|$|" + escape(logdisImg) + "|$|" + escape(logDisSM) + "|$|" + escape(logDisURL) + "|$|" + escape(logDisKey) + "|$|" + escape(logQuote) + "|$|" + escape(Tags) + "|$|" + escape(mCateID) + "|$|" + escape(Message) + "|$|" + escape(id)
		'            0                      1                        2                         3                         4                              5                6                             7                         8                         9                             10                        11              12                        13                        14                              15                       16                     17      18 
        title = UnEscape(SplitReadForm(0))    
        cname = UnEscape(SplitReadForm(1))   
        Message = UnEscape(SplitReadForm(18)) 
		e_tags =  UnEscape(SplitReadForm(16)) 
		ctype = UnEscape(SplitReadForm(2))
		logWeather = UnEscape(SplitReadForm(3)) 
		logLevel = UnEscape(SplitReadForm(4))
		logcomorder = UnEscape(SplitReadForm(5))
		logDisComment = UnEscape(SplitReadForm(6))
		logIsTop = UnEscape(SplitReadForm(7))
		'logIsHidden = Request.QueryString("logIsHidden")
		logMeta = UnEscape(SplitReadForm(8))
		logFrom = UnEscape(SplitReadForm(9))
		logFromURL = UnEscape(SplitReadForm(10))
		logdisImg = UnEscape(SplitReadForm(11))
		logDisSM = UnEscape(SplitReadForm(12))
		logDisURL = UnEscape(SplitReadForm(13))
		logDisKey = UnEscape(SplitReadForm(14))
		logQuote = UnEscape(SplitReadForm(15))
        cCateID = UnEscape(SplitReadForm(17))
		
        SaveId = UnEscape(SplitReadForm(19))   
            
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
		lArticle.categoryID = cCateID
        postLog = lArticle.editLog(SaveId)    
        Set lArticle = Nothing   
            
        response.write "草稿于 " & DateToStr(now(), "Y-m-d H:I:S") & " <font color='#AF5500'>更新</font>成功,请不要刷新本页,以免丢失信息!|$|0|$|"&SaveId    
    else    
        response.write "非法提交数据"   
    end if
'-------------[防盗链]---------------
ElseIf request("action") = "Antidown" or request("action") = "Antimdown" then
	If ChkPost() Then
		Dim down, showdownstr, id
			Session(CookieName & "Antimdown") = "pjblog_Antimdown"
			id = Checkxss(request("id"))
			if instr(id, "&") > 0 then id = split(id,"&")(0)
			Set down = conn.execute("select FilesPath,FilesCounts from blog_Files where id="&id)
				response.clear()
				If request("action") = "Antimdown" and memName = empty Then
					showdownstr = getFileIcons(getFileInfo(down(0))(9))&" 该文件只允许会员下载! <a href=""login.asp"" accesskey=""L"">登录</a> | <a href=""register.asp"">注册</a>"
				Else
					showdownstr = getFileIcons(getFileInfo(down(0))(9))&" <a href="""&request("downurl")&""" target=""_blank"">"&trim(checkstr(request("main")))&"</a><font color=""#999999"">("&getFileInfo(down(0))(0)&")</font><br/><font color=""#999999"">["&Datetostr(getFileInfo(down(0))(10),"Y-m-d H:I A")&"; 下载次数:"&down(1)&"]</font>"
				End If
			response.write showdownstr
	else
		response.write "非法提交数据"
	end if
'-------------[Ajax增加新分类]---------------
ElseIf Request.QueryString("action") = "AddNewCate" then
	If ChkPost() Then
		Dim AddNewCateArray, IntCount, NewcatePart, NRs, ReturnID
		
		NewcatePart = "newadd_" & randomStr(8)
		IntCount = (conn.Execute("select count(*) from blog_Category")(0) + 1)
		NewCateName = Checkxss(Request.QueryString("newcate"))
		
		'*********************************
		Set NRs = Server.CreateObject("Adodb.Recordset")
		NRs.open "blog_Category", Conn, 1, 2
			NRs.AddNew
			ReturnID = NRs("cate_ID")
			NRs("cate_Name") = NewCateName
			NRs("cate_Order") = Int(IntCount)
			NRs("cate_Intro") = NewCateName
			NRs("cate_OutLink") = False
			NRs("cate_URL") = ""
			NRs("cate_local") = 0
			NRs("cate_icon") = "images/icons/1.gif"
			NRs("cate_Secret") = 0
			NRs("cate_Part") = NewcatePart
			NRs.update
		NRs.Close
		Set NRs = Nothing
		CategoryList(2)
		'*********************************
		
		Response.Write(ReturnID)
	Else
		response.write "非法提交数据"
	End If
'-------------[Ajax找回密码]---------------
ElseIf Request.QueryString("action") = "GetPassReturnInfo" Then
	If ChkPost() Then
		Dim Password, PName, Rs
		Password = CheckStr(UnEscape(Request.QueryString("password")))
		PName = CheckStr(UnEscape(Request.QueryString("name")))
		Set Rs = Conn.Execute("Select * From [blog_Member] Where [mem_Name]='"&PName&"'")
		If Rs.Bof Or Rs.Eof Then
			Response.Write("none")
		Else
			If Password = Trim(Rs("mem_Answer")) Then
				Response.Write(Rs("mem_ID"))
			Else
				Response.Write("wrong")
			End If
		End If
		Set Rs = Nothing
	Else
		response.write "非法提交数据"
	End If
ElseIf Request.QueryString("action") = "UpdatePass" Then
	If ChkPost() Then
		Dim u_ID, u_q, u_a
		u_ID = CheckStr(UnEscape(Request.QueryString("id")))
		u_q = CheckStr(UnEscape(Request.QueryString("q")))
		u_a = CheckStr(UnEscape(Request.QueryString("a")))
		Conn.Execute("Update [blog_Member] Set [mem_Question]='"&u_q&"',[mem_Answer]='"&u_a&"' Where [mem_ID]="&u_ID)
		Response.Write("1")
	Else
		response.write "非法提交数据"
	End If
ElseIf Request.QueryString("action") = "CheckPostName" Then
	If ChkPost() Then
		Dim zName, Zs
		zName = CheckStr(UnEscape(Request.QueryString("name")))
		Set Zs = Conn.Execute("Select * From [blog_Member] Where [mem_Name]='"&zName&"'")
		If Zs.Bof Or Zs.Eof Then
			Response.write "0"
		Else
			Response.Write(UnCheckStr(zName) & "|$|" & Zs("mem_Question"))
		End If
	Else
		response.write "非法提交数据"
	End If
ElseIf Request.QueryString("action") = "updatepassto" Then
	If ChkPost() Then
		Dim e_Pass, e_RePass, e_ID, e_Rs, e_hash, d_pass
		e_ID = CheckStr(UnEscape(Request.QueryString("id")))
		e_Pass = CheckStr(UnEscape(Request.QueryString("pass")))
		e_RePass = CheckStr(UnEscape(Request.QueryString("repass")))
		Set e_Rs = Server.CreateObject("Adodb.Recordset")
		e_Rs.open "Select * From [blog_Member] Where [mem_ID]="&e_ID, Conn, 1, 3
			e_hash = e_Rs("mem_salt")
			d_pass = SHA1(e_Pass&e_hash)
			e_Rs("mem_Password") = d_pass
			e_Rs.update
		e_Rs.Close
		Set e_Rs = nothing
		response.Write("1")
	Else
		response.write "非法提交数据"
	End If
'-------------[Ajax首页评论审核]---------------
ElseIf Request.QueryString("action") = "IndexAudit" Then
	AType = Int(CheckStr(UnEscape(Request.QueryString("type"))))
	SaveId = Int(CheckStr(UnEscape(Request.QueryString("id"))))
	BlogId = Int(CheckStr(UnEscape(Request.QueryString("blogid"))))
	If len(SaveId) = 0 Then Response.Write("评论参数错误") : Response.End()
	If len(BlogId) = 0 Then Response.Write("日志参数错误") : Response.End()
	If AType = 0 Then
		Conn.Execute("Update [blog_Comment] Set comm_IsAudit=true Where [comm_ID]="&SaveId)
		Response.Write("0")
	ElseIf AType = 1 Then
		Conn.Execute("Update [blog_Comment] Set comm_IsAudit=false Where [comm_ID]="&SaveId)
		Response.Write("1")
	Else
		Response.Write("9")
	End If
	NewComment(2)
	If blog_postFile = 2 Then PostArticle BlogId, False
'-------------[Ajax评论加载]---------------
ElseIf Request.QueryString("action") = "ReadArticleComentByCommentID" Then
	SaveId = CheckStr(UnEscape(Request.QueryString("CommentIDStr")))
	Dim SplitSaveId, SplitSi, SplitRs, SplitStr
	SplitSaveId = Split(SaveId, "|$|")
	SplitStr = ""
	For SplitSi = 0 to (Ubound(SplitSaveId) - 1)
		Set SplitRs = Conn.Execute("Select comm_ID,comm_Content,comm_Author,comm_PostTime,comm_DisSM,comm_DisUBB,comm_DisIMG,comm_AutoURL,comm_PostIP,comm_AutoKEY,comm_IsAudit From blog_Comment Where comm_ID="&SplitSaveId(SplitSi))
'   0         1            2           3            4           5         6             7           8             9         10
		If Not SplitRs.Eof And Not SplitRs.Bof Then
			If blog_AuditOpen Then
				If stat_Admin Then
					If SplitRs(10) Then
						SplitStr = SplitStr & Escape(UBBCode(HtmlEncode(SplitRs(1)),SplitRs(4),blog_commUBB,blog_commIMG,SplitRs(7),SplitRs(9))) & "|$|"
					Else
						SplitStr = SplitStr & Escape("[未审核评论,仅管理员可见]:&nbsp;" & UBBCode(HtmlEncode(SplitRs(1)),SplitRs(4),blog_commUBB,blog_commIMG,SplitRs(7),SplitRs(9))) & "|$|"
					End If
				Else
					If SplitRs(10) Then	
						SplitStr = SplitStr & Escape(UBBCode(HtmlEncode(SplitRs(1)),SplitRs(4),blog_commUBB,blog_commIMG,SplitRs(7),SplitRs(9))) & "|$|"
					Else
						SplitStr = SplitStr & "[未审核评论,仅管理员可见]|$|"
					End If
				End If
			Else
				SplitStr = SplitStr & Escape(UBBCode(HtmlEncode(SplitRs(1)),SplitRs(4),blog_commUBB,blog_commIMG,SplitRs(7),SplitRs(9))) & "|$|"
			End If
		End If
		Set SplitRs = nothing
	Next
	Response.Write(SplitStr & "end")
Else
	response.write "非法操作!"
End If
%>
