<%
'==================================
'  日志类文件
'    更新时间: 2006-6-2
'==================================

'--------------------输出文件头------------------
sub outCardHead(title) ' Output WML Head
 response.write "<head><meta forua=""true"" http-equiv=""Cache-Control"" content=""max-age=0"" /></head>"
 response.write "<card id=""MainCard"" title="""&toUnicode(title)&""">"
 if not blog_wap then 
  response.write "<p>"&toUnicode("Blog不允许通过wap访问.")&"</p>"
  response.write "</card>" 
  response.write "</wml>" 
  response.end 
 end if
end sub

'--------------------输出wap 首页----------------
sub outDefault 'Output Default
  outCardHead ("欢迎光临")
  response.write "<p><a href=""wap.asp"">"&toUnicode(HTMLDecode(SiteName))&"</a><br/>&nbsp;</p>"
  response.write "<p><a href=""wap.asp?do=showLog"">"&toUnicode("最新日志")&"</a></p>"
  response.write "<p><a href=""wap.asp?do=showComment"">"&toUnicode("最新评论")&"</a></p>"
  dim Arr_Category,Category_Len,i
  CategoryList(1)
  Category_code=""
  Arr_Category=Application(CookieName&"_blog_Category")
  if ubound(Arr_Category,1)=0 then exit Sub
  Category_Len=ubound(Arr_Category,2)
			For i=0 to Category_Len
				 if not Arr_Category(4,i) then
					  if cbool(Arr_Category(10,i)) then
						   if stat_ShowHiddenCate or stat_Admin then Category_code=Category_code&("&nbsp;- <a href=""wap.asp?do=showLog&amp;cateID="&Arr_Category(0,i)&""">"&toUnicode(Arr_Category(1,i))&"</a>&nbsp;("&Arr_Category(7,i)&")<br/>")
					   else
						   Category_code=Category_code&("&nbsp;- <a href=""wap.asp?do=showLog&amp;cateID="&Arr_Category(0,i)&""">"&toUnicode(Arr_Category(1,i))&"</a>&nbsp;("&Arr_Category(7,i)&")<br/>")
					  end if
				 end if
			Next
  response.write "<p>"&Category_code&"</p>"
  response.write "<p><a href=""wap.asp?do=showTag"">"&toUnicode("标签云集")&"</a></p>"
  outControl
  outCardFoot
  Conn.Close
  Set Conn=Nothing
end sub

'--------------------输出最新文章----------------
sub outLog
  outCardHead ("欢迎光临")
  dim wPageSize
  wPageSize=blog_wapNum
  response.write "<p><a href=""wap.asp"">"&toUnicode(HTMLDecode(SiteName))&"</a><br/>&nbsp;</p>"
  Dim dbRow,i,webLog,strSQL,CanRead,Log_Num
		Dim CT,ViewTag,getCate
		   CT="最新日志"
		   ViewTag=checkstr(Request.QueryString("tag"))
		   set getCate=new Category
		   IF IsInteger(cateID)=True Then
			 getCate.load(cateID)
			 CT="分类: "&getCate.cate_Name
		 	 if getCate.cate_Secret then
				   if not stat_ShowHiddenCate and not stat_Admin then  
					  response.write "<p>"&toUnicode("找不到任何日志")&"</p>"
					  outCardFoot
				   exit sub
			 	 end if
		 	  end if
	    End If
	    
		If len(ViewTag)>0 then
			dim getTag,getTID
			set getTag=new tag
			getTID=getTag.getTagID(ViewTag)
			if getTID<>0 then 
			   SQLFiltrate=SQLFiltrate & " log_tag LIKE '%{"&getTID&"}%' AND "
			   Url_Add=Url_Add & "tag="&Server.URLEncode(ViewTag)&"&amp;"
			   CT="Tag: "&ViewTag
			end if
			set getTag=nothing
		end if	  
	  strSQL="log_ID,log_Title,log_CateID,log_Author,log_PostTime,log_IsShow,log_IsTop"
	  if stat_ShowHiddenCate or stat_Admin then
		SQL="SELECT "&strSQL&" FROM blog_Content "&SQLFiltrate&" log_IsDraft=false ORDER BY log_PostTime DESC"
	   else
		SQL="SELECT "&strSQL&" FROM blog_Content As T,blog_Category As C "&SQLFiltrate&" T.log_CateID=C.cate_ID and C.cate_Secret=false and log_IsDraft=false ORDER BY log_PostTime DESC"
	  end if
	
	Set webLog=Server.CreateObject("Adodb.Recordset")
	webLog.Open SQL,Conn,1,1
	SQLQueryNums=SQLQueryNums+1

	if webLog.EOF or webLog.BOF then
		  response.write "<p>"&toUnicode("找不到任何日志")&"</p>"
		  outCardFoot
		  ReDim dbRow(0,0)
		  exit sub
	 else
		webLog.PageSize=wPageSize
		webLog.AbsolutePage=CurPage
		Log_Num=webLog.RecordCount
		dbRow=webLog.getrows(wPageSize)
	end if
	webLog.close
	set webLog=nothing
	Conn.Close
	Set Conn=Nothing
	
	if ubound(dbRow,1)<>0 then
		  response.write "<p>"&toUnicode(CT)&"</p>"
		for i=0 to ubound(dbRow,2)
		  CanRead=false
		  If stat_Admin=True then CanRead=True
		  If dbRow(5,i) Then CanRead=True
		  If dbRow(5,i)=False And dbRow(2,i)=memName then CanRead=True
		  if CanRead then 
			  response.write "<p>&nbsp;&nbsp;"&(i+wPageSize*(CurPage-1)+1)&". <a href=""wap.asp?do=showLogDetail&amp;id="&dbRow(0,i)&""">"&toUnicode(dbRow(1,i))&"</a>"
			  if dbRow(5,i)=False then response.write toUnicode(" [隐]")
			  response.write "</p>"
			else
			  response.write "<p>&nbsp;&nbsp;"&(i+wPageSize*(CurPage-1)+1)&". "&toUnicode("[隐藏日志]")&"</p>"
	      end if
		next
		call getPage (Log_Num,CurPage,wPageSize,Url_Add)
	end if
  
  outControl
  outCardFoot
end sub

sub outLogDetail
  outCardHead ("欢迎光临")
   response.write "<p><a href=""wap.asp"">"&toUnicode(HTMLDecode(SiteName))&"</a><br/>&nbsp;</p>"
 dim logID,lArticle,loadArticle
  logID=clng(request.QueryString("id"))
  
  set lArticle = new logArticle
    loadArticle=lArticle.getLog(logID)
    if loadArticle(0)=0 then 
      Dim getCate
      set getCate=new Category
	  getCate.load(int(lArticle.categoryID))  
         If (lArticle.logAuthor=memName And lArticle.logIsShow=False) Or stat_Admin or lArticle.logIsShow=true then 
		else
	 		response.write "<p>"&toUnicode("没有权限查看该日志.")&"</p>"
			outCardFoot
		end if
		
		If (Not getCate.cate_Secret) Or (lArticle.logAuthor=memName And getCate.cate_Secret) Or stat_Admin Or (getCate.cate_Secret and stat_ShowHiddenCate) Then
		else
	 		response.write "<p>"&toUnicode("没有权限查看该日志.")&"</p>"
			outCardFoot
		end if
		
		response.write "<p>"&"<b>"&toUnicode("标题:")&"</b> "&toUnicode(lArticle.logTitle)
		if lArticle.logIsShow=False then response.write toUnicode(" [隐]")
		response.write "</p>"
 		response.write "<p>"&"<b>"&toUnicode("作者:")&"</b> "&toUnicode(lArticle.logAuthor)&"</p>"
 		response.write "<p>"&"<b>"&toUnicode("日期:")&"</b> "&toUnicode(DateToStr(lArticle.logPubTime,"Y-m-d H:I A"))&"</p>"
 		response.write "<p>"&"<b>"&toUnicode("分类:")&"</b> <a href=""wap.asp?do=showLog&amp;cateID="&lArticle.categoryID&""">"&toUnicode(getCate.cate_Name)&"</a></p>"
 		if lArticle.logEditType=0 then 
	 		response.write "<p>"&"<b>"&toUnicode("内容:")&"</b> "&HtmlToText(UnCheckStr(lArticle.logMessage))&"</p>"
        else
	 		response.write "<p>"&"<b>"&toUnicode("内容:")&"</b> "&HtmlToText(UBBCode(HTMLEncode(lArticle.logMessage),0,0,0,1,1))&"</p>"
        end if
  		response.write "<p> + <a href=""#CommentCard"">"&toUnicode("查看当前日志评论")&"</a> ("&lArticle.logCommentCount&")</p>"
    else
  		response.write "<p>"&toUnicode(loadArticle(1))&"</p><p><a href=""wap.asp"">"&toUnicode("返回首页")&"</a></p>"
		outCardFoot
		exit sub
  end if
  outControl
  outCardFoot
  
  if stat_CommentAdd then
	  response.write "<card id=""postCommentCard"">"
	  response.write "<p><b>"&toUnicode("标题:")&"</b> <a href=""#MainCard"">"&toUnicode(lArticle.logTitle)&"</a><br/></p>"
	  response.write "<p><br/><b>发表评论:</b></p><p>"
	  if memName<>Empty then 
		  response.write toUnicode("昵称: ")&"<b><i>"&memName&"</i></b><br/>"
	   else
		  response.write toUnicode("昵称: ")&"<input emptyok=""false"" name=""userName"" size=""10"" maxlength=""20"" type=""text"" value="""" /><br/>"
	  end if
	  response.write toUnicode("评论内容: ")&"<br/> <input emptyok=""false"" name=""message""  maxlength="""&blog_commLength&""" format=""false"" value="""" /><br/>"
	  response.write "<anchor>"&toUnicode("发表")&"<go href=""wap.asp?do=postComment"" method=""post"">"
	  if memName<>Empty then 
		  response.write "<postfield name=""userName"" value="""&memName&""" />"
	  else
		  response.write "<postfield name=""userName"" value=""$(userName)"" />"
	  end if
	  response.write "<postfield name=""message"" value=""$(message)"" />"
	  response.write "<postfield name=""id"" value="""&logID&""" />"
	  response.write "</go></anchor>"
	  response.write "&nbsp;<a href=""#MainCard"">"&toUnicode("返回")&"</a></p>"
   else
	  response.write "<card id=""postCommentCard"">"
	  response.write "<p><b>"&toUnicode("标题:")&"</b> <a href=""#MainCard"">"&toUnicode(lArticle.logTitle)&"</a></p>"
	  response.write "<p><br/>你没有权限发表评论</p>"
  end if
  outCardFoot
    
  response.write "<card id=""CommentCard"">"
  dim blog_Comment,comDesc,dbRow,i
  if lArticle.logCommentOrder then comDesc="Desc" else comDesc="Asc" end If
  SQL="SELECT comm_ID,comm_Content,comm_Author,comm_PostTime,comm_DisSM,comm_DisUBB,comm_DisIMG,comm_AutoURL,comm_PostIP,comm_AutoKEY FROM blog_Comment WHERE blog_ID="&logID&" ORDER BY comm_PostTime "&comDesc
  set blog_Comment=conn.execute (SQL)
  IF blog_Comment.EOF AND blog_Comment.BOF Then
  		response.write "<p>"&toUnicode("暂无评论")&"</p><p><a href=""#MainCard"">"&toUnicode("返回")&"</a></p>"
	    outCardFoot
	    exit sub
  	else
		dbRow=blog_Comment.getrows()
  end if
  response.write "<p><b>"&toUnicode("标题:")&"</b> <a href=""#MainCard"">"&toUnicode(lArticle.logTitle)&"</a></p>"
  response.write"<p><b>"&toUnicode("评论内容:")&"</b></p>"
  
	if ubound(dbRow,1)<>0 then
		for i=0 to ubound(dbRow,2)
			  response.write"<p>== <b>"&toUnicode(dbRow(2,i))&"</b> <small>"&DateToStr(dbRow(3,i),"Y-m-d H:I A")&"</small> ----</p>"
			  response.write"<p>"&HtmlToText(UBBCode(HtmlEncode(dbRow(1,i)),dbRow(4,i),dbRow(5,i),dbRow(6,i),dbRow(7,i),dbRow(9,i)))&" </p>"
			  response.write"<p> </p>"
		next
	end if  
  outCardFoot
 
  blog_Comment.close
  set blog_Comment=nothing
  set lArticle = nothing
  Conn.Close
  Set Conn=Nothing
end sub

'--------------------输出导航----------------
sub getPage (lN,cP,wP,uA)
		dim getPages,URL
		If lN Mod Cint(wP)=0 Then
			getPages=Int(lN/wP)
		Else
			getPages=Int(lN/wP)+1
		End If
		URL=Request.ServerVariables("Script_Name")&uA
			  response.write "<p> "
			  if int(cP)<>1 then 
				 response.write "<a href="""&Url&"do=showLog&amp;page="&(cP-1)&""">"& toUnicode("上一页") &"</a>" & " "
			   else
			     response.write toUnicode("上一页") & " "
			  end if
			  
			  if getPages<>int(cP) then 
				   response.write "<a href="""&Url&"do=showLog&amp;page="&(cP+1)&""">"& toUnicode("下一页") &"</a>"
			   else
				   response.write toUnicode("下一页") 
			  end if
			  response.write "</p>"
end sub

'--------------------输出最新评论----------------
sub outComment
  outCardHead ("欢迎光临")
  response.write "<p><a href=""wap.asp"">"&toUnicode(HTMLDecode(SiteName))&"</a><br/>&nbsp;</p>"
  dim blog_Comment,comDesc,dbRow,i
  SQL="SELECT top 10 comm_ID,comm_Content,comm_Author,comm_PostTime,comm_DisSM,comm_DisUBB,comm_DisIMG,comm_AutoURL,comm_AutoKEY,blog_ID FROM blog_Comment ORDER BY comm_PostTime desc"
  set blog_Comment=conn.execute (SQL)
  IF blog_Comment.EOF AND blog_Comment.BOF Then
  		response.write "<p>"&toUnicode("暂无评论")&"</p><p><a href=""wap.asp"">"&toUnicode("返回首页")&"</a></p>"
	    outCardFoot
	    exit sub
  	else
		dbRow=blog_Comment.getrows()
  end if
  blog_Comment.close
  set blog_Comment=nothing
  Conn.Close
  Set Conn=Nothing
  
	if ubound(dbRow,1)<>0 then
		for i=0 to ubound(dbRow,2)
			  response.write"<p>== <b>"&toUnicode(dbRow(2,i))&"</b> <small>"&DateToStr(dbRow(3,i),"Y-m-d H:I A")&"</small> ----</p>"
			  response.write"<p>"&HtmlToText(UBBCode(HtmlEncode(dbRow(1,i)),dbRow(4,i),dbRow(5,i),dbRow(6,i),dbRow(7,i),dbRow(8,i)))&" </p>"
			  response.write"<p> + <a href=""wap.asp?do=showLogDetail&amp;id="&dbRow(9,i)&"#CommentCard"">"&toUnicode("查看完整评论")&"</a></p>"
			  response.write"<p> </p>"
		next
	end if  
  outCardFoot
end sub

'--------------------输出Tag列表----------------
sub outTagList
  outCardHead ("标签云集")
  response.write "<p><a href=""wap.asp"">"&toUnicode(HTMLDecode(SiteName))&"</a><br/>&nbsp;</p><p>"
  dim log_Tag,log_TagItem
	For Each log_TagItem IN Arr_Tags
		log_Tag=Split(log_TagItem,"||")
			%>
	          &nbsp;- <a href="wap.asp?do=showLog&amp;tag=<%=Server.URLEncode(log_Tag(1))%>"><%=toUnicode(log_Tag(1))%></a>&nbsp;(<%=log_Tag(2)%>)<br/>
			<%
 	Next
  response.write "</p><p><br/><a href=""wap.asp"">"&toUnicode("返回首页")&"</a></p>"
  outControl
  outCardFoot
  Conn.Close
  Set Conn=Nothing
end sub

'--------------------登录框----------------
sub outLogin
  outCardHead ("登入Blog")
  response.write "<p><a href=""wap.asp"">"&toUnicode(HTMLDecode(SiteName))&"</a><br/>&nbsp;</p><p>"
  response.write toUnicode("用户名: ")&"<input emptyok=""false"" name=""userName"" size=""10"" maxlength=""20"" type=""text"" value="""" /><br/>"
  response.write toUnicode("密　码: ")&"<input emptyok=""false"" name=""Password"" size=""10"" maxlength=""32"" type=""password"" /><br/>"
  response.write "<anchor>"&toUnicode("登入")&"<go href=""wap.asp?do=CheckUser"" method=""post"">"
  response.write "<postfield name=""userName"" value=""$(userName)"" />"
  response.write "<postfield name=""Password"" value=""$(Password)"" />"
  response.write "<postfield name=""validate"" value=""0000"" />"
  response.write "</go></anchor>"
  response.write "&nbsp;<a href=""wap.asp"">"&toUnicode("返回")&"</a></p>"
  outCardFoot
end sub

'--------------------用户登录----------------
sub outLoginUser
	 Dim loginUser
	 response.write "<head><meta forua=""true"" http-equiv=""Cache-Control"" content=""max-age=0"" /></head>"
	 response.write "<card id=""splashscreen"" ontimer=""wap.asp"" title="""&toUnicode("登入信息")&""">"
	 
	 if not blog_wapLogin then
	     response.write "<timer value=""5000""/>"
		 response.write "<p>"&toUnicode("Blog不允许wap登录")&"</p>"
		 response.write "<p><a href=""wap.asp"">"&toUnicode("返回首页")&"</a></p>"
	 else
		 Session("GetCode")=Request.form("validate")
		 loginUser=login(Request.form("userName"),Request.form("Password"))
		 if not loginUser(3) then 
		     response.write "<timer value=""5000""/>"
			 response.write "<p>"&toUnicode("登入失败！")&"</p>"
			 response.write "<p><br/><a href=""wap.asp?do=Login"">"&toUnicode("重新登入")&"</a><br/><a href=""wap.asp"">"&toUnicode("返回首页")&"</a></p>"
		 else
		     response.write "<timer value=""30""/>"
			 response.write "<p>"&toUnicode("登入成功！")&toUnicode(Request.form("userName"))&toUnicode("欢迎你回来.")&"</p>"
		 	 response.write "<p>"&toUnicode("3秒后自动返回首页返回首页")&"</p>"
			 response.write "<p><br/><a href=""wap.asp"">"&toUnicode("返回首页")&"</a></p>"
		 end if
	  end if
	  outCardFoot
	  Conn.Close
	  Set Conn=Nothing
end sub

'--------------------用户登出----------------
sub outLogout
  logout(true)
  response.write "<head><meta forua=""true"" http-equiv=""Cache-Control"" content=""max-age=0"" /></head>"
  response.write "<card id=""splashscreen"" ontimer=""wap.asp"" title="""&toUnicode("登出信息")&""">"
  response.write "<timer value=""30""/>"
  response.write "<p>"&toUnicode("欢迎你已成功退出.")&"</p>"
  response.write "<p>"&toUnicode("3秒后自动返回首页返回首页")&"</p>"
  response.write "<p><br/><a href=""wap.asp"">"&toUnicode("返回首页")&"</a></p>"
  response.write "</card>"
  Conn.Close
  Set Conn=Nothing
end sub


'--------------------控制面板----------------
sub outControl
 response.write "<p>&nbsp;<br/>"
 if memName<>Empty then response.write "<b>"&toUnicode(memName)&"</b>"&toUnicode("，欢迎你!")&"<br/>"
 'if stat_AddAll=true or stat_Add=true then response.write "<a href=""blogpost.asp"">"&toUnicode("发表新日志")&"</a>"
 if request.QueryString("do") = "showLogDetail" and stat_CommentAdd and blog_wapComment then  response.write "<br/><a href=""#postCommentCard"">"&toUnicode("发表评论")&"</a>"
 if memName<>Empty then
      'if stat_AddAll and stat_Add then response.write "<br/><a href=""wap.asp?do=postLog"">"&toUnicode("发表日志")&"</a>"
	 response.write "<br/><a href=""wap.asp?do=Logout"">"&toUnicode("登出")&"</a>"
 else
     if blog_wapLogin then response.write "<br/><a href=""wap.asp?do=Login"">"&toUnicode("登入")&"</a>"
 end if  
 response.write "</p>"
end sub

'--------------------输出文件未----------------
sub outCardFoot 'Output Foot
 response.write "<p><br/>"&toUnicode("────────────")&"</p>"
 response.write "<p><a href=""wap.asp"">"&toUnicode(HTMLDecode(SiteName))&"</a></p>"
 response.write "<p><a href=""http://www.pjhome.net/wap.asp"">PJBlog2&nbsp;v"&blog_version&"</a>&nbsp;Inside.</p>"
 response.write "<p>Processed&nbsp;In&nbsp;"&FormatNumber(Timer()-StartTime,3,-1)&"&nbsp;ms</p>"
 response.write "<do type=""prev"" label="""&toUnicode("返回")&"""><prev/></do>"
 response.write "</card>"
end sub

sub postLog 
 
end sub

sub outPostComment
  dim postComment
  postComment=wap_CommentPost
  
  response.write "<head><meta forua=""true"" http-equiv=""Cache-Control"" content=""max-age=0"" /></head>"
  response.write "<card id=""splashscreen"" ontimer=""wap.asp"" title="""&toUnicode("发表评论")&""">"
  response.write "<p>"&postComment&"</p>"
  
  outCardFoot
  Conn.Close
  Set Conn=Nothing
end sub

function wap_CommentPost
  wap_CommentPost=""
  dim username,post_logID,post_Message,LastMSG,FlowControl,ReInfo
  username=trim(CheckStr(request.form("userName")))
  post_logID=clng(CheckStr(request.form("id")))
  post_Message=CheckStr(request.form("message"))
  
  set LastMSG=conn.execute("select top 1 comm_Content from blog_Comment where blog_ID="&post_logID&" order by comm_ID desc")
  if LastMSG.eof then 
     FlowControl=false
   else
    if LastMSG("comm_Content")=post_Message then FlowControl=true
  end if
  
  if filterSpam(post_Message,"spam.xml") and stat_Admin=false then
	  wap_CommentPost="<b>"&toUnicode("评论中包含被屏蔽的字符")&"</b><br/><a href=""wap.asp?do=showLogDetail&amp;id="&post_logID&"#postCommentCard"">"&toUnicode("返回")&"</a>"
      exit function 
  end if
  
  if FlowControl then 
	  wap_CommentPost="<b>"&toUnicode("禁止恶意灌水！")&"</b><br/><a href=""wap.asp?do=showLogDetail&amp;id="&post_logID&"#postCommentCard"">"&toUnicode("返回")&"</a>"
      exit function 
  end if 

  if DateDiff("s",Request.Cookies(CookieName)("memLastPost"),Now())<blog_commTimerout then 
	  wap_CommentPost="<b>"&toUnicode("发言太快,请 "&blog_commTimerout&" 秒后再发表评论")&"</b><br/><a href=""wap.asp?do=showLogDetail&amp;id="&post_logID&"#postCommentCard"">"&toUnicode("返回")&"</a>"
      exit function  
  end if
  
  if len(username)<1 then
	  wap_CommentPost="<b>"&toUnicode("请输入你的昵称.")&"</b><br/><a href=""wap.asp?do=showLogDetail&amp;id="&post_logID&"#postCommentCard"">"&toUnicode("返回")&"</a>"
      exit function  
  end if
  
  if IsValidUserName(username)=false then
	 wap_CommentPost="<b>"&toUnicode("非法用户名！")&"</b><br/><a href=""wap.asp?do=showLogDetail&amp;id="&post_logID&"#postCommentCard"">"&toUnicode("返回")&"</a>"
	 exit function
 end if

  if not stat_CommentAdd then
	 wap_CommentPost="<b>"&toUnicode("你没有权限发表评论")&"</b><br/><a href=""wap.asp?do=showLogDetail&amp;id="&post_logID&"#postCommentCard"">"&toUnicode("返回")&"</a>"
	 exit function
  end if
  
  if not blog_wapComment then
	 wap_CommentPost="<b>"&toUnicode("Blog不开放Wap发表评论")&"</b><br/><a href=""wap.asp?do=showLogDetail&amp;id="&post_logID&"#postCommentCard"">"&toUnicode("返回")&"</a>"
	 exit function
  end if
  
  if len(post_Message)<1 then
	 wap_CommentPost="<b>"&toUnicode("不允许发表空评论")&"</b><br/><a href=""wap.asp?do=showLogDetail&amp;id="&post_logID&"#postCommentCard"">"&toUnicode("返回")&"</a>"
	 exit function
  end if
  
  if len(post_Message)>blog_commLength then
	 wap_CommentPost="<b>"&toUnicode("评论超过最大字数限制")&"</b><br/><a href=""wap.asp?do=showLogDetail&amp;id="&post_logID&"#postCommentCard"">"&toUnicode("返回")&"</a>"
	 exit function
  end if
  
  dim checkMem
  if memName=empty then
       set checkMem=Conn.ExeCute("select top 1 mem_id from blog_Member where mem_Name='"&username&"'")
       if not checkMem.eof then
		 wap_CommentPost="<b>"&toUnicode("该用户名已存在，无法发表评论")&"</b><br/><a href=""wap.asp?do=showLogDetail&amp;id="&post_logID&"#postCommentCard"">"&toUnicode("返回")&"</a>"
    	 exit function
       end if
  end if 
  
  if Conn.ExeCute("select log_DisComment from blog_Content where log_ID="&post_logID)(0) then 
	 wap_CommentPost="<b>"&toUnicode("该日志不允许发表任何评论")&"</b><br/><a href=""wap.asp?do=showLogDetail&amp;id="&post_logID&"#postCommentCard"">"&toUnicode("返回")&"</a>"
	 exit function
  end if
  
  dim post_DisSM,post_DisURL,post_DisKEY,post_disImg,post_DisUBB
  post_DisSM=0
  post_DisURL=1
  post_DisKEY=1
  
  post_disImg=1
  post_DisUBB=0
  
 Dim AddComm
 AddComm=array(array("blog_ID",post_logID),array("comm_Content",post_Message),array("comm_Author",username),array("comm_DisSM",post_DisSM),array("comm_DisUBB",post_DisUBB),array("comm_DisIMG",post_disImg),array("comm_AutoURL",post_DisURL),Array("comm_PostIP",getIP),Array("comm_AutoKEY",post_DisKEY))
 DBQuest "blog_Comment",AddComm,"insert"
 'Conn.ExeCute("INSERT INTO blog_Comment(blog_ID,comm_Content,comm_Author,comm_DisSM,comm_DisUBB,comm_DisIMG,comm_AutoURL,comm_PostIP,comm_AutoKEY) VALUES ("&post_logID&",'"&post_Message&"','"&username&"',"&post_DisSM&","&post_DisUBB&","&post_disImg&","&post_DisURL&",'"&getIP()&"',"&post_DisKEY&")")
 Conn.ExeCute("update blog_Content set log_CommNums=log_CommNums+1 where log_ID="&post_logID)
 Conn.ExeCute("update blog_Info set blog_CommNums=blog_CommNums+1")
 Response.Cookies(CookieName)("memLastpost")=Now()
 getInfo(2)
 NewComment(2)
 
 if memName<>empty then conn.execute("update blog_Member set mem_PostComms=mem_PostComms+1 where mem_Name='"&memName&"'")
 PostArticle post_logID
 wap_CommentPost="<b>"&toUnicode("你成功地对该日志发表了评论")&"</b><br/><a href=""wap.asp?do=showLogDetail&amp;id="&post_logID&"#CommentCard"">"&toUnicode("返回")&"</a>"
end function

function toUnicode(str) 'To Unicode
    dim i,a
	For i = 1 to Len (str)
		a=Mid(str, i, 1)
		toUnicode=toUnicode & "&#x" & Hex(Ascw(a)) & ";"
	next
end function

function toUnicodeJS(str) 'To Unicode
    dim i,a,ignoreStr
    ignoreStr ="1234567890`~!@#$%^&*()_+=-{}][|\"":;'?><,./qwertyuioplkjhgfdsazxcvbnmQWERTYUIOPLKJHGFDSAZXCVBNM"
	For i = 1 to Len (str)
		a=Mid(str, i, 1)
	    if Instr(ignoreStr, a)>0 then 
		     toUnicodeJS=toUnicodeJS & a
	     else
			 toUnicodeJS=toUnicodeJS & "&#x" & Hex(Ascw(a)) & ";"
		end if
	next
end function
%>

<script runat="server" Language="javascript">
	function HtmlToText(str) {
	    //filter HTMLToUBBAndText
		str = str.replace(/\r/g,"");
		str = str.replace(/on(load|click|dbclick|mouseover|mousedown|mouseup)="[^"]+"/ig,"");
		if (blog_wapImg) str = str.replace(/<img[^>]+src="([^"]+)"[^>]*>/ig,"\n[img]$1[/img]\n");
		if (blog_wapHTML) {
			str = str.replace(/<a[^>]+href="([^"]+)"[^>]*>(.*?)<\/a>/ig,"[url=$1]$2[/url]");
			str = str.replace(/<([\/]?)b>/ig,"[$1b]");
			str = str.replace(/<([\/]?)strong>/ig,"[$1b]");
			str = str.replace(/<p>(.*?)<\/p>/ig,"$1<br/>");
			str = str.replace(/<([\/]?)i>/ig,"[$1i]");
			str = str.replace(/<([\/]?)u>/ig,"[$1u]");
		}
		str = str.replace(/<br\/>/ig,"\n");
		str = str.replace(/<[^>]*?>/g,"");
		str = str.replace(/&/g,"&amp;");
		str = str.replace(/&amp;(.*?[^\s]);/g,"&$1;");
		str = str.replace(/\n+/g,"<br/>");

		if (blog_wapHTML) {
		//filter UBBToText
			str=str.replace(/(\[i\])(.*?)(\[\/i\])/ig,"<i>$2</i>")
			str=str.replace(/(\[u\])(.*?)(\[\/u\])/ig,"<u>$2</u>")
			str=str.replace(/(\[b\])(.*?)(\[\/b\])/ig,"<b>$2</b>")		
			if (blog_wapURL) 
			  str=str.replace(/(\[url=(.[^\[]*)\])(.*?)(\[\/url\])/ig,"<a href=\"$2\">$3</a>");
			 else
			  str=str.replace(/(\[url=(.[^\[]*)\])(.*?)(\[\/url\])/ig,"[<i><u>$3</u></i>]");
		}
		if (blog_wapImg) str=str.replace(/\[img\](.+?)\[\/img\]/ig,"<img src=\"$1\" alt=\"Image\"/>");
	
		return toUnicodeJS(str).replace(/&#x20;/g," ");
	}
</script>