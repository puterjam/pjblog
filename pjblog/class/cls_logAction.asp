<%
'==================================
'  日志编辑类
'    更新时间: 2006-1-22
'==================================
class logArticle
  Private weblog
  public categoryID,logTitle,logAuthor,logEditType
  public logIsShow,logIsDraft,logWeather,logLevel,logCommentOrder
  public logDisableComment,logIsTop,logFrom,logFromURL
  public logDisableImage,logDisableSmile,logDisableURL,logDisableKeyWord
  public logQuote,logMessage,logIntro,logIntroCustom,logTags,logPublishTimeType,logPubTime,logTrackback,logCommentCount,logQuoteCount,logViewCount
  Private logUbbFlags,PubTime,sqlString
  
  Private Sub Class_Initialize()
   Set weblog=Server.CreateObject("ADODB.RecordSet")
         categoryID = 0
         logTitle = ""
         logEditType = 1
         logIntroCustom = 0
         logIntro = ""
         logAuthor = "null"
         logWeather = "sunny"
         logLevel = "level3"
         logCommentOrder = 1
         logDisableComment = 0
         logIsShow = true
         logIsTop = false
         logIsDraft = false
         logFrom = "本站原创"
         logFromURL = siteURL
         logDisableImage = 0
         logDisableSmile = 0
         logDisableURL = 0
         logDisableKeyWord = 0
         logCommentCount = 0
         logQuoteCount = 0
         logViewCount = 0
         logMessage = ""
         logTrackback = ""
         logTags = ""
	     logPubTime="2006-1-1 00:00:00"
	     logPublishTimeType="now"
  end sub
  
  Private Sub Class_Terminate()
  Set weblog=nothing
  end sub
 
  '*********************************************
  '发表新日志
  '*********************************************
  public function postLog()
    postLog=array(-4,"准备发表日志",-1)
    weblog.Open "blog_Content",Conn,1,2
    SQLQueryNums=SQLQueryNums+1
    
    if stat_AddAll<>true and stat_Add<>true Then
      postLog=array(-3,"没有权限发表日志",-1)
      exit function
    end if
             
         '-------------------处理Tags--------------------
         dim tempTags,loadTagString,loadTags,loadTag,getTags
         tempTags=Split(CheckStr(logTags),",")
         
	     set getTags=new Tag
        
	    dim post_tag,post_taglist
	    post_taglist=""
	        
	    '添加新的Tag
	    for each post_tag in tempTags
            if len(trim(post_tag))>0 then
	 	     post_taglist = post_taglist & "{" & getTags.insert(CheckStr(trim(post_tag))) & "}"
	 	   end if
	    next
	    logTags = post_taglist
	    call Tags(2)
	    set getTags=nothing
        '--------------处理日期---------------------
        if CheckStr(logPublishTimeType)="now" then
           PubTime = DateToStr(now(),"Y-m-d H:I:S")
         else
           PubTime = DateToStr(CheckStr(logPubTime),"Y-m-d H:I:S")
        end if
        
        '---------------分割日志--------------------
         if logIntroCustom=1 then
	          if int(logEditType)=1 then
		           logIntro = closeUBB(CheckStr(HTMLEncode(logIntro)))
	          else
		           logIntro = closeHTML(CheckStr(logIntro))
	          end if
         else
	         if int(logEditType)=1 then
	                if blog_SplitType then 
		                logIntro = closeUBB(SplitLines(CheckStr(HTMLEncode(logMessage)),blog_introLine))
	                else
		                logIntro = closeUBB(CutStr(CheckStr(HTMLEncode(logMessage)),blog_introChar))
	                end if
	       	 else
	                logIntro = closeHTML(SplitLines(CheckStr(logMessage),blog_introLine))
	         end if
        end If

         '日志基本状态 
         logIsShow=cbool(logIsShow)
         logCommentOrder=cbool(logCommentOrder)
	 	 logDisableComment=cbool(logDisableComment)
         logIsTop=cbool(logIsTop)
         logIsDraft=cbool(logIsDraft)

         'UBB 特别属性
         if logDisableSmile=1 then logDisableSmile=1 else logDisableSmile=0
         if logDisableImage=1 then logDisableImage=1 else logDisableImage=0
         if logDisableURL=1 then logDisableURL=0 else logDisableURL=1
         if logDisableKeyWord=1 then logDisableKeyWord=0 else logDisableKeyWord=1
         if logIntroCustom=1 then logIntroCustom=0 else logIntroCustom=1
         logUbbFlags=logDisableSmile & "0" & logDisableImage & logDisableURL & logDisableKeyWord & logIntroCustom

		 weblog.addNew
         weblog("log_CateID")=CheckStr(categoryID)
         weblog("log_Author")=CheckStr(logAuthor)
         weblog("log_Title")=CheckStr(logTitle)
         weblog("log_weather")=CheckStr(logWeather)
         weblog("log_Level")=CheckStr(logLevel)
         weblog("log_From")=CheckStr(logFrom)
         weblog("log_FromURL")=CheckStr(logFromURL)
         weblog("log_Content")=CheckStr(logMessage)
         weblog("log_Intro")=logIntro
         weblog("log_tag")=logTags
         weblog("log_ubbFlags")=logUbbFlags
         weblog("log_IsShow")=logIsShow
         weblog("log_IsTop")=logIsTop
	     weblog("log_PostTime")=PubTime
         weblog("log_IsDraft")=logIsDraft
         weblog("log_DisComment")=logDisableComment
         weblog("log_edittype")=logEditType
         weblog("log_comorder")=logCommentOrder
         SQLQueryNums=SQLQueryNums+2
         weblog.update
	   	 weblog.close
	   	 
	   	 
	   	    '------------------统计日志-----------------------------
	        Dim PostLogID
	        PostLogID=Conn.ExeCute("SELECT TOP 1 log_ID FROM blog_Content ORDER BY log_ID DESC")(0)
	        Conn.ExeCute("UPDATE blog_Member SET mem_PostLogs=mem_PostLogs+1 WHERE mem_Name='"&logAuthor&"'")
	        if not logIsDraft then
		        Conn.ExeCute("UPDATE blog_Info SET blog_LogNums=blog_LogNums+1")
		        Conn.ExeCute("UPDATE blog_Category SET cate_count=cate_count+1 where cate_ID="&categoryID)
	            SQLQueryNums=SQLQueryNums+2
	        end if
        
        	   	 
            '-------------------输出静态日志档案--------------------
            dim preLog,nextLog
            '输出日志到文件
            PostArticle PostLogID
	        
	        '输出附近的日志到文件
	 	    set preLog=Conn.Execute("SELECT TOP 1 log_Title,log_ID FROM blog_Content WHERE log_PostTime<#"&PubTime&"# and log_IsShow=true and log_IsDraft=false ORDER BY log_PostTime DESC")
	        set nextLog=Conn.Execute("SELECT TOP 1 log_Title,log_ID FROM blog_Content WHERE log_PostTime>#"&PubTime&"# and log_IsShow=true and log_IsDraft=false ORDER BY log_PostTime ASC")
	        if not preLog.eof then PostArticle preLog("log_ID")
	        if not nextLog.eof then PostArticle nextLog("log_ID")
	        
			call updateCache
			
		    Session(CookieName&"_LastDo")="AddArticle"
		    session(CookieName&"_draft_"&logAuthor)=conn.execute("select count(log_ID) from blog_Content where log_Author='"&logAuthor&"' and log_IsDraft=true")(0)
	        SQLQueryNums=SQLQueryNums+1
	        

        	        
	        if logIsDraft then
			    postLog=array(1,"日志成功保存为草稿",PostLogID)
	        else
			    postLog=array(0,"恭喜!日志发表成功",PostLogID)
	        end if
	        
	        '-------------------引用通告-------------------
	        IF logTrackback<>Empty And logIsShow=True and logIsDraft=false Then
Dim log_QuoteEvery,log_QuoteArr,logid,LastID
                                set LastID=Conn.Execute("SELECT TOP 1 log_ID FROM blog_Content ORDER BY log_ID DESC")
                                logid=LastID("log_ID")
                         log_QuoteArr=Split(logTrackback,",")
                         For Each log_QuoteEvery In log_QuoteArr
                                 Trackback Trim(log_QuoteEvery), siteURL&"default.asp?id="&logid, logTitle, CutStr(CheckStr(logIntro),252), siteName
                                set LastID=Nothing
	         	Next
	        End IF
  end function
  
  '*********************************************
  '编辑日志
  '*********************************************
  public function editLog(id)
      editLog=array(-4,"准备编辑日志",-1)
    if isEmpty(id) then
      getLog=array(-5,"ID号不能为空")
      exit function
    end if
    if not IsInteger(id) then
      editLog=array(-1,"非法ID号",-1)
      exit function
    end if
    
    sqlString="SELECT top 1 * FROM blog_Content WHERE log_ID="&id&""
    weblog.Open sqlString,Conn,1,3
    SQLQueryNums=SQLQueryNums+1
    
    if weblog.eof or weblog.bof then
      editLog=array(-2,"无法找到相应文章",-1)
      exit function
    end if

    if stat_EditAll<>true and (stat_Edit and weblog("log_Author")=logAuthor)<>true Then
      editLog=array(-3,"您没有权限编辑日志",-1)
      exit function
    end if
    
    logAuthor=weblog("log_Author")
         Conn.ExeCute("UPDATE blog_Category SET cate_count=cate_count-1 where cate_ID="&weblog("log_CateID"))
         Conn.ExeCute("UPDATE blog_Category SET cate_count=cate_count+1 where cate_ID="&CheckStr(categoryID))

             
         '-------------------处理Tags--------------------
         dim tempTags,loadTagString,loadTags,loadTag,getTags
         tempTags=Split(CheckStr(logTags),",")
         loadTagString=weblog("log_tag")
         
	     set getTags=new Tag
	    
	     '清除旧的Tag
         if len(loadTagString)>0 then
	         loadTagString=replace(loadTagString,"}{",",")
	         loadTagString=replace(loadTagString,"}","")
	         loadTagString=replace(loadTagString,"{","")
	         loadTags=Split(loadTagString,",")
	        
	         for each loadTag in loadTags
		          getTags.remove loadTag
	         next
         end if
        
	    dim post_tag,post_taglist
	    post_taglist=""
	        
	    '添加新的Tag
	    for each post_tag in tempTags
            if len(trim(post_tag))>0 then
	 	     post_taglist = post_taglist & "{" & getTags.insert(CheckStr(trim(post_tag))) & "}"
	 	   end if
	    next
	    logTags = post_taglist
	    call Tags(2)
	    set getTags=nothing
        '--------------处理日期---------------------
        if CheckStr(logPublishTimeType)="now" then
           PubTime = DateToStr(now(),"Y-m-d H:I:S")
         else
           PubTime = DateToStr(CheckStr(logPubTime),"Y-m-d H:I:S")
        end if
        
        '---------------分割日志--------------------
         if logIntroCustom=1 then
	          if int(logEditType)=1 then
		           logIntro = closeUBB(CheckStr(HTMLEncode(logIntro)))
	          else
		           logIntro = closeHTML(CheckStr(logIntro))
	          end if
         else
	         if int(logEditType)=1 then
	                if blog_SplitType then 
		                logIntro = closeUBB(SplitLines(CheckStr(HTMLEncode(logMessage)),blog_introLine))
	                else
		                logIntro = closeUBB(CutStr(CheckStr(HTMLEncode(logMessage)),blog_introChar))
	                end if
	       	 else
	                logIntro = closeHTML(SplitLines(CheckStr(logMessage),blog_introLine))
	         end if
        end If

         '日志基本状态 
         logIsShow=cbool(logIsShow)
         logCommentOrder=cbool(logCommentOrder)
	 	 logDisableComment=cbool(logDisableComment)
         logIsTop=cbool(logIsTop)
         logIsDraft=cbool(logIsDraft)
         
         'UBB 特别属性
         if logDisableSmile=1 then logDisableSmile=1 else logDisableSmile=0
         if logDisableImage=1 then logDisableImage=1 else logDisableImage=0
         if logDisableURL=1 then logDisableURL=0 else logDisableURL=1
         if logDisableKeyWord=1 then logDisableKeyWord=0 else logDisableKeyWord=1
         if logIntroCustom=1 then logIntroCustom=0 else logIntroCustom=1
         logUbbFlags=logDisableSmile & "0" & logDisableImage & logDisableURL & logDisableKeyWord & logIntroCustom

         if logIsDraft=false then weblog("log_Modify")="[本日志由 "&memName&" 于 "&DateToStr(now(),"Y-m-d H:I A")&" 编辑]"
         if logIsDraft=false and weblog("log_IsDraft")<>logIsDraft then
	        Conn.ExeCute("UPDATE blog_Info SET blog_LogNums=blog_LogNums+1")
	        Conn.ExeCute("UPDATE blog_Category SET cate_count=cate_count+1 where cate_ID=" & CheckStr(categoryID))
            SQLQueryNums=SQLQueryNums+2
         end if

         weblog("log_Title")=CheckStr(logTitle)
         weblog("log_weather")=CheckStr(logWeather)
         weblog("log_Level")=CheckStr(logLevel)
         weblog("log_From")=CheckStr(logFrom)
         weblog("log_FromURL")=CheckStr(logFromURL)
         weblog("log_Content")=CheckStr(logMessage)
         weblog("log_Intro")=logIntro
         weblog("log_CateID")=CheckStr(categoryID)
         weblog("log_tag")=logTags
         weblog("log_ubbFlags")=logUbbFlags
         weblog("log_IsShow")=logIsShow
         weblog("log_IsTop")=logIsTop
	     weblog("log_PostTime")=PubTime
         weblog("log_IsDraft")=logIsDraft
         weblog("log_DisComment")=logDisableComment
         weblog("log_edittype")=logEditType
         weblog("log_comorder")=logCommentOrder
         SQLQueryNums=SQLQueryNums+2
         weblog.update
	   	 weblog.close
	   	 
            dim preLog,nextLog
            
            '-------------------输出静态日志档案--------------------
            '输出日志到文件
            PostArticle logid
	        
	        '输出附近的日志到文件
	 	    set preLog=Conn.Execute("SELECT TOP 1 log_Title,log_ID FROM blog_Content WHERE log_PostTime<#"&PubTime&"# and log_IsShow=true and log_IsDraft=false ORDER BY log_PostTime DESC")
	        set nextLog=Conn.Execute("SELECT TOP 1 log_Title,log_ID FROM blog_Content WHERE log_PostTime>#"&PubTime&"# and log_IsShow=true and log_IsDraft=false ORDER BY log_PostTime ASC")
	        if not preLog.eof then PostArticle preLog("log_ID")
	        if not nextLog.eof then PostArticle nextLog("log_ID")
	        
			call updateCache
			
	 	    Session(CookieName&"_LastDo")="EditArticle"
	        Session(CookieName&"_draft_"&logAuthor)=conn.execute("select count(log_ID) from blog_Content where log_Author='"&logAuthor&"' and log_IsDraft=true")(0)
	        SQLQueryNums=SQLQueryNums+1
	        
	        if logIsDraft then
			    editLog=array(1,"日志成功保存为草稿",id)
	        else
			    editLog=array(0,"恭喜!日志编辑成功",id)
	        end if
	        
	        '-------------------引用通告-------------------
	        IF logTrackback<>Empty And logIsShow=True and logIsDraft=false Then
	         	Dim log_QuoteEvery,log_QuoteArr
	         	log_QuoteArr=Split(logTrackback,",")
	         	For Each log_QuoteEvery In log_QuoteArr
	         		Trackback Trim(log_QuoteEvery), siteURL&"default.asp?id="&logid, logTitle, CutStr(CheckStr(logIntro),252), siteName
	         	Next
	        End IF
  end function
  
  '*********************************************
  '删除日志
  '*********************************************
  public function deleteLog(id)
   deleteLog=array(-4,"准备删除")
   if isEmpty(id) then
      getLog=array(-5,"ID号不能为空")
      exit function
    end if  
  
   if not IsInteger(id) then
      deleteLog=array(-1,"非法ID号")
      exit function
    end if
    
    sqlString="SELECT top 1 * FROM blog_Content WHERE log_ID="&id&""
    weblog.Open sqlString,Conn,1,3
    SQLQueryNums=SQLQueryNums+1
    
    if weblog.eof or weblog.bof then
      deleteLog=array(-2,"找不到相应文章")
      exit function
    end if
  
    if stat_DelAll<>true and (stat_Del and weblog("log_Author")=logAuthor)<>true Then
      deleteLog=array(-3,"没有权限删除")
      exit function
    end if
    
    dim Pdate,getTag
    dim tempTags,loadTagString,loadTags,loadTag,getTags,post_tag
    Pdate=weblog("log_PostTime")
    Conn.ExeCute("UPDATE blog_Member SET mem_PostLogs=mem_PostLogs-1 WHERE mem_Name='"&weblog("log_Author")&"'")
     if not weblog("log_IsDraft") then
         Conn.ExeCute("UPDATE blog_Category SET cate_count=cate_count-1 where cate_ID="&weblog("log_CateID"))
	     Conn.ExeCute("UPDATE blog_Info SET blog_LogNums=blog_LogNums-1")
	     Conn.ExeCute("update blog_Info set blog_CommNums=blog_CommNums-"&weblog("log_CommNums"))  
     end if
     
    loadTag=weblog("log_tag")
	set getTag=new Tag
	    
	    '清除旧的Tag
        if len(loadTag)>0 then
	         loadTag=replace(loadTag,"}{",",")
	         loadTag=replace(loadTag,"}","")
	         loadTag=replace(loadTag,"{","")
	         loadTags=Split(loadTag,",")
	        
	         for each post_tag in loadTags
		          getTag.remove post_tag
	         next
        end if
     call Tags(2)
     set getTag=nothing
     dim preLog,nextLog
     Conn.ExeCute("DELETE * FROM blog_Content WHERE log_ID="&id)
     Conn.ExeCute("DELETE * FROM blog_Comment WHERE blog_ID="&id)
     DeleteFiles Server.MapPath("post/"&logid&".asp")
     DeleteFiles Server.MapPath("cache/"&logid&".asp")
     DeleteFiles Server.MapPath("cache/c_"&logid&".js")
     
 	    set preLog=Conn.Execute("SELECT TOP 1 log_Title,log_ID FROM blog_Content WHERE log_PostTime<#"&Pdate&"# and log_IsShow=true and log_IsDraft=false ORDER BY log_PostTime DESC")
        set nextLog=Conn.Execute("SELECT TOP 1 log_Title,log_ID FROM blog_Content WHERE log_PostTime>#"&Pdate&"# and log_IsShow=true and log_IsDraft=false ORDER BY log_PostTime ASC")
        '输出附近的日志到文件
	    if not preLog.eof then PostArticle preLog("log_ID")
	    if not nextLog.eof then PostArticle nextLog("log_ID")
     SQLQueryNums=SQLQueryNums+5
     weblog.close
     
	 call updateCache
	 
 	 Session(CookieName&"_LastDo")="DelArticle"
     session(CookieName&"_draft_"&logAuthor)=conn.execute("select count(log_ID) from blog_Content where log_Author='"&logAuthor&"' and log_IsDraft=true")(0)
     SQLQueryNums=SQLQueryNums+1    
     deleteLog=array(0,"删除成功")
	
  end function
  
  '*********************************************
  '获得日志
  '*********************************************
  public function getLog(id)
   dim getTag
   getLog=array(-3,"准备提取日志")
   if isEmpty(id) then
      getLog=array(-4,"ID号不能为空")
      exit function
    end if
       
   if not IsInteger(id) then
      getLog=array(-1,"非法ID号")
      exit function
    end if
    
    sqlString="SELECT top 1 log_CateID,log_Author,log_Title,log_edittype,log_ubbFlags,log_Intro,log_weather,log_Level,log_comorder,log_DisComment,log_IsShow,log_IsTop,log_IsDraft,log_From,log_FromURL,log_Content,log_tag,log_PostTime,log_CommNums,log_QuoteNums,log_ViewNums FROM blog_Content WHERE log_ID="&id&""
    weblog.Open sqlString,Conn,1,1
    SQLQueryNums=SQLQueryNums+1
    
    if weblog.eof or weblog.bof then
      getLog=array(-2,"找不到相应文章")
      exit function
    end if
    
         categoryID = weblog("log_CateID")
         logAuthor = weblog("log_Author")
         logTitle = weblog("log_Title")
         logEditType = weblog("log_edittype")
         logIntroCustom = mid(weblog("log_ubbFlags"),6,1)
         logIntro = weblog("log_Intro")
         logWeather = weblog("log_weather")
         logLevel = weblog("log_Level")
         logCommentOrder = weblog("log_comorder")
         logDisableComment = weblog("log_DisComment")
         logIsShow = weblog("log_IsShow")
         logIsTop = weblog("log_IsTop")
         logIsDraft = weblog("log_IsDraft")
         logFrom = weblog("log_From")
         logFromURL = weblog("log_FromURL")
         logDisableImage = mid(weblog("log_ubbFlags"),3,1)
         logDisableSmile = mid(weblog("log_ubbFlags"),1,1)
         logDisableURL = mid(weblog("log_ubbFlags"),4,1)
         logDisableKeyWord = mid(weblog("log_ubbFlags"),5,1)
         logMessage = weblog("log_Content")
         logCommentCount = weblog("log_CommNums")
         logQuoteCount = weblog("log_QuoteNums")
         logViewCount = weblog("log_ViewNums")
         logTrackback = ""
         set getTag=new tag
         logTags = getTag.filterEdit(weblog("log_tag"))
		 set getTag=nothing
	     logPubTime = weblog("log_PostTime")
	     logPublishTimeType = "now"
	     
	     weblog.close
	     getLog=array(0,"成功获取日志")
	     
  end function
  
  '*********************************************
  '删除文件
  '*********************************************
  Private Function DeleteFiles(FilePath)
      Dim FSO
      Set FSO=Server.CreateObject("Scripting.FileSystemObject")
      IF FSO.FileExists(FilePath) Then
      	FSO.DeleteFile FilePath,True
      	DeleteFiles = True
      Else
      	DeleteFiles = false
      End IF
      Set FSO = Nothing
  End Function
  
  '*********************************************
  '更新缓存
  '*********************************************
  Private sub updateCache
     call archive(2)
     call CategoryList(2)
     call getInfo(2)
     call NewComment(2)
	 call Calendar("","","",2)
	 
	 if blog_postFile then
	     dim lArticle
		 set lArticle = new ArticleCache
		 lArticle.SaveCache
		 set lArticle = nothing
	 end if
  end sub
end Class
%>

<%
'======================================================
'  PJblog2 静态缓存类
'======================================================
class ArticleCache
  Private cacheList
  Private cacheStream
  Private errorCode

  Private Sub Class_Initialize()
    cacheList=""
  end sub
  
  Private Sub Class_Terminate()

  end sub
  
  Private function clearT(str)
    dim tempLen
    tempLen=len(str)
    if tempLen>0 then
      str=left(str,tempLen-1)
    end if
    clearT=str
  end function
  
  Private function LoadIntro(id,aRight,outType)
   dim getIntro,tempI,TempStr,getC,author
   getIntro=LoadFile("cache/" & id & ".asp")
   if getIntro = "error" then
   		if stat_Admin then
		response.write "<div style=""color:#f00"">编号为[" + id+ "]的日志读取失败！建议您重新 <a href=""blogedit.asp?id="&id&""" title=""编辑该日志"" accesskey=""E"">编辑</a> 此文章获得新的缓存</div>"
		end if
		exit function
   end if
   getIntro=split(getIntro,"<"&"%ST(A)%"&">")
   author=trim(getIntro(1))
   if outType="list" then
       if cbool(int(aRight)) or stat_Admin or (not cbool(int(aRight)) and memName=author) then 
	  	   tempI=getIntro(4)
  	   else
	  	   tempI=getIntro(6)
  	   end if
	   tempI=Replace(tempI,"<$log_viewC$>",getIntro(2))
	   response.write tempI
	else
		TempStr=""
		if stat_EditAll or (stat_Edit and memName=author) then 
			 TempStr=TempStr&" | <a href=""blogedit.asp?id="&id&""" title=""编辑该日志"" accesskey=""E""><img src=""images/icon_edit.gif"" alt="""" border=""0"" style=""margin-bottom:-2px""/></a> "
		end if
			   
		if stat_DelAll or (stat_Del and memName=author) then 
	    	 TempStr=TempStr&" | <a href=""blogedit.asp?action=del&amp;id="&id&""" onclick=""if (!window.confirm('是否要删除该日志')) return false"" title=""删除该日志"" accesskey=""K""><img src=""images/icon_del.gif"" alt="""" border=""0"" style=""margin-bottom:-2px""/></a>"
		end if	
       if cbool(int(aRight)) or stat_Admin or (not cbool(int(aRight)) and memName=author) then 
	  	   tempI=getIntro(3)
  	   else
	  	   tempI=getIntro(5)
  	   end if
  	   tempI=Replace(tempI,"<"&"%Article In PJblog2%"&">","")
	   tempI=Replace(tempI,"<$editRight$>",TempStr)
	   tempI=Replace(tempI,"<$log_viewC$>",getIntro(2))
	   response.write tempI
   end if
  end function
  
  Private function LoadFile(ByVal File)
  	  On Error Resume Next
	  LoadFile = "error"
	  With cacheStream
        .Type = 2
        .Mode = 3
        .Open
        .Charset = "utf-8"
        .Position = cacheStream.Size
        .LoadFromFile Server.MapPath(File)
        If Err Then
           .Close
           Err.Clear
	       exit function
        End If
        LoadFile = .ReadText
        .Close
    End With  
  end function
  
  public function outHTML(loadType,outType,title)
	    dim re, strMatchs, strMatch,i,j,id,aRight,hiddenC
	    Set cacheStream = Server.CreateObject("ADODB.Stream")
	    Set re=new RegExp
		re.IgnoreCase =True
		re.Global=True
 		re.Pattern="\[""([^\r]*?)"";([^\r]*?);\(([^\r]*?)\)\]"
		Set strMatchs=re.Execute(cacheList)
		For Each strMatch in strMatchs
		    if loadType=strMatch.SubMatches(0) then 
			    dim aList,pageSize
			    pageSize = blogPerPage
			    if outType="list" then pageSize=pageSize*4
			    aList=split(strMatch.SubMatches(2),",")
			    hiddenC=strMatch.SubMatches(1)
			    if stat_Admin Or stat_ShowHiddenCate then hiddenC=0
			    if (ubound(aList)+1-hiddenC)>0 then
					    %>
					      <div class="pageContent" style="text-align:Right;overflow:hidden;height:18px;line-height:140%"><span style="float:left"><%=title%></span><%=MultiPage(ubound(aList)+1-hiddenC,pageSize,CurPage,Url_Add,"","float:Left")%> 预览模式: <a href="<%=Url_Add%>distype=normal" accesskey="1">普通</a> | <a href="<%=Url_Add%>distype=list" accesskey="2">列表</a></div>
					    <%
					    if outType="list" then response.write "<div class=""Content-body"" style=""text-align:Left""><table cellpadding=""2"" cellspacing=""2"" width=""100%"">"
					    i=0
					    Do Until i >= pageSize
					      j = i + (CurPage-1)*pageSize
					      if j<=ubound(aList) then
					          id=split(aList(j),"|")(1)
							  aRight=split(aList(j),"|")(0)
						      LoadIntro id,aRight,outType
						      i=i+1
					       else
							   if outType="list" then response.write "</table></div>"
							    %>
								 <div class="pageContent"><%=MultiPage(ubound(aList)+1-hiddenC,pageSize,CurPage,Url_Add,"","float:Left")%></div>
							    <%
					    		exit for
					      end if
					    Loop
					    if outType="list" then response.write "</table></div>"
					    %>
						 <div class="pageContent"><%=MultiPage(ubound(aList)+1-hiddenC,pageSize,CurPage,Url_Add,"","float:Left")%></div>
					    <%
			    else
				     response.write "<b>抱歉，没有找到任何日志！</b>"
			    end if
			    set re=nothing
			    exit function
		    end if
		Next
	    set re = nothing
	    Set cacheStream = nothing
  end function
    
  public function loadCache
    dim LoadList
    if not blog_postFile then loadCache=false:exit function
	LoadList=LoadFromFile("cache/listCache.asp")
    if LoadList(0)=0 then
     cacheList=LoadList(1)
     loadCache=true
    else
     loadCache=false
    end if
  end function
  
  public function SaveCache
	if not blog_postFile then exit function
    Dim LogList,LogListArray,SaveList,CateDic,CateHDic,TagsDic
    set CateDic=Server.CreateObject("Scripting.Dictionary")
    set CateHDic=Server.CreateObject("Scripting.Dictionary")
    set TagsDic=Server.CreateObject("Scripting.Dictionary")

    SQL="select T.log_ID,T.log_CateID,T.log_IsShow,C.cate_Secret FROM blog_Content As T,blog_Category As C where T.log_CateID=C.cate_ID and log_IsDraft=false ORDER BY log_IsTop ASC,log_PostTime DESC"
    set LogList=conn.execute(SQL)
    if LogList.EOF or LogList.BOF then
	     SaveList=SaveToFile("[""A"";0;()]" & chr(13) & "[""G"";0;()]","cache/listCache.asp")
	     set LogList=nothing
	     exit function
    end if
    LogListArray = LogList.GetRows()
	set LogList=nothing
	dim i,AList,AListC,GList,GListC,outIndex,tempS,tempCS,hiddenC
	AList=""
	AListC=0
	GList=""
	GListC=0
	outIndex=""
	for i=0 to ubound(LogListArray,2)
	  tempS=1
	  hiddenC=1
	  'response.write LogListArray(0,i) & " "
	  if not LogListArray(2,i) then tempS=0
	  if not LogListArray(3,i) then 
	   tempCS=0:hiddenC=0
	   GList=GList & tempS & "|" & LogListArray(0,i) & ","
	   GListC=GListC+hiddenC
	  end if
	  
	  AList=AList & tempS & "|" & LogListArray(0,i) & ","
	  AListC=AListC+hiddenC
	  if not CateDic.Exists("C"&LogListArray(1,i)) then
		  CateDic.add "C"&LogListArray(1,i),tempS & "|" & LogListArray(0,i)&","
		else
	 	  CateDic.Item("C"&LogListArray(1,i))=CateDic.Item("C"&LogListArray(1,i)) & tempS & "|" & LogListArray(0,i) & ","
	  end if
	  
	  if not CateHDic.Exists("CH"&LogListArray(1,i)) then
		  CateHDic.add "CH"&LogListArray(1,i),hiddenC
		else
	 	  CateHDic.Item("CH"&LogListArray(1,i))=CateHDic.Item("CH"&LogListArray(1,i)) + hiddenC
	  end if
	  
	 next
	outIndex=outIndex & "[""A"";"&AListC&";("&clearT(AList)&")] " & chr(13)
	outIndex=outIndex & "[""G"";"&GListC&";("&clearT(GList)&")] " & chr(13)
	dim CateKeys,CateItems,CateHKeys,CateHItems
	CateKeys=CateDic.Keys
	CateItems=CateDic.Items
	CateHKeys=CateHDic.Keys
	CateHItems=CateHDic.Items
	for i=0 to CateDic.Count-1
		outIndex=outIndex & "["""&CateKeys(i)&""";"&CateHItems(i)&";("&clearT(CateItems(i))&")] " & chr(13)
	next
	
	SaveList=SaveToFile(outIndex,"cache/listCache.asp")
	
	set CateDic=nothing
	set CateHDic=nothing
	set TagsDic=nothing
  end function
  
end class
%>

<%
'======================================================
'  PJblog2 动态文章保存
'======================================================

sub PostArticle(LogID)
 if not blog_postFile then exit sub
  dim SaveArticle,LoadTemplate1,LoadTemplate2,Temp1,Temp2,TempStr,log_View,preLogC,nextLogC
 '读取日志模块
  LoadTemplate1=LoadFromFile("Template/Article.asp")
  LoadTemplate2=LoadFromFile("Template/ArticleList.asp")
  if LoadTemplate1(0)=0 and LoadTemplate2(0)=0 then '读取成功后写入信息
    '读取分类信息
    Temp1=LoadTemplate1(1)
    Temp2=LoadTemplate2(1)
    
    '读取日志内容
    SQL="SELECT TOP 1 * FROM blog_Content WHERE log_ID=" & LogID
    SQLQueryNums=SQLQueryNums+1
    Set log_View=conn.execute(SQL) 
     dim blog_Cate,blog_CateArray,comDesc
      Dim getCate,getTags
      set getCate=new Category
      set getTags=new tag
	  getCate.load(int(log_View("log_CateID"))) '获取分类信息
     
     Temp1=Replace(Temp1,"<$Cate_icon$>",getCate.cate_icon)
     Temp1=Replace(Temp1,"<$Cate_Title$>",getCate.cate_Name)
     Temp1=Replace(Temp1,"<$log_CateID$>",log_View("log_CateID"))
     Temp1=Replace(Temp1,"<$LogID$>",LogID)   
     Temp1=Replace(Temp1,"<$log_Title$>",HtmlEncode(log_View("log_Title")))
     Temp1=Replace(Temp1,"<$log_Author$>",log_View("log_Author"))
     Temp1=Replace(Temp1,"<$log_PostTime$>",DateToStr(log_View("log_PostTime"),"Y-m-d"))
 
     Temp2=Replace(Temp2,"<$Cate_icon$>",getCate.cate_icon)
     Temp2=Replace(Temp2,"<$Cate_Title$>",getCate.cate_Name)
     Temp2=Replace(Temp2,"<$log_CateID$>",log_View("log_CateID"))
     Temp2=Replace(Temp2,"<$LogID$>",LogID)   
     Temp2=Replace(Temp2,"<$log_Title$>",HtmlEncode(log_View("log_Title")))
     Temp2=Replace(Temp2,"<$log_Author$>",log_View("log_Author"))
     Temp2=Replace(Temp2,"<$log_PostTime$>",DateToStr(log_View("log_PostTime"),"Y-m-d"))
     Temp2=Replace(Temp2,"<$log_viewCount$>",log_View("log_ViewNums"))
 
     if log_View("log_IsTop") then
	     Temp2=Replace(Temp2,"<$ShowButton$>","<div class=""BttnE"" onclick=""TopicShow(this,'log_"&LogID&"')""></div>")
	     Temp2=Replace(Temp2,"<$ShowStyle$>"," style=""display:none""")
	   else
	     Temp2=Replace(Temp2,"<$ShowButton$>","")
	     Temp2=Replace(Temp2,"<$ShowStyle$>","")
     end if
 
     Temp1=Replace(Temp1,"<$log_weather$>",log_View("log_weather"))
     Temp1=Replace(Temp1,"<$log_level$>",log_View("log_level"))
     Temp1=Replace(Temp1,"<$log_Author$>",log_View("log_Author"))
     Temp1=Replace(Temp1,"<$log_IsShow$>",log_View("log_IsShow"))
     if log_View("log_IsShow") then 
	     Temp1=Replace(Temp1,"<$log_hiddenIcon$>","")
	     Temp2=Replace(Temp2,"<$log_hiddenIcon$>","")
     else
	     Temp1=Replace(Temp1,"<$log_hiddenIcon$>","<img src=""images/icon_lock.gif"" style=""margin:0px 0px -3px 2px;"" alt="""" />")
	     Temp2=Replace(Temp2,"<$log_hiddenIcon$>","<img src=""images/icon_lock.gif"" style=""margin:0px 0px -3px 2px;"" alt="""" />")
     end if
     
     if len(log_View("log_tag"))>0 then 
	     Temp1=Replace(Temp1,"<$log_tag$>",getTags.filterHTML(log_View("log_tag")))
	     Temp2=Replace(Temp2,"<$log_tag$>","<p>Tags: "&getTags.filterHTML(log_View("log_tag"))&"</p>")
     else
	     Temp1=Replace(Temp1,"<$log_tag$>","")
	     Temp2=Replace(Temp2,"<$log_tag$>","")
     end if
	 if log_View("log_comorder") then comDesc="Desc" else comDesc="Asc" end if
     Temp1=Replace(Temp1,"<$comDesc$>",comDesc)
     Temp1=Replace(Temp1,"<$log_DisComment$>",log_View("log_DisComment"))
     
  	 if log_View("log_edittype")=1 then
  	 	Temp1=Replace(Temp1,"<$ArticleContent$>",UnCheckStr(UBBCode(HtmlEncode(log_View("log_Content")),mid(log_View("log_ubbFlags"),1,1),mid(log_View("log_ubbFlags"),2,1),mid(log_View("log_ubbFlags"),3,1),mid(log_View("log_ubbFlags"),4,1),mid(log_View("log_ubbFlags"),5,1))))
	    Temp2=Replace(Temp2,"<$log_Intro$>",UnCheckStr(UBBCode(log_View("log_Intro"),mid(log_View("log_ubbFlags"),1,1),mid(log_View("log_ubbFlags"),2,1),mid(log_View("log_ubbFlags"),3,1),mid(log_View("log_ubbFlags"),4,1),mid(log_View("log_ubbFlags"),5,1))))
	  	if log_View("log_Intro")<>HtmlEncode(log_View("log_Content")) then
			Temp2=Replace(Temp2,"<$log_readMore$>","<p><a href=""article.asp?id="&LogID&""" class=""more"">查看更多...</a></p>")
		  else
			Temp2=Replace(Temp2,"<$log_readMore$>","")
		end if
  	 else
  		Temp1=Replace(Temp1,"<$ArticleContent$>",UnCheckStr(log_View("log_Content")))
	    Temp2=Replace(Temp2,"<$log_Intro$>",UnCheckStr(log_View("log_Intro")))
		if log_View("log_Intro")<>log_View("log_Content") then
			Temp2=Replace(Temp2,"<$log_readMore$>","<p><a href=""article.asp?id="&LogID&""" class=""more"">查看更多...</a></p>")
		  else
			Temp2=Replace(Temp2,"<$log_readMore$>","")
		end if	    
  	 end if
  	
  	
  	
     if len(log_View("log_Modify"))>0 then 
       Temp1=Replace(Temp1,"<$log_Modify$>",log_View("log_Modify")&"<br/>")
      else
       Temp1=Replace(Temp1,"<$log_Modify$>","")
     end If
     
     Temp1=Replace(Temp1,"<$log_FromUrl$>",log_View("log_FromUrl"))
     Temp1=Replace(Temp1,"<$log_From$>",log_View("log_From"))
     Temp1=Replace(Temp1,"<$trackback$>",SiteURL&"trackback.asp?tbID="&LogID&"&amp;action=view")
     
     Temp1=Replace(Temp1,"<$log_CommNums$>",log_View("log_CommNums"))
     Temp1=Replace(Temp1,"<$log_QuoteNums$>",log_View("log_QuoteNums"))
     
     Temp2=Replace(Temp2,"<$log_CommNums$>",log_View("log_CommNums"))
     Temp2=Replace(Temp2,"<$log_QuoteNums$>",log_View("log_QuoteNums"))
     
     
     Temp1=Replace(Temp1,"<$log_IsDraft$>",log_View("log_IsDraft"))
	 
	 set preLogC=Conn.Execute("SELECT TOP 1 log_Title,log_ID FROM blog_Content WHERE log_PostTime<#"&DateToStr(log_View("log_PostTime"),"Y-m-d H:I:S")&"# and log_IsDraft=false ORDER BY log_PostTime DESC")
	 set nextLogC=Conn.Execute("SELECT TOP 1 log_Title,log_ID FROM blog_Content WHERE log_PostTime>#"&DateToStr(log_View("log_PostTime"),"Y-m-d H:I:S")&"# and log_IsDraft=false ORDER BY log_PostTime ASC")
     
     dim BTemp
     BTemp=""
		 if not preLogC.eof then
				 BTemp = BTemp & "<a href=""?id="&preLogC("log_ID")&""" title=""上一篇日志: "&preLogC("log_Title")&""" accesskey="",""><img border=""0"" src=""images/Cprevious.gif"" alt=""""/>上一篇</a>"
			else
				 BTemp = BTemp & "<img border=""0"" src=""images/Cprevious1.gif"" alt=""这是最新一篇日志""/>上一篇"
		 end if
		 if not nextLogC.eof then
				BTemp = BTemp & " | <a href=""?id="&nextLogC("log_ID")&""" title=""下一篇日志: "&nextLogC("log_Title")&""" accesskey="".""><img border=""0"" src=""images/Cnext.gif"" alt=""""/>下一篇</a>"
			else
				BTemp = BTemp & " | <img border=""0"" src=""images/Cnext1.gif"" alt=""这是最后一篇日志""/>下一篇"
		 end if
	 Temp1=Replace(Temp1,"<$log_Navigation$>",BTemp)

     SaveArticle=SaveToFile(Temp1,"post/" & LogID & ".asp")
	 SaveArticle=SaveToFile(Temp2,"cache/" & LogID & ".asp")
     set getCate=nothing
     set getTags=nothing
     'getCate.cate_Secret or (not log_View("Log_IsShow"))
   end if
 end sub
%>