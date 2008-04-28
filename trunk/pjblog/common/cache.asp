<%
'***************PJblog2 缓存处理*******************
' PJblog2 Copyright 2006
' Update:2006-1-25
'**************************************************

'-------------------------Blog基本参数--------------------------
Dim blog_Infos,SiteName,SiteUrl,blogPerPage,blog_LogNums,blog_CommNums,blog_MemNums
Dim blog_VisitNums,blogBookPage,blog_MessageNums,blogcommpage,blogaffiche
Dim blogabout,blogcolsize,blog_colNums,blog_TbCount,blog_showtotal,blog_commTimerout
Dim blog_commUBB,blog_commImg,blog_version,blog_UpdateDate,blog_DefaultSkin,blog_SkinName,blog_SplitType
Dim blog_ImgLink,blog_postFile,blog_postCalendar,log_SplitType,blog_introChar,blog_introLine
Dim blog_validate,Register_UserNames,Register_UserName,FilterIPs,FilterIP,blog_Title
Dim blog_commLength,blog_downLocal,blog_DisMod,blog_Disregister,blog_master,blog_email,blog_CountNum
Dim blog_wapNum,blog_wapImg,blog_wapHTML,blog_wapLogin,blog_wapComment,blog_wap,blog_wapURL

'=========================日志基本信息缓存=======================
Sub getInfo(ByVal action)
	Dim blog_Infos
	'--------------写入基本信息缓存------------------
	  IF Not IsArray(Application(CookieName&"_blog_Infos")) or action=2 Then
		Dim log_Infos
		SQL="select top 1 blog_Name,blog_URL,blog_PerPage,blog_LogNums,blog_CommNums,blog_MemNums," & _
		    "blog_VisitNums,blog_BookPage,blog_MessageNums,blog_commPage,blog_affiche," & _
		    "blog_about,blog_colPage,blog_colNums,blog_tbNums,blog_showtotal," & _
		    "blog_FilterName,blog_FilterIP,blog_commTimerout,blog_commUBB,blog_commImg," & _
		    "blog_postFile,blog_postCalendar,blog_DefaultSkin,blog_SkinName,blog_SplitType," & _
		    "blog_introChar,blog_introLine,blog_validate,blog_Title,blog_ImgLink," & _
		    "blog_commLength,blog_downLocal,blog_DisMod,blog_Disregister,blog_master,blog_email,blog_CountNum," & _
		    "blog_wapNum,blog_wapImg,blog_wapHTML,blog_wapLogin,blog_wapComment,blog_wap,blog_wapURL" & _
		    " from blog_Info"
		Set log_Infos=Conn.Execute(SQL)
		SQLQueryNums=SQLQueryNums+1
		blog_Infos=log_Infos.GetRows()
		Set log_Infos=nothing
		Application.Lock
		Application(CookieName&"_blog_Infos")=blog_Infos
		Application.UnLock
	   Else
		blog_Infos=Application(CookieName&"_blog_Infos")
	   End IF
	   
	'--------------读取基本信息缓存------------------
	if action<>2 then
	    SiteName=blog_Infos(0,0)'站点名字
	    SiteURL=blog_Infos(1,0)'站点地址
	    blogPerPage=int(blog_Infos(2,0))'每页日志数
	    blog_LogNums=int(blog_Infos(3,0))'日志总数
	    blog_CommNums=int(blog_Infos(4,0))'评论总数
	    blog_MemNums=int(blog_Infos(5,0))'会员总数
	    blog_VisitNums=int(blog_Infos(6,0))'访问量
	    blogBookPage=int(blog_Infos(7,0))'每页留言数(备用)
	    blog_MessageNums=int(blog_Infos(8,0))'留言总数(备用)
	    blogcommpage=int(blog_Infos(9,0))'每页评论数
	    blogaffiche=blog_Infos(10,0)'公告
	    blogabout=blog_Infos(11,0)'备案信息
	    blogcolsize=int(blog_Infos(12,0))'每页书签数(备用)
	    blog_colNums=int(blog_Infos(13,0))'书签总数(备用)
	    blog_TbCount=int(blog_Infos(14,0))'引用通告总数
	    blog_showtotal=CBool(blog_Infos(15,0))'是否显示统计(备用)
	    Register_UserNames=blog_Infos(16,0)'注册名字过滤
	    Register_UserName=Split(Register_UserNames,"|")
	    FilterIPs=blog_Infos(17,0)'IP地址过滤
	    FilterIP=Split(FilterIPs,"|")
	    blog_commTimerout=int(blog_Infos(18,0))'发表评论时间间隔
	    blog_commUBB=int(blog_Infos(19,0))'是否禁用评论UBB代码
	    blog_commIMG=int(blog_Infos(20,0))'是否禁用评论贴图
	    blog_postFile=CBool(blog_Infos(21,0)) '动态输出日志文件
	    blog_postCalendar=CBool(blog_Infos(22,0)) '动态输出日志日历文件
	    blog_DefaultSkin=blog_Infos(23,0)'默认界面
	    blog_SkinName=blog_Infos(24,0)'界面名称
	    blog_SplitType=CBool(blog_Infos(25,0))'日志分割类型
	    blog_introChar=blog_Infos(26,0)'日志预览最大字符数
	    blog_introLine=blog_Infos(27,0)'日志预览切割行数
	    blog_validate=CBool(blog_Infos(28,0))'发表评论是否都需要验证
		blog_Title=blog_Infos(29,0)'Blog副标题
		blog_ImgLink=CBool(blog_Infos(30,0))'是否在首页显示图片友情链接
		blog_commLength=int(blog_Infos(31,0))'评论长度
		blog_downLocal=CBool(blog_Infos(32,0))'是否使用防盗链下载
		blog_DisMod=CBool(blog_Infos(33,0))'默认显示内容
		blog_Disregister=CBool(blog_Infos(34,0))'是否允许注册
		blog_master=blog_Infos(35,0)'blog管理员姓名
		blog_email=blog_Infos(36,0)'blog管理员邮件地址
		blog_CountNum=blog_Infos(37,0)'访客统计最大次数
		
	    blog_wapNum=int(blog_Infos(38,0))'Wap 文章列表数量
	    blog_wapImg=CBool(blog_Infos(39,0))'Wap 文章显示图片
	    blog_wapHTML=CBool(blog_Infos(40,0))'Wap 文章使用简单HTML
	    blog_wapLogin=CBool(blog_Infos(41,0))'Wap 允许登录
	    blog_wapComment=CBool(blog_Infos(42,0))'Wap 允许评论	
	    blog_wap=CBool(blog_Infos(43,0))'使用 wap
	    blog_wapURL=CBool(blog_Infos(44,0))'使用 wap 转换文章超链接
	    
	    blog_version="2.7 Build 05"'当前PJBlog版本号
	    blog_UpdateDate="2007-11-09"'PJBlog最新更新时间

	 end if  
End Sub
'======================End Sub=======================


'-------------------------Blog权限变量---------------
Dim stat_title,stat_AddAll,stat_EditAll,stat_DelAll,stat_Add,stat_Edit,stat_Del,stat_CommentAdd
Dim stat_CommentDel,stat_Admin,stat_code,UP_FileType,UP_FileSize,UP_FileTypes,stat_FileUpLoad
Dim stat_CommentEdit,stat_ShowHiddenCate

'=====================日志权限缓存===================
Sub UserRight(ByVal action) '读取日志权限
Dim blog_Status
   '--------------写入日志权限缓存------------------
  IF Not IsArray(Application(CookieName&"_blog_rights")) or action=2 Then
		Dim log_Status,log_StatusList
		SQL="select stat_name,stat_title,stat_Code,stat_attSize,stat_attType from blog_status"
		Set log_Status=Conn.Execute(SQL)
		SQLQueryNums=SQLQueryNums+1
	    blog_Status=log_Status.GetRows()
		Set log_Status=Nothing
		
		Application.Lock
		Application(CookieName&"_blog_rights")=blog_Status
		Application.UnLock
   Else
		blog_Status=Application(CookieName&"_blog_rights")
   End IF
   
	'--------------写入日志权限缓存------------------
  if action<>2 then
	   Dim blog_Status_Len,i
	   blog_Status_Len=ubound(blog_Status,2)
			For i=0 to blog_Status_Len
			  if blog_Status(0,i)=memStatus then
		            stat_title=blog_Status(1,i)
		            FillRight blog_Status(2,i)
		            UP_FileSize=blog_Status(3,i)
		            UP_FileTypes=blog_Status(4,i)
		            UP_FileType=Split(UP_FileTypes,"|")
		            'exit Sub
			  end if
			Next
  end if
End Sub

sub FillRight(StatusCode) '写入权限变量
	 stat_AddAll=CBool(mid(StatusCode,1,1))
	 stat_Add=CBool(mid(StatusCode,2,1))
	 stat_EditAll=CBool(mid(StatusCode,3,1))
	 stat_Edit=CBool(mid(StatusCode,4,1))
	 stat_DelAll=CBool(mid(StatusCode,5,1))
	 stat_Del=CBool(mid(StatusCode,6,1))
	 stat_CommentAdd=CBool(mid(StatusCode,7,1))
	 stat_CommentEdit=CBool(mid(StatusCode,8,1))
	 stat_CommentDel=CBool(mid(StatusCode,9,1))
	 stat_FileUpLoad=CBool(mid(StatusCode,10,1))
	 stat_Admin=CBool(mid(StatusCode,11,1))
	 stat_ShowHiddenCate=CBool(mid(StatusCode,12,1))
end sub
'=========================End Sub========================



'========================日志分类缓存=========================
Dim Category_code
Sub CategoryList(ByVal action) '日志分类
'写入日志分类
'action=0 横向菜单 action=1 树状菜单 action=2重建分类

	'--------------写入日志分类缓存------------------
	 Dim Arr_Category,i
	 IF Not IsArray(Application(CookieName&"_blog_Category")) or action=2 Then
	 	 Dim log_Category
		 TempVar=""
		 SQL="SELECT cate_ID,cate_Name,cate_Order,cate_Intro,cate_OutLink,cate_URL,cate_icon,cate_count,cate_Lock,cate_local,cate_Secret FROM blog_Category ORDER BY cate_Order ASC"
		 Set log_Category=Conn.Execute(SQL)
		 SQLQueryNums=SQLQueryNums+1
		 if log_Category.eof or log_Category.bof then
			  ReDim Arr_Category(0,0)
		 else
			  Arr_Category=log_Category.GetRows()
		 end if
		 Set log_Category=Nothing
		 Application.Lock
		 Application(CookieName&"_blog_Category")=Arr_Category
		 Application.UnLock
	 Else
		 Arr_Category=Application(CookieName&"_blog_Category")
	 End IF

	Dim Category_Len,Menu_Diver
	'--------------输出日志横向菜单------------------
	if action=0 then
		    Menu_Diver=""
			Response.Write("<div id=""menu""><div id=""Left""></div><div id=""Right""></div><ul><li class=""menuL""></li>")

		    if ubound(Arr_Category,1)=0 then Response.Write("<li class=""menuR""></li></ul></div>"):exit Sub
		    
		    Category_Len=ubound(Arr_Category,2)
		    
			For i=0 to Category_Len
				    if int(Arr_Category(9,i))=0 or int(Arr_Category(9,i))=1 then
				    Response.Write(Menu_Diver)
				 	  if Arr_Category(4,i) then
				 	     if cbool(Arr_Category(10,i)) then
					 	     if stat_ShowHiddenCate or stat_Admin then Response.Write("<li><a class=""menuA"" href="""&Arr_Category(5,i)&""" title="""&Arr_Category(3,i)&""">"&Arr_Category(1,i)&"</a></li>")
				 	      else
						     Response.Write("<li><a class=""menuA"" href="""&Arr_Category(5,i)&""" title="""&Arr_Category(3,i)&""">"&Arr_Category(1,i)&"</a></li>")
				 	     end if
				  	  else
				 	     if cbool(Arr_Category(10,i)) then
					 	     if stat_ShowHiddenCate or stat_Admin then Response.Write("<li><a class=""menuA"" href=""default.asp?cateID="&Arr_Category(0,i)&""" title="""&Arr_Category(3,i)&""">"&Arr_Category(1,i)&"</a></li>")
				 	      else
						     Response.Write("<li><a class=""menuA"" href=""default.asp?cateID="&Arr_Category(0,i)&""" title="""&Arr_Category(3,i)&""">"&Arr_Category(1,i)&"</a></li>")
				 	     end if
				      end if
					   Menu_Diver="<li class=""menuDiv""></li>"
				    end if
			Next
			
			Response.Write("<li class=""menuR""></li></ul></div>")
	end if

	if action=1 then
			Category_code=""
			if ubound(Arr_Category,1)=0 then exit Sub
			    
			Category_Len=ubound(Arr_Category,2)
			    
			For i=0 to Category_Len
			    if int(Arr_Category(9,i))=0 or int(Arr_Category(9,i))=2 then
				 if Arr_Category(4,i) then
					  if cbool(Arr_Category(10,i)) then
						   if stat_ShowHiddenCate or stat_Admin then Category_code=Category_code&("<img src="""&Arr_Category(6,i)&""" border=""0"" style=""margin:3px 4px -4px 0px;"" alt="""&Arr_Category(3,i)&"""/><a class=""CategoryA"" href="""&Arr_Category(5,i)&""" title="""&Arr_Category(3,i)&""">"&Arr_Category(1,i)&"</a><br/>")
					   else
						   Category_code=Category_code&("<img src="""&Arr_Category(6,i)&""" border=""0"" style=""margin:3px 4px -4px 0px;"" alt="""&Arr_Category(3,i)&"""/><a class=""CategoryA"" href="""&Arr_Category(5,i)&""" title="""&Arr_Category(3,i)&""">"&Arr_Category(1,i)&"</a><br/>")
					  end if
				 else
					  if cbool(Arr_Category(10,i)) then
						   if stat_ShowHiddenCate or stat_Admin then Category_code=Category_code&("<img src="""&Arr_Category(6,i)&""" border=""0"" style=""margin:3px 4px -4px 0px;"" alt="""&Arr_Category(3,i)&"""/><a class=""CategoryA"" href=""default.asp?cateID="&Arr_Category(0,i)&""" title="""&Arr_Category(3,i)&""">"&Arr_Category(1,i)&" ["&Arr_Category(7,i)&"]</a> <a href=""feed.asp?cateID="&Arr_Category(0,i)&""" title=""订阅该分类内容""><img src=""images/rss.png"" border=""0"" style=""margin:3px 4px -1px 0px;"" alt=""""/></a><br/>")
					   else
						   Category_code=Category_code&("<img src="""&Arr_Category(6,i)&""" border=""0"" style=""margin:3px 4px -4px 0px;"" alt="""&Arr_Category(3,i)&"""/><a class=""CategoryA"" href=""default.asp?cateID="&Arr_Category(0,i)&""" title="""&Arr_Category(3,i)&""">"&Arr_Category(1,i)&" ["&Arr_Category(7,i)&"]</a> <a href=""feed.asp?cateID="&Arr_Category(0,i)&""" title=""订阅该分类内容""><img src=""images/rss.png"" border=""0"" style=""margin:3px 4px -1px 0px;"" alt=""""/></a><br/>")
					  end if
				 end if
				end if
			Next
	end if
End Sub
'========================End Sub===============================


'========================日志归档缓存============================
function archive(ByVal action)'日志归档
	Dim blog_archive,i
	'-----------------写入日志归档缓存--------------------
	IF Not IsArray(Application(CookieName&"_blog_archive")) or action=2 Then
		Dim log_archives
		SQL="SELECT Count(log_ID) AS [count], Year([log_PostTime]) AS PostYear, Month([log_PostTime]) AS PostMonth " &_
	        "FROM blog_Content where blog_Content.log_IsDraft=false "&_
	        "GROUP BY Year([log_PostTime]), Month([log_PostTime]) "&_
	        "ORDER BY Year([log_PostTime]) Desc, Month([log_PostTime]) ASC"
		Set log_archives=Conn.Execute(SQL)
		SQLQueryNums=SQLQueryNums+1
		if log_archives.eof or log_archives.bof then
		    ReDim blog_archive(0,0)
		else
			blog_archive=log_archives.GetRows()
		end if
		Set log_archives=Nothing

		Application.Lock
		Application(CookieName&"_blog_archive")=blog_archive
		Application.UnLock
	Else
		blog_archive=Application(CookieName&"_blog_archive")
	End IF
	
	'-----------------读取日志归档缓存--------------------
	if action<>2 then
  Dim archive_item_Len,Month_array,TempYear,MonthCounter
  if ubound(blog_archive,1)=0 then archive="":exit function
  Month_array=Array("01月","02月","03月","04月","05月","06月","07月","08月","09月","10月","11月","12月")
  archive_item_Len=ubound(blog_archive,2)
  TempYear=blog_archive(1,0)
  MonthCounter=0
   For i=0 to archive_item_Len
    IF i=0 Then archive="<a class=""sideA"" style=""margin:0px 0px 0px -2px;"" href=""default.asp?log_Year="&blog_archive(1,i)&""" title=""查看"&blog_archive(1,i)&"年的日志"">"&blog_archive(1,i)&"</a>"
    IF blog_archive(1,i)=TempYear Then
   archive=archive&"<a style=""margin-right:5px;"" href=""default.asp?log_Year="&blog_archive(1,i)&"&log_Month="&blog_archive(2,i)&""" title="""&blog_archive(1,i)&"年"&blog_archive(2,i)&"月有"&blog_archive(0,i)&"篇日志"">"&Month_array(blog_archive(2,i)-1)&"</a>"
   MonthCounter=MonthCounter+1
   IF MonthCounter=5 Then MonthCounter=0:archive=archive&"<br/>"
  Else
   MonthCounter=1
   archive=archive&"<a class=""sideA"" style=""margin:6px 0px 0px -2px;"" href=""default.asp?log_Year="&blog_archive(1,i)&""" title=""查看"&blog_archive(1,i)&"年的日志"">"&blog_archive(1,i)&"</a>"
   archive=archive&"<a style=""margin-right:5px;"" href=""default.asp?log_Year="&blog_archive(1,i)&"&log_Month="&blog_archive(2,i)&""" title="""&blog_archive(1,i)&"年"&blog_archive(2,i)&"月有"&blog_archive(0,i)&"篇日志"">"&Month_array(blog_archive(2,i)-1)&"</a>"
   TempYear=blog_archive(1,i)
  End IF
   Next
 end if
end function
'=====================End Function========================



'=====================最新评论缓存=====================
function NewComment(ByVal action)
	Dim blog_Comment,ShowLen,i
	ShowLen=10 '显示最新评论预览数量
	'-----------------写入最新评论缓存--------------------
	IF Not IsArray(Application(CookieName&"_blog_Comment")) or action=2 Then
		Dim log_Comments
		SQL="SELECT top "&ShowLen&" comm_ID,blog_ID,comm_Author,comm_Content,comm_PostTime" &_
		    " FROM blog_Comment as C,blog_Content as T,blog_Category as A where C.blog_ID=T.log_ID and T.log_IsShow=true and T.log_CateID=A.cate_ID and A.cate_Secret=false order by C.comm_PostTime Desc"
		Set log_Comments=Conn.Execute(SQL)
		SQLQueryNums=SQLQueryNums+1
		if log_Comments.eof or log_Comments.bof then
			    ReDim blog_Comment(0,0)
			else
				blog_Comment=log_Comments.GetRows(ShowLen)
			end if
		Set log_Comments=Nothing
		Application.Lock
		Application(CookieName&"_blog_Comment")=blog_Comment
		Application.UnLock
	Else
		blog_Comment=Application(CookieName&"_blog_Comment")
	End IF

	'-----------------读取最新评论缓存--------------------
	 if action<>2 then
		  dim Comment_Item_Len
		  if ubound(blog_Comment,1)=0 then NewComment="":exit function
		  Comment_Item_Len=ubound(blog_Comment,2)
		  For i=0 to Comment_Item_Len
				 NewComment=NewComment&"<a class=""sideA"" href=""default.asp?id="&blog_Comment(1,i)&"#comm_"&blog_Comment(0,i)&""" title="""&blog_Comment(2,i)&" 于 "&blog_Comment(4,i)&" 发表评论"&CHR(10)&CCEncode(CutStr(DelQuote(blog_Comment(3,i)),100))&""">"&CCEncode(CutStr(DelQuote(blog_Comment(3,i)),25))&"</a>"
		  Next
		 end if
end function
'=====================End Function========================

'====================写入标签Tag缓存=====================
Dim Arr_Tags
function Tags(ByVal action)
 IF Not IsArray(Application(CookieName&"_blog_Tags")) or action=2 Then
	Dim log_Tags,log_TagsList
	Set log_Tags=Conn.Execute("SELECT tag_id,tag_name,tag_count FROM blog_tag")
	SQLQueryNums=SQLQueryNums+1
	TempVar=""
	Do While Not log_Tags.EOF
		log_TagsList=log_TagsList&TempVar&log_Tags("tag_id")&"||"&log_Tags("tag_name")&"||"&log_Tags("tag_count")
		TempVar=","
		log_Tags.MoveNext
	Loop
	Set log_Tags=Nothing
	Arr_Tags=Split(log_TagsList,",")
	Application.Lock
	Application(CookieName&"_blog_Tags")=Arr_Tags
	Application.UnLock
 Else
	Arr_Tags=Application(CookieName&"_blog_Tags")
 End IF
end Function
'======================End Function========================

'====================写入表情符号缓存=====================
Dim Arr_Smilies
function Smilies(ByVal action)
 IF Not IsArray(Application(CookieName&"_blog_Smilies")) or action=2 Then
	Dim log_Smilies,log_SmiliesList
	Set log_Smilies=Conn.Execute("SELECT sm_ID,sm_Image,sm_Text FROM blog_Smilies")
	SQLQueryNums=SQLQueryNums+1
	TempVar=""
	Do While Not log_Smilies.EOF
		log_SmiliesList=log_SmiliesList&TempVar&log_Smilies("sm_ID")&"|"&log_Smilies("sm_Image")&"|"&log_Smilies("sm_Text")
		TempVar=","
		log_Smilies.MoveNext
	Loop
	Set log_Smilies=Nothing
	Arr_Smilies=Split(log_SmiliesList,",")
	Application.Lock
	Application(CookieName&"_blog_Smilies")=Arr_Smilies
	Application.UnLock
 Else
	Arr_Smilies=Application(CookieName&"_blog_Smilies")
 End IF
end Function
'======================End Function========================

'======================写入关键字列表======================
Dim Arr_Keywords
function Keywords(ByVal action)
IF Not IsArray(Application(CookieName&"_blog_Keywords")) or action=2 Then
	Dim log_Keywords,log_KeywordsList
	Set log_Keywords=Conn.Execute("SELECT key_ID,key_Text,key_URL,key_Image FROM blog_Keywords")
	SQLQueryNums=SQLQueryNums+1
	TempVar=""
	Do While Not log_Keywords.EOF
		IF log_Keywords("key_Image")<>Empty Then
			log_KeywordsList=log_KeywordsList&TempVar&log_Keywords("key_ID")&"$|$"&log_Keywords("key_Text")&"$|$"&log_Keywords("key_URL")&"$|$"&log_Keywords("key_Image")
		Else
			log_KeywordsList=log_KeywordsList&TempVar&log_Keywords("key_ID")&"$|$"&log_Keywords("key_Text")&"$|$"&log_Keywords("key_URL")&"$|$None"
		End IF
		TempVar="|$|"
		log_Keywords.MoveNext
	Loop
	Set log_Keywords=Nothing
	Arr_Keywords=Split(log_KeywordsList,"|$|")
	Application.Lock
	Application(CookieName&"_blog_Keywords")=Arr_Keywords
	Application.UnLock
Else
	Arr_Keywords=Application(CookieName&"_blog_Keywords")
End IF
end function 
'======================End Function=========================

'=======================写入首页链接列表====================
Dim Arr_Bloglinks
function Bloglinks(ByVal action)
IF Not IsArray(Application(CookieName&"_blog_Bloglinks")) or action=2 Then
	Dim log_Bloglinks,log_BloglinksList
	Set log_BlogLinks=Conn.ExeCute("SELECT link_Name,link_URL,link_Image FROM blog_Links WHERE link_IsMain=True ORDER BY link_Order ASC")
	SQLQueryNums=SQLQueryNums+1
	TempVar=""
	Do While Not log_BlogLinks.EOF
		IF log_BlogLinks("link_Image")<>Empty Then
			log_BloglinksList=log_BloglinksList&TempVar&log_BlogLinks("link_Name")&"$|$"&log_BlogLinks("link_URL")&"$|$"&log_BlogLinks("link_Image")
		Else
			log_BloglinksList=log_BloglinksList&TempVar&log_BlogLinks("link_Name")&"$|$"&log_BlogLinks("link_URL")&"$|$None"
		End IF
		TempVar="|$|"
		log_BlogLinks.MoveNext
	Loop
	Set log_BlogLinks=Nothing
	Arr_Bloglinks=Split(log_BloglinksList,"|$|")
	Application.Lock
	Application(CookieName&"_blog_Bloglinks")=Arr_Bloglinks
	Application.UnLock
Else
	Arr_Bloglinks=Application(CookieName&"_blog_Bloglinks")
End IF

if action=1 then
   Dim Arr_Bloglink,Arr_BloglinkItem,ImgLink,TextLink
   Bloglinks=""
   for each Arr_Bloglink in Arr_Bloglinks
    Arr_BloglinkItem=Split(Arr_Bloglink,"$|$")
    if blog_ImgLink then
      if Arr_BloglinkItem(2)="None" then
       TextLink=TextLink&"<a class=""sideA"" href="""&Arr_BloglinkItem(1)&""" target=""_blank"" title="""&Arr_BloglinkItem(0)&""">"&Arr_BloglinkItem(0)&"</a>"
      else
       ImgLink=ImgLink&"<a href="""&Arr_BloglinkItem(1)&""" target=""_blank"" title="""&Arr_BloglinkItem(0)&"""><img src="""&Arr_BloglinkItem(2)&""" border=""0"" alt=""""/></a> "
      end if
     else
      Bloglinks=Bloglinks&"<a class=""sideA"" href="""&Arr_BloglinkItem(1)&""" target=""_blank"" title="""&Arr_BloglinkItem(0)&""">"&Arr_BloglinkItem(0)&"</a>"
    end if
   next
   if blog_ImgLink then Bloglinks=ImgLink&TextLink
end if
end function 
'=====================End Function=======================


'======================自定义模块缓存=====================
Dim side_html_default,side_html,content_html_Top_default,content_html_Top,content_html_Bottom_default,content_html_Bottom,function_Plugin

function log_module(ByVal action)
Dim blog_modules
side_html_default="" '首页侧栏代码
side_html="" '普通页面侧栏代码
content_html_Top_default="" '首页内容代码顶部
content_html_Top="" '普通页面内容代码顶部
content_html_Bottom_default="" '首页内容代码底部
content_html_Bottom="" '普通页面内容代码底部
function_Plugin="" 'Blog功能插件

IF Not IsArray(Application(CookieName&"_blog_module")) or action=2 Then
    dim blog_module,blog_module_array,TempDiv
    TempDiv=""
    SQL="SELECT type,title,name,HtmlCode,IndexOnly,SortID,PluginPath,InstallFolder FROM blog_module where IsHidden=false order by SortID"
    Set blog_module=Conn.ExeCute(SQL)
    SQLQueryNums=SQLQueryNums+1
          do until blog_module.eof 
	             if blog_module("type")="sidebar" then
	                side_html_default=side_html_default&"<div id=""Side_"&blog_module("name")&""" class=""sidepanel"">"
	                if len(blog_module("title"))>0 then side_html_default=side_html_default&"<h4 class=""Ptitle"">"&blog_module("title")&"</h4>"
	                side_html_default=side_html_default&"<div class=""Pcontent"">"&blog_module("HtmlCode")&"</div><div class=""Pfoot""></div></div>"
	              if blog_module("IndexOnly")=false then
	                side_html=side_html&"<div id=""Side_"&blog_module("name")&""" class=""sidepanel"">"
	               if len(blog_module("title"))>0 then side_html=side_html&"<h4 class=""Ptitle"">"&blog_module("title")&"</h4>"
	                side_html=side_html&"<div class=""Pcontent"">"&blog_module("HtmlCode")&"</div><div class=""Pfoot""></div></div>"
	              end if
	             end if
	             if blog_module("type")="content" and blog_module("name")<>"ContentList" then
	               if blog_module("SortID")<0 then
	                    content_html_Top_default=content_html_Top_default&"<div id=""Content_"&blog_module("name")&""" class=""content-width"">"&blog_module("HtmlCode")&"</div>"
	                  if blog_module("IndexOnly")=false then
	                    content_html_Top=content_html_Top&"<div id=""Content_"&blog_module("name")&""" class=""content-width"">"&blog_module("HtmlCode")&"</div>"
	                  end if
	               else
	                    content_html_Bottom_default=content_html_Bottom_default&"<div id=""Content_"&blog_module("name")&""" class=""content-width"">"&blog_module("HtmlCode")&"</div>"
	                  if blog_module("IndexOnly")=false then
	                    content_html_Bottom=content_html_Bottom&"<div id=""Content_"&blog_module("name")&""" class=""content-width"">"&blog_module("HtmlCode")&"</div>"
	                  end if
	               end if
	             end if
	             if blog_module("type")="function" then
	              function_Plugin=function_Plugin&TempDiv&blog_module("name")&"%|%"&blog_module("PluginPath")&"%|%"&blog_module("InstallFolder")
	              TempDiv="$*$"
	             end if
              blog_module.movenext
          loop
    Set blog_module=Nothing
    blog_modules=array(side_html_default,side_html,content_html_Top_default,content_html_Top,content_html_Bottom_default,content_html_Bottom,function_Plugin)
	Application.Lock
	Application(CookieName&"_blog_module")=blog_modules
	Application.UnLock
Else
	blog_modules=Application(CookieName&"_blog_module")
End IF

if action<>2 then
     side_html_default=UnCheckStr(blog_modules(0)) '首页侧栏代码
     side_html=UnCheckStr(blog_modules(1)) '普通页面侧栏代码
     content_html_Top_default=UnCheckStr(blog_modules(2)) '首页内容代码顶部
     content_html_Top=UnCheckStr(blog_modules(3)) '普通页面内容代码顶部
     content_html_Bottom_default=UnCheckStr(blog_modules(4)) '首页内容代码底部
     content_html_Bottom=UnCheckStr(blog_modules(5)) '普通页面内容代码底部
     function_Plugin=blog_modules(6) 'Blog功能插件
end if

end function
'========================End function=========================


'======================重新加载Blog缓存=====================
sub reloadcache
 getInfo(2)
 UserRight(2)
 CategoryList(2)
 archive(2)
 NewComment(2)
 Tags(2)
 Smilies(2)
 Keywords(2)
 Bloglinks(2)
 log_module(2)
 Calendar "","","",2
end sub
%>