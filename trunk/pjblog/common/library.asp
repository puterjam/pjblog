<%
'***************PJblog2 模块与类处理*******************
' PJblog2 Copyright 2005
' Update:2005-10-20
'**************************************************

'**********************************************
'BLOG日历
'**********************************************
function Calendar(C_Year,C_Month,C_Day,update)  
	Dim C_YM,S_Date,E_Date,isDo,Dclass,RS_Month,Link_TF,i,DayCount,DayStr,NotCM,TS_Date,TE_Date
	IF C_Year=Empty Then C_Year=Year(Now())
	IF C_Month=Empty Then C_Month=Month(Now())
	IF C_Day=Empty Then C_Day=0
	C_Year=Cint(C_Year)
	C_Month=Cint(C_Month)
	C_Day=Cint(C_Day)
	C_YM=C_Year & "-" & C_Month
    dim PY,PM,NY,NM
	PM=C_Month-1
	if PM<1 then PM=12:PY=C_Year-1 else PY=C_Year
 	NM=C_Month+1
	if NM>12 then NM=1:NY=C_Year+1 else NY=C_Year
    Calendar="<div id=""Calendar_Body""><div id=""Calendar_Top""><div id=""LeftB"" onclick=""location.href='default.asp?log_Year="&PY&"&amp;log_Month="&PM&"'""></div><div id=""RightB"" onclick=""location.href='default.asp?log_Year="&NY&"&amp;log_Month="&NM&"'""></div>"&C_Year&"年"&C_Month&"月</div>"
    Calendar=Calendar & "<div id=""Calendar_week""><ul class=""Week_UL""><li><font color=""#FF0000"">日</font></li><li>一</li><li>二</li><li>三</li><li>四</li><li>五</li><li>六</li></ul></div>"

'--->计算当前月份的日期 
    i=weekday(C_YM & "-" & 1)-1
    TS_Date=DateSerial(C_Year,C_Month,1-i)
	TE_Date=DateAdd("d",42,TS_Date)
    S_Date=year(TS_Date)&"-"&month(TS_Date)&"-"&day(TS_Date)
    E_Date=year(TE_Date)&"-"&month(TE_Date)&"-"&day(TE_Date)

'--->保存日志日历缓存
Dim Link_Count,Link_Days,CalendarArray,doUpdate,upTime
upTime=Year(Now())&"-"&Month(Now())
doUpdate=false

	if Not IsArray(Application(CookieName&"_blog_Calendar")) then '判断日期更新条件
		doUpdate=true
    elseif Application(CookieName&"_blog_Calendar")(1)<>upTime then
		doUpdate=true
    elseif upTime<>C_Year&"-"&C_Month then
		doUpdate=true
    elseif update=2 then
		doUpdate=true
    end if

	if doUpdate then
		ReDim Link_Days(4,0)
		Link_Count=0
		SQL="SELECT C.log_id,C.log_title,C.log_PostTime,C.log_IsShow FROM blog_Content as C,blog_Category as A where C.log_PostTime Between #"&S_Date&" 00:00:00# And #"&E_Date&" 23:59:59# and C.log_IsDraft=false and C.log_CateID=A.cate_ID and A.cate_Secret=false ORDER BY C.log_PostTime"
		Set RS_Month=Conn.Execute(SQL)
		SQLQueryNums=SQLQueryNums+1
		Dim the_Day,TempTitle,TempCount,TempSplit
		the_Day=0
		TempCount=0
		TempTitle=""
		Do While NOT RS_Month.EOF
			IF Day(RS_Month("log_PostTime"))<>the_Day Then
				the_Day=Day(RS_Month("log_PostTime"))
				ReDim PreServe Link_Days(4,Link_Count)
				Link_Days(0,Link_Count)=Year(RS_Month("log_PostTime"))
				Link_Days(1,Link_Count)=Month(RS_Month("log_PostTime"))
				Link_Days(2,Link_Count)=Day(RS_Month("log_PostTime"))
				Link_Days(3,Link_Count)="default.asp?log_Year="&Year(RS_Month("log_PostTime"))&"&amp;log_Month="&Month(RS_Month("log_PostTime"))&"&amp;log_Day="&Day(RS_Month("log_PostTime"))
			    TempCount=1
			    if RS_Month("log_IsShow") then
				    TempTitle=chr(13) & " - " & RS_Month("log_title")
				  else
				    TempTitle=chr(13) & " - [隐藏日志]"
				end if
				Link_Days(4,Link_Count)="当天共写了" & TempCount &"篇日志" & TempTitle
				Link_Count=Link_Count+1
			Else
			    TempCount=TempCount+1
			    if RS_Month("log_IsShow") then
				    Link_Days(4,Link_Count-1) = Link_Days(4,Link_Count-1) & chr(10) & " - " & RS_Month("log_title")
				  else
				    Link_Days(4,Link_Count-1) = Link_Days(4,Link_Count-1) & chr(10) & " - [隐藏日志]"
				end if
			    
			    TempSplit = split(Link_Days(4,Link_Count-1),chr(13))
			    TempSplit(0)="当天共写了" & TempCount &"篇日志" & chr(13)
			    if ubound(TempSplit)>0 then Link_Days(4,Link_Count-1)=TempSplit(0) & TempSplit(1)
			End IF
			RS_Month.MoveNext
		Loop
		Set RS_Month=Nothing
		 'response.write "<script>alert('动态日历')</script>"
		if upTime=C_Year&"-"&C_Month then
			CalendarArray=array(Link_Days,upTime)
			Application.Lock
			Application(CookieName&"_blog_Calendar")=CalendarArray
			Application.UnLock
		 'response.write "<script>alert('写日历缓存')</script>"
		end if
	 else
		Link_Days=Application(CookieName&"_blog_Calendar")(0)
		Link_Count=ubound(Link_Days,2)+1
		'response.write "<script>alert('静态日历')</script>"
	 end if	
	
	if update=2 then exit function
	
	dim DayEnd,Calendar_Count
	Calendar_Count=0
	DayEnd=false
    DayCount=0:Dclass="":DayStr="":isDo=0:NotCM=1
    do Until month(S_Date)<>C_Month and NotCM=7
	 	 if DayCount>6 then
			   Calendar=Calendar & "<div class=""Calendar_Day""><ul class=""Day_UL"">"&DayStr&"</ul></div>"
			   DayCount=0
		       DayStr=""
		 end if
		 if Calendar_Count=Link_Count then Calendar_Count=Link_Count-1:DayEnd=true
	     if month(S_Date)=C_Month then NotCM=0
	     if month(S_Date)<>C_Month then
		      Dclass="class=""otherday"""
			  NotCM=NotCM+1
		 elseif year(S_Date)=year(now()) and month(S_Date)=month(now()) and day(S_Date)=day(now()) then 
			  Dclass="class=""today"""
		 else
			  Dclass=""
		 end if
		  if Link_Count>0 then
			     if Link_Days(1,Calendar_Count)=month(S_Date) and Link_Days(2,Calendar_Count)=day(S_Date) and DayEnd=false then
					    if month(S_Date)<>C_Month then
					          Dclass="class=""otherday"""
				        elseif day(S_Date)=C_Day then
						     Dclass="class=""click"""
				        elseif C_Year=year(now()) and C_Month=month(now()) and day(S_Date)=day(now()) then 
							 Dclass="class=""DayD"""
						else
							 Dclass="class=""haveD"""
						end if
				      DayStr=DayStr&"<li class=""DayA""><a "&Dclass&" href="""&Link_Days(3,Calendar_Count)&""" title="""&Link_Days(4,Calendar_Count)&""">"&day(S_Date)&"</a></li>"  
				  	  Calendar_Count=Calendar_Count+1
			     else
				   	  DayStr=DayStr&"<li class=""DayA""><a "&Dclass&">"&day(S_Date)&"</a></li>"
			     end if
			 else
				   	  DayStr=DayStr&"<li class=""DayA""><a "&Dclass&">"&day(S_Date)&"</a></li>"
		  end if
		  DayCount=DayCount+1
		  S_Date=DateAdd("d",1,S_Date)       
    loop
    Calendar=Calendar & "</div>"
End function


'**********************************************
'用户面板
'**********************************************
function userPanel()
 userPanel=""
 if memName<>Empty then userPanel=userPanel&" <b>"&memName&"</b>，欢迎你!<br/>你的权限: "&stat_title&"<br/><br/>"
 if stat_Admin=true then userPanel=userPanel+"<a href=""control.asp"" class=""sideA"" accesskey=""3"">系统管理</a>"
 if stat_AddAll=true or stat_Add=true then userPanel=userPanel+"<a href=""blogpost.asp"" class=""sideA"" accesskey=""N"">发表新日志</a>"
 if (stat_AddAll=true or stat_Add=true) and (stat_EditAll or stat_Edit) then
  if isEmpty(session(CookieName&"_draft_"&memName)) then
   session(CookieName&"_draft_"&memName)=conn.execute("select count(log_ID) from blog_Content where log_Author='"&memName&"' and log_IsDraft=true")(0)
   SQLQueryNums=SQLQueryNums+1
  end if
  if session(CookieName&"_draft_"&memName)>0 then
      userPanel=userPanel+"<a href=""default.asp?display=draft"" class=""sideA"" accesskey=""D""><strong>编辑草稿 ["&session(CookieName&"_draft_"&memName)&"]</strong></a>"
   else
      userPanel=userPanel+"<a href=""default.asp?display=draft"" class=""sideA"" accesskey=""D"">编辑草稿</a>"
  end if
 end if
 if memName<>Empty then 
  userPanel=userPanel&"<a href=""member.asp?action=edit"" class=""sideA"" accesskey=""M"">修改个人资料</a><a href=""login.asp?action=logout"" class=""sideA"" accesskey=""Q"">退出系统</a>"
 else
  userPanel=userPanel&"<a href=""login.asp"" class=""sideA"" accesskey=""L"">登录</a><a href=""register.asp"" class=""sideA"" accesskey=""U"">用户注册</a>"
 end if  
end function


'**********************************************
'输出日志统计信息
'**********************************************
function info_code(str)
    dim vOnline
    vOnline=getOnline
    str=replace(str,"$blog_LogNums$",blog_LogNums)
    str=replace(str,"$blog_CommNums$",blog_CommNums)
    str=replace(str,"$blog_TbCount$",blog_TbCount)
    str=replace(str,"$blog_MessageNums$",blog_MessageNums)
    str=replace(str,"$blog_MemNums$",blog_MemNums)
    str=replace(str,"$blog_VisitNums$",blog_VisitNums)
    str=replace(str,"$blog_OnlineNums$",vOnline)
	info_code=str
end function

'**********************************************
'获取在线人数
'**********************************************
function getOnline
	getOnline=1
	if len(Application(CookieName&"_onlineCount"))>0 then
		  if DateDiff("s",Application(CookieName&"_userOnlineCountTime"),now())>60 then
			    Application.Lock()
			    Application(CookieName&"_online")=Application(CookieName&"_onlineCount")
			    Application(CookieName&"_onlineCount")=1
			    Application(CookieName&"_onlineCountKey")=randomStr(2)
			    Application(CookieName&"_userOnlineCountTime")=now()
			    Application.Unlock()
		  else
			    if Session(CookieName&"userOnlineKey")<>Application(CookieName&"_onlineCountKey") then
				      Application.Lock()
				      Application(CookieName&"_onlineCount")=Application(CookieName&"_onlineCount")+1
				      Application.Unlock()
				      Session(CookieName&"userOnlineKey")=Application(CookieName&"_onlineCountKey")
			    end if
		  end if
	else
		  Application.Lock
		  Application(CookieName&"_online")=1
		  Application(CookieName&"_onlineCount")=1
		  Application(CookieName&"_onlineCountKey")=randomStr(2)
		  Application(CookieName&"_userOnlineCountTime")=now()
		  Application.Unlock
	end if
	getOnline=Application(CookieName&"_online")
end function


'**********************************************
'侧边模版处理
'**********************************************
sub Side_Module_Replace()
'日历处理
    Dim Cal_code
	Cal_code=Calendar(log_Year,log_Month,log_Day,1)
    side_html_default=replace(side_html_default,"$calendar_code$",Cal_code)
    side_html=replace(side_html,"$calendar_code$",Cal_code)
    
'用户面板处理
    Dim user_code
    user_code=userPanel
    side_html_default=replace(side_html_default,"$user_code$",user_code)
    side_html=replace(side_html,"$user_code$",user_code)
    
'归档面板处理
    Dim archive_code
	archive_code=archive(1)
    side_html_default=replace(side_html_default,"$archive_code$",archive_code)
    side_html=replace(side_html,"$archive_code$",archive_code)
    
'树形分类处理
    CategoryList(1)
    side_html_default=replace(side_html_default,"$Category_code$",Category_code)
    side_html=replace(side_html,"$Category_code$",Category_code)

'显示统计信息
    side_html_default=info_code(side_html_default)
    side_html=info_code(side_html)

'处理最新评论内容
    Dim Comment_code
	Comment_code=NewComment(1)
    side_html_default=replace(side_html_default,"$comment_code$",Comment_code)
    side_html=replace(side_html,"$comment_code$",Comment_code)
    
'处理友情链接内容
    Dim Link_Code
	Link_Code=Bloglinks(1)
    side_html_default=replace(side_html_default,"$Link_Code$",Link_Code)
    side_html=replace(side_html,"$Link_Code$",Link_Code)
end sub




'==============================================================
' Blog Class
'==============================================================

'*******************************************
'  分类读取Class
'*******************************************
Class Category
	Public cate_ID
	Public cate_Name
	Public cate_Order
	Public cate_Intro
	Public cate_OutLink
	Public cate_URL
	Public cate_icon
	Public cate_count
	Public cate_Lock
	Public cate_local
	Public cate_Secret
	Private LastID
	Private Loaded
	
  Private Sub Class_Initialize()
		 cate_ID=0
		 cate_Name=""
		 cate_Order=0
		 cate_Intro=""
		 cate_OutLink=False
		 cate_URL=""
		 cate_icon=""
		 cate_count=""
		 cate_Lock=False
		 cate_local=""
		 cate_Secret=False
		 LastID=-99
		 Loaded=false
  end sub
  
  Private Sub Class_Terminate()

  end sub
  
 
  Public sub Reload
	 CategoryList(2) '更新分类缓存
  end sub

  Public function Load(ID)
   Dim blog_Cate,blog_CateArray,Category_Len,i
   if int(ID)=LastID then exit function
   if Not IsArray(Application(CookieName&"_blog_Category")) then Reload
   blog_CateArray=Application(CookieName&"_blog_Category")
   	if ubound(blog_CateArray,1)=0 then exit function
	Category_Len=ubound(blog_CateArray,2)
		For i=0 to Category_Len
			    if int(blog_CateArray(0,i))=int(ID) Then
				  cate_ID=blog_CateArray(0,i)
				  cate_Name=blog_CateArray(1,i)
				  cate_Order=blog_CateArray(2,i)
				  cate_Intro=blog_CateArray(3,i)
				  cate_OutLink=blog_CateArray(4,i)
				  cate_URL=blog_CateArray(5,i)
				  cate_icon=blog_CateArray(6,i)
				  cate_count=blog_CateArray(7,i)
				  cate_Lock=blog_CateArray(8,i)
				  cate_local=blog_CateArray(9,i)
				  cate_Secret=blog_CateArray(10,i)
				  LastID=int(ID)
				  Loaded=true
				  exit function
				end If
	   Next 
  end function
end Class


'*******************************************
'  Tag Class
'*******************************************
Class Tag

  Private Sub Class_Initialize()
	 IF Not IsArray(Application(CookieName&"_blog_Tags")) Then Reload
  end sub
  
  Private Sub Class_Terminate()

  end sub
  
  Public sub Reload
	 Tags(2) '更新Tag缓存
  end sub  
  
  
  Public function insert(tagName) '插入标签,返回ID号
   if checkTag(tagName) then
      conn.execute("update blog_tag set tag_count=tag_count+1 where tag_name='"&tagName&"'")
      insert=conn.execute("select top 1 tag_id from blog_tag where tag_name='"&tagName&"'")(0)
    else
      conn.execute("insert into blog_tag (tag_name,tag_count) values ('"&tagName&"',1)")
      insert=conn.execute("select top 1 tag_id from blog_tag order by tag_id desc")(0)
   end if
  end function
  
  
  Public function remove(tagID) '清除标签
   if checkTagID(tagID) then
     conn.execute("update blog_tag set tag_count=tag_count-1 where tag_id="&tagID)
  end if
  end function
  
  Public function filterHTML(str) '过滤标签
   	If isEmpty(str) Or isNull(str) Or len(str)=0 Then
        Exit Function
   		filterHTML=str
	 else
        dim log_Tag,log_TagItem
		For Each log_TagItem IN Arr_Tags
	   	    log_Tag=Split(log_TagItem,"||")
			str=replace(str,"{"&log_Tag(0)&"}","<a href=""default.asp?tag="&Server.URLEncode(log_Tag(1))&""">"&log_Tag(1)&"</a><a href=""http://technorati.com/tag/"&log_Tag(1)&""" rel=""tag"" style=""display:none"">"&log_Tag(1)&"</a> ")
		Next
		Dim re
		Set re=new RegExp
		re.IgnoreCase =True
		re.Global=True
        re.Pattern="\{(\d)\}"
      	str=re.Replace(str,"")
		filterHTML=str
	end if
  end function
  
  Public function filterEdit(str) '过滤标签进行编辑
   	If isEmpty(str) Or isNull(str) Or len(str)=0 Then
        Exit Function
   		filterEdit=str
	 else
        dim log_Tag,log_TagItem
		For Each log_TagItem IN Arr_Tags
	   	    log_Tag=Split(log_TagItem,"||")
			str=replace(str,"{"&log_Tag(0)&"}",log_Tag(1)&",")
		Next
		Dim re
		Set re=new RegExp
		re.IgnoreCase =True
		re.Global=True
        re.Pattern="\{(\d)\}"
      	str=re.Replace(str,"")
		filterEdit=left(str,len(str)-1)
	end if
  end function
 
  Private function checkTag(tagName) '检测是否存在此标签（根据名称）
   checkTag=false
   dim log_Tag,log_TagItem
	For Each log_TagItem IN Arr_Tags
		log_Tag=Split(log_TagItem,"||")
		if lcase(log_Tag(1))=lcase(tagName) then checkTag=true:exit function
	Next
  end function
  
  Private function checkTagID(tagID) '检测是否存在此标签（根据ID）
   checkTagID=false
   dim log_Tag,log_TagItem
	For Each log_TagItem IN Arr_Tags
		log_Tag=Split(log_TagItem,"||")
		if int(log_Tag(0))=int(tagID) then checkTagID=true:exit function
	Next
  end function
 
 Public function getTagID(tagName) '获得Tag的ID
   getTagID=0
   dim log_Tag,log_TagItem
	For Each log_TagItem IN Arr_Tags
		log_Tag=Split(log_TagItem,"||")
		if lcase(log_Tag(1))=lcase(ClearHTML(tagName)) then getTagID=log_Tag(0):exit function
	Next
 end function
end Class

%>