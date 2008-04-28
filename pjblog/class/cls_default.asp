<%
'==================================
'  系统首页类文件
'    更新时间: 2006-1-22
'==================================

'**********************************************
'日志列表处理
'**********************************************
function ContentList()'日志列表
		Dim webLog,webLogArr,webLogArrLen,Log_Num,PageCount,CanRead,ViewType,ViewDraft,strSQL,ViewTag
		Dim getCate,ArticleList
		PageCount=0
		set getCate=new Category
		ViewDraft=checkstr(Request.QueryString("display"))
		ViewTag=checkstr(Request.QueryString("tag"))
		CanRead=False

 		
		If len(checkstr(Request.QueryString("distype")))>0 Then
		     Response.Cookies(CookieNameSetting)("ViewType")=checkstr(Request.QueryString("distype"))
		else
		  if len(Request.Cookies(CookieNameSetting)("ViewType"))<1 then
		     if blog_DisMod then
		        Response.Cookies(CookieNameSetting)("ViewType")="list" 
		      else
		        Response.Cookies(CookieNameSetting)("ViewType")="normal"
		     end if
		  end if
		end if
		
		Dim CT 
		   CT=""
		   IF IsInteger(cateID)=True Then
			 getCate.load(cateID)
			 CT="分类: "&getCate.cate_Name&" | "
		 	 if getCate.cate_Secret then
				   if not stat_ShowHiddenCate and not stat_Admin then  %>
					   <div style="margin:10px 0px 10px 0px"><strong>抱歉，没有找到任何日志！</strong></div>
				   <%
				   exit function
			 	 end if
		 	  end if
	    End If

		If Request.Cookies(CookieNameSetting)("ViewType")="list" Then ViewType="list" Else ViewType="normal"

		if ViewType="list" then
		  strSQL="log_ID,log_CateID,log_Author,log_Title,log_PostTime,log_IsShow,log_CommNums,log_QuoteNums,log_ViewNums,log_IsTop"
		 else
		  strSQL="log_ID,log_CateID,log_Author,log_Title,log_PostTime,log_IsShow,log_CommNums,log_QuoteNums,log_ViewNums,log_IsTop,log_Intro,log_Content,log_edittype,log_DisComment,log_ubbFlags,log_tag"
		end if
		
		'row序号: 0     ,1         ,2         ,3        ,4           ,5         ,6           ,7            ,8           ,9        ,10       ,11         ,12          ,13            ,14          ,15
		If len(ViewTag)>0 then
			dim getTag,getTID
			set getTag=new tag
			getTID=getTag.getTagID(ViewTag)
			if getTID<>0 then 
			   SQLFiltrate=SQLFiltrate & " log_tag LIKE '%{"&getTID&"}%' AND "
			   Url_Add=Url_Add & "tag="&Server.URLEncode(ViewTag)&"&amp;"
			   CT="Tag: "&ViewTag&" | "
			end if
			set getTag=nothing
		end if
	  
'=================Load Cache List============================
	    set ArticleList=new ArticleCache
		if ArticleList.loadCache and len(ViewTag)<1 and IsInteger(log_Year)=false and IsInteger(log_Month)=false and IsInteger(log_Day)=false and ViewDraft<>"draft" then
		  if IsInteger(cateID)=True then
			  ArticleList.outHTML "C"&cateID,ViewType,CT
		  else
		    if stat_Admin Or stat_ShowHiddenCate then
			  ArticleList.outHTML "A",ViewType,CT
			 else
			  ArticleList.outHTML "G",ViewType,CT
			end if
		  end if
		  exit function
		end if
		
		
'=================Load DB List===============================
	  if stat_ShowHiddenCate or stat_Admin then
		SQL="SELECT "&strSQL&" FROM blog_Content "&SQLFiltrate&" log_IsDraft=false ORDER BY log_IsTop ASC,log_PostTime DESC"
	   else
		SQL="SELECT "&strSQL&" FROM blog_Content As T,blog_Category As C "&SQLFiltrate&" T.log_CateID=C.cate_ID and C.cate_Secret=false and log_IsDraft=false ORDER BY log_IsTop ASC,log_PostTime DESC"
	  end if
	  
		'if stat_ShowHiddenCate or stat_Admin then

			If ViewDraft="draft" and len(memName)>0 then
				ViewType="list"
				SQL="SELECT "&strSQL&" FROM blog_Content "&SQLFiltrate&" log_IsDraft=true and log_Author='"&memName&"' ORDER BY log_IsTop ASC,log_PostTime DESC"
			end if
		Set webLog=Server.CreateObject("Adodb.Recordset")

		webLog.Open SQL,CONN,1,1
		SQLQueryNums=SQLQueryNums+1

		If webLog.EOF or webLog.BOF Then 
			   If ViewDraft="draft" then%>
					   <div style="margin:10px 0px 10px 0px"><strong>抱歉，没有找到任何草稿！</strong></div>
				   <%else%>
					   <div style="margin:10px 0px 10px 0px"><strong>抱歉，没有找到任何日志！</strong></div>
			   <%end if
			   exit function
		 Else
		   If ViewDraft="draft" then Url_Add=Url_Add&"display=draft&"
		   If ViewType="list" Then blogPerPage=blogPerPage*4
		   webLog.PageSize=blogPerPage
		   webLog.AbsolutePage=CurPage
		   Log_Num=webLog.RecordCount
		   if webLog.EOF or webLog.BOF then %>
				   <div style="margin:10px 0px 10px 0px"><strong>抱歉，没有找到任何日志！</strong></div>
			   <%
			   exit function
		   end if
		   webLogArr=webLog.GetRows(Log_Num)
		   webLog.close
		   Set webLog=Nothing
		   webLogArrLen=Ubound(webLogArr,2)
		   If ViewDraft="draft" then%>
		      <div class="pageContent" style="text-align:Right;overflow:hidden;height:18px;line-height:140%"><span style="float:left">草稿列表</span><%=MultiPage(Log_Num,blogPerPage,CurPage,Url_Add,"","float:Left")%></div>
		   <%else%>
		      <div class="pageContent" style="text-align:Right;overflow:hidden;height:18px;line-height:140%"><span style="float:left"><%=CT%></span><%=MultiPage(Log_Num,blogPerPage,CurPage,Url_Add,"","float:Left")%> 预览模式: <a href="<%=Url_Add%>distype=normal" accesskey="1">普通</a> | <a href="<%=Url_Add%>distype=list" accesskey="2">列表</a></div>
		   <%end if
		   If ViewType="list" Then%>
		     <div class="Content-body" style="text-align:Left"><table cellpadding="2" cellspacing="2" width="100%">
		   <%end if
		   Do Until PageCount=webLogArrLen+1 or PageCount=blogPerPage
			    IF IsInteger(cateID)=False Then
					getCate.load(webLogArr(1,PageCount))
			    end if
				'是否有权限查看日记
				If stat_Admin=True then CanRead=True
				If webLogArr(5,PageCount) Then CanRead=True
				If webLogArr(5,PageCount)=False And webLogArr(2,PageCount)=memName then CanRead=True
					If ViewType="list" Then
							 '====================================
							 '  列表模式
							 '====================================
							 OutList webLogArr,PageCount,getCate,ViewDraft,CanRead
						 Else
							 '====================================
							 '  正常模式
							 '====================================
							 OutNomal webLogArr,PageCount,getCate,CanRead
						End If	
				PageCount=PageCount+1
		        CanRead=False
		   Loop
		If ViewType="list" Then%>
		 </table></div>
		<%end if%>
		 <div class="pageContent"><%=MultiPage(Log_Num,blogPerPage,CurPage,Url_Add,"","float:Left")%></div>
		<%end If
end Function

		
' ----------------------- 输出普通模式--------------------
function OutNomal(webLogArr,PageCount,getCate,CanRead)
		if getCate.cate_Secret then
		  if not stat_ShowHiddenCate and not stat_Admin then exit function
		end if
		dim getTag
		set getTag=new tag
		%>
		<div class="Content">
		<div class="Content-top"><div class="ContentLeft"></div><div class="ContentRight"></div>
		<%If webLogArr(9,PageCount)=True Then%>
		 <div class="BttnE" onclick="TopicShow(this,'log_<%=webLogArr(0,PageCount)%>')"></div>
		<%end if%>
		 <h1 class="ContentTitle"><img src="<%=getCate.cate_icon%>" style="margin:0px 2px -4px 0px;" alt="" class="CateIcon"/>
		<%If CanRead Then%>
			<a class="titleA" href="article.asp?id=<%=webLogArr(0,PageCount)%>"><%=HtmlEncode(webLogArr(3,PageCount))%></a>
		<%Else%>
			<a class="titleA" href="article.asp?id=<%=webLogArr(0,PageCount)%>">[隐藏日志]</a>
		<%end If
		if webLogArr(5,PageCount)=false or getCate.cate_Secret then %>
			<img src="images/icon_lock.gif" style="margin:0px 0px -3px 2px;" alt="" />
		<%end if%>
		</h1>
		<h2 class="ContentAuthor">作者:<%=webLogArr(2,PageCount)%>&nbsp; 日期:<%=DateToStr(webLogArr(4,PageCount),"Y-m-d")%></h2></div>
		  <div id="log_<%=webLogArr(0,PageCount)%>"<%if webLogArr(9,PageCount)=true then %> style="display:none"<%end if%>>
		<%
		if CanRead then 
			if webLogArr(12,PageCount)=1 then%>
					<div class="Content-body"><%=UnCheckStr(UBBCode(webLogArr(10,PageCount),mid(webLogArr(14,PageCount),1,1),mid(webLogArr(14,PageCount),2,1),mid(webLogArr(14,PageCount),3,1),mid(webLogArr(14,PageCount),4,1),mid(webLogArr(14,PageCount),5,1)))%>
					<%if webLogArr(10,PageCount)<>HtmlEncode(webLogArr(11,PageCount)) then%>
						<p><a href="article.asp?id=<%=webLogArr(0,PageCount)%>" class="more">查看更多...</a></p>
					<%end if%>
			<%else%>
					<div class="Content-body"><%=UnCheckStr(webLogArr(10,PageCount))%>
					<%if webLogArr(10,PageCount)<>webLogArr(11,PageCount) then%>
						<p><a href="default.asp?id=<%=webLogArr(0,PageCount)%>" class="more">查看更多...</a></p>
					<%end if
			end if
			if len(webLogArr(15,PageCount))>0 then
			%>
			 <p>Tags: <%=getTag.filterHTML(webLogArr(15,PageCount))%></p>
			<%
			end if
		Else%>
			<div class="Content-body">该日志是隐藏日志，只有管理员或发布者可以查看！
		<%end if%>
							 
		</div><div class="Content-bottom">
		<div class="ContentBLeft"></div><div class="ContentBRight"></div>
		分类:<a href="default.asp?cateID=<%=webLogArr(1,PageCount)%>" title="<%=getCate.cate_Intro%>"><%=getCate.cate_Name%></a> | <a href="?id=<%=webLogArr(0,PageCount)%>">固定链接</a> |
							
		<%if webLogArr(13,PageCount)=true then%>
			 禁止评论 
		<%Else%>
			 <a href="article.asp?id=<%=webLogArr(0,PageCount)%>#comm_top">评论: <%=webLogArr(6,PageCount)%></a>
		<%end If%>
			 | 引用: <%=webLogArr(7,PageCount)%> | 查看次数: <%=webLogArr(8,PageCount)%>
				<%if stat_EditAll or (stat_Edit and webLogArr(2,PageCount)=memName) then%> 
					 | <a href="blogedit.asp?id=<%=webLogArr(0,PageCount)%>"><img src="images/icon_edit.gif" alt="" border="0" style="margin-bottom:-2px"/></a>
				<%end if%>
				<%if stat_DelAll or (stat_Del and webLogArr(2,PageCount)=memName)  then%> 
					 | <a href="blogedit.asp?action=del&amp;id=<%=webLogArr(0,PageCount)%>" onclick="if (!window.confirm('是否要删除该日志')) return false"><img src="images/icon_del.gif" alt="" border="0" style="margin-bottom:-2px"/></a>
				<%end if%>				 
			   </div>
			</div></div>
<%
    set getTag=nothing
end function


' ----------------------- 输出列表模式 --------------------
function OutList(webLogArr,PageCount,getCate,ViewDraft,CanRead)
		dim logLink,logIcon
		if getCate.cate_Secret then
		  if not stat_ShowHiddenCate and not stat_Admin then exit function
		end if%>
		<tr><td valign="top">
		<%If ViewDraft="draft" then 
			logLink="blogedit.asp?id="&webLogArr(0,PageCount)
			logIcon="<a href=""blogedit.asp?id="&webLogArr(0,PageCount)&""" title=""编辑草稿""><img border=""0"" alt=""编辑草稿"" src=""images/drafts.gif"" style=""margin:0px 4px -2px 0px""/></a>"
		else
			logLink="article.asp?id="&webLogArr(0,PageCount)
			logIcon="<a href=""default.asp?cateID="&webLogArr(1,PageCount)&""" ><img border=""0"" alt=""查看 "&getCate.cate_Name&" 的日志"" src="""&getCate.cate_icon&""" style=""margin:0px 2px -3px 0px""/></a>"
		end if
							 
		If webLogArr(9,PageCount) Then %><b><%end If%>
		<%=logIcon%>
		<%If CanRead Then%>
			<a href="<%=logLink%>" title="作者:<%=webLogArr(2,PageCount)%> 日期:<%=DateToStr(webLogArr(4,PageCount),"Y-m-d")%>"><%=HtmlEncode(webLogArr(3,PageCount))%></a>
		<%Else%>
			<a href="<%=logLink%>">[隐藏日志]</a>
		<%end If
								 
		if webLogArr(5,PageCount)=false or getCate.cate_Secret then %>
			<img src="images/icon_lock.gif" style="margin:0px 0px -3px 2px;" alt=""/>
		<%end if%>
		</td>
		<%If webLogArr(9,PageCount) Then %></b><%end If%>
		<%If not ViewDraft="draft" then %>
			<td valign="top" width="60"><nobr><a href="article.asp?id=<%=webLogArr(0,PageCount)%>#comm_top" title="评论"><%=webLogArr(6,PageCount)%></a> | <span title="引用通告"><%=webLogArr(7,PageCount)%></span> | <span title="阅读次数"><%=webLogArr(8,PageCount)%></span></nobr></td>		
		<%else%>
		    <td valign="top" width="60"><nobr><%=webLogArr(2,PageCount)%></span></nobr></td>		
		<%end if%>
		</tr>
<%end function%>