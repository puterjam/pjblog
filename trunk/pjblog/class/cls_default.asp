<%
'==================================
'  系统首页类文件
'    更新时间: 2006-1-22
'==================================

'**********************************************
'日志列表处理
'**********************************************

Function ContentList()'日志列表
    Dim webLog, webLogArr, webLogArrLen, Log_Num, PageCount, CanRead, ViewType, ViewDraft, strSQL, ViewTag, Readpw
    Dim getCate, ArticleList
    PageCount = 0
    Set getCate = New Category
    ViewDraft = checkstr(Request.QueryString("display"))
    ViewTag = checkstr(Request.QueryString("tag"))
    CanRead = False


    If Len(checkstr(Request.QueryString("distype")))>0 Then
        Response.Cookies(CookieNameSetting)("ViewType") = checkstr(Request.QueryString("distype"))
    Else
        If Len(Request.Cookies(CookieNameSetting)("ViewType"))<1 Then
            If blog_DisMod Then
                Response.Cookies(CookieNameSetting)("ViewType") = "list"
            Else
                Response.Cookies(CookieNameSetting)("ViewType") = "normal"
            End If
        End If
    End If

    Dim CT
    CT = ""
    If IsInteger(cateID) = True Then
        getCate.load(cateID)
        CT = "分类: "&getCate.cate_Name&""
        If getCate.cate_Secret Then
            If Not stat_ShowHiddenCate And Not stat_Admin Then
%>
					   <div style="margin:10px 0px 10px 0px"><strong>抱歉，没有找到任何日志！</strong></div>
				   <%
				Exit Function
			End If
		End If
	End If

If Request.Cookies(CookieNameSetting)("ViewType") = "list" Then ViewType = "list" Else ViewType = "normal"

If ViewType = "list" Then
    strSQL = "log_ID,log_CateID,log_Author,log_Title,log_PostTime,log_IsShow,log_CommNums,log_QuoteNums,log_ViewNums,log_IsTop,log_Readpw,log_Pwtitle"
Else
    strSQL = "log_ID,log_CateID,log_Author,log_Title,log_PostTime,log_IsShow,log_CommNums,log_QuoteNums,log_ViewNums,log_IsTop,log_Intro,log_Content,log_edittype,log_DisComment,log_ubbFlags,log_tag,log_Readpw,log_Pwtitle"
End If

'row序号: 0     ,1         ,2         ,3        ,4           ,5         ,6           ,7            ,8           ,9        ,10       ,11         ,12          ,13            ,14          ,15
If Len(ViewTag)>0 Then
    Dim getTag, getTID
    Set getTag = New tag
    getTID = getTag.getTagID(ViewTag)
    If getTID<>0 Then
        SQLFiltrate = SQLFiltrate & " log_tag LIKE '%{"&getTID&"}%' AND "
        Url_Add = Url_Add & "tag="&Server.URLEncode(ViewTag)&"&amp;"
        CT = "Tag: "&ViewTag&""
    End If
    Set getTag = Nothing
End If

'=================Load Cache List============================
Set ArticleList = New ArticleCache
If ArticleList.loadCache And Len(ViewTag)<1 And IsInteger(log_Year) = False And IsInteger(log_Month) = False And IsInteger(log_Day) = False And ViewDraft<>"draft" Then
    If IsInteger(cateID) = True Then
        ArticleList.outHTML "C"&cateID, ViewType, CT
    Else
        If stat_Admin Or stat_ShowHiddenCate Then
            ArticleList.outHTML "A", ViewType, CT
        Else
            ArticleList.outHTML "G", ViewType, CT
        End If
    End If
    Exit Function
End If

'=================Load DB List===============================
If stat_ShowHiddenCate Or stat_Admin Then
    SQL = "SELECT "&strSQL&" FROM blog_Content "&SQLFiltrate&" log_IsDraft=false ORDER BY log_IsTop ASC,log_PostTime DESC"
Else
    SQL = "SELECT "&strSQL&" FROM blog_Content As T,blog_Category As C "&SQLFiltrate&" T.log_CateID=C.cate_ID and C.cate_Secret=false and log_IsDraft=false ORDER BY log_IsTop ASC,log_PostTime DESC"
End If

'if stat_ShowHiddenCate or stat_Admin then

If ViewDraft = "draft" And Len(memName)>0 Then
    ViewType = "list"
    SQL = "SELECT "&strSQL&" FROM blog_Content "&SQLFiltrate&" log_IsDraft=true and log_Author='"&memName&"' ORDER BY log_IsTop ASC,log_PostTime DESC"
End If
Set webLog = Server.CreateObject("Adodb.Recordset")

webLog.Open SQL, CONN, 1, 1
SQLQueryNums = SQLQueryNums + 1

If webLog.EOF Or webLog.BOF Then
    If ViewDraft = "draft" Then
%>
					   <div style="margin:10px 0px 10px 0px"><strong>抱歉，没有找到任何草稿！</strong></div>
				   <%else%>
					   <div style="margin:10px 0px 10px 0px"><strong>抱歉，没有找到任何日志！</strong></div>
			   <%End If
Exit Function
Else
    If ViewDraft = "draft" Then Url_Add = Url_Add&"display=draft&"
    If ViewType = "list" Then blogPerPage = blogPerPage * 4
    webLog.PageSize = blogPerPage
    webLog.AbsolutePage = CurPage
    Log_Num = webLog.RecordCount
    If webLog.EOF Or webLog.BOF Then
%>
				   <div style="margin:10px 0px 10px 0px"><strong>抱歉，没有找到任何日志！</strong></div>
			   <%
Exit Function
End If
webLogArr = webLog.GetRows(Log_Num)
webLog.Close
Set webLog = Nothing
webLogArrLen = UBound(webLogArr, 2)
If ViewDraft = "draft" Then
%>
		      <div class="pageContent" style="text-align:Right;overflow:hidden;height:18px;line-height:140%"><span style="float:left">草稿列表</span><%=MultiPage(Log_Num,blogPerPage,CurPage,Url_Add,"","float:Left","")%></div>
		   <%else%>
		      <div class="pageContent" style="text-align:Right;overflow:hidden;height:18px;line-height:140%"><span style="float:left"><%=CT%></span>预览模式: <a href="<%=Url_Add%>distype=normal" accesskey="1">普通</a> | <a href="<%=Url_Add%>distype=list" accesskey="2">列表</a></div>
		   <%End If
If ViewType = "list" Then
%>
		     <div class="Content-body" style="text-align:Left"><table cellpadding="2" cellspacing="2" width="100%">
		   <%End If
Do Until PageCount = webLogArrLen + 1 Or PageCount = blogPerPage
    If IsInteger(cateID) = False Then
        getCate.load(webLogArr(1, PageCount))
    End If
    '是否有权限查看日记
    If ViewType="list" Then
	Readpw=Trim(webLogArr(10,PageCount))
    Else
	Readpw=Trim(webLogArr(16,PageCount))
    End If
    If stat_Admin = True Then CanRead = True
    If webLogArr(5, PageCount) Then CanRead = True
    If webLogArr(5, PageCount) = False And webLogArr(2, PageCount) = memName Then CanRead = True
    If Readpw<>"" and Session("ReadPassWord_"&webLogArr(0,PageCount)) = Readpw then CanRead = True
    
    If ViewType = "list" Then
        '====================================
        '  列表模式
        '====================================
        OutList webLogArr, PageCount, getCate, ViewDraft, CanRead
    Else
        '====================================
        '  正常模式
        '====================================
        OutNomal webLogArr, PageCount, getCate, CanRead
    End If
    PageCount = PageCount + 1
    CanRead = False
Loop
If ViewType = "list" Then
%>
		 </table></div>
		<%end if%>
		 <div class="pageContent"><%=MultiPage(Log_Num,blogPerPage,CurPage,Url_Add,"","","")%></div>
		<%End If
End Function


' ----------------------- 输出普通模式--------------------

Function OutNomal(webLogArr, PageCount, getCate, CanRead)
    If getCate.cate_Secret Then
        If Not stat_ShowHiddenCate And Not stat_Admin Then Exit Function
    End If
    Dim getTag,aUrl
    Set getTag = New tag
	If blog_postFile>1 Then
		aUrl = caload(webLogArr(0,PageCount))
	else
		aUrl = "article.asp?id=" & webLogArr(0,PageCount)
	end if
%>
		<div class="Content">
		<div class="Content-top"><div class="ContentLeft"></div><div class="ContentRight"></div>
		<%If webLogArr(9,PageCount)=True Then%>
		 <div class="BttnE" onclick="TopicShow(this,'log_<%=webLogArr(0,PageCount)%>')"></div>
		<%end if%>
		 <h1 class="ContentTitle"><img src="<%=getCate.cate_icon%>" style="margin:0px 2px -4px 0px;" alt="" class="CateIcon"/>
		<%If CanRead Then%>
			<a class="titleA" href="<%=aUrl%>"><%=HtmlEncode(webLogArr(3,PageCount))%></a>
		<%Else%>
			<a class="titleA" href="article.asp?id=<%=webLogArr(0,PageCount)%>"><%If webLogArr(17,PageCount) = False then%><%=HtmlEncode(webLogArr(3,PageCount))%><%ElseIf Trim(webLogArr(16,PageCount)) <> "" Then%>[加密日志]<%Else%>[私密日志]<%End If%></a>
		<%End If
			If webLogArr(5, PageCount) = False Or getCate.cate_Secret Then
			%>
			<%If Trim(webLogArr(16,PageCount)) <> "" Then%><img src="images/icon_lock2.gif" style="margin:0px 0px -3px 2px;" alt="加密日志" /><%Else%><img src="images/icon_lock1.gif" style="margin:0px 0px -3px 2px;" alt="私密日志" /><%End If%>
			<%end if%>
		</h1>
		<h2 class="ContentAuthor">作者:<%=webLogArr(2,PageCount)%>&nbsp; 日期:<%=DateToStr(webLogArr(4,PageCount),"Y-m-d")%></h2></div>
		  <div id="log_<%=webLogArr(0,PageCount)%>"<%if webLogArr(9,PageCount)=true then %> style="display:none"<%end if%>>
		<%
If CanRead Then
    If webLogArr(12, PageCount) = 1 Then
%>
					<div class="Content-body"><%=UnCheckStr(UBBCode(webLogArr(10,PageCount),mid(webLogArr(14,PageCount),1,1),mid(webLogArr(14,PageCount),2,1),mid(webLogArr(14,PageCount),3,1),mid(webLogArr(14,PageCount),4,1),mid(webLogArr(14,PageCount),5,1)))%>
					<%if webLogArr(10,PageCount)<>HtmlEncode(webLogArr(11,PageCount)) then%>
						<p class="readMore"><a href="<%=aUrl%>" class="more"><span>查看更多...</span></a></p>
					<%end if%>
			<%else%>
					<div class="Content-body"><%=UnCheckStr(webLogArr(10,PageCount))%>
					<%if webLogArr(10,PageCount)<>webLogArr(11,PageCount) then%>
						<p class="readMore"><a href="default.asp?id=<%=webLogArr(0,PageCount)%>" class="more"><span>查看更多...</span></a></p>
					<%End If
End If
If Len(webLogArr(15, PageCount))>0 Then

%>
			 <p class="tags">Tags: <%=getTag.filterHTML(webLogArr(15,PageCount))%></p>
			<%
End If
Else
%>
			<div class="Content-body">
			<%if Trim(webLogArr(16,PageCount))<>"" then%>
			该日志是加密日志，需要输入正确密码才可以查看！
			<%else%>
			该日志是私密日志，只有管理员或发布者可以查看！
			<%end if%>
		<%end if%>
							 
		</div><div class="Content-bottom">
		<div class="ContentBLeft"></div><div class="ContentBRight"></div>
		分类:<a href="default.asp?cateID=<%=webLogArr(1,PageCount)%>" title="<%=getCate.cate_Intro%>"><%=getCate.cate_Name%></a> | <a href="?id=<%=webLogArr(0,PageCount)%>">固定链接</a> |
							
		<%if webLogArr(13,PageCount)=true then%>
			 禁止评论 
		<%Else%>
			 <a href="<%=aUrl%>#comm_top">评论: <%=webLogArr(6,PageCount)%></a>
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
Set getTag = Nothing
End Function


' ----------------------- 输出列表模式 --------------------

Function OutList(webLogArr, PageCount, getCate, ViewDraft, CanRead)
    Dim logLink, logIcon,aUrl
    If getCate.cate_Secret Then
        If Not stat_ShowHiddenCate And Not stat_Admin Then Exit Function
    End If
	If blog_postFile>1 Then
		aUrl = caload(webLogArr(0,PageCount))
	else
		aUrl = "article.asp?id=" & webLogArr(0,PageCount)
	end if
%>
		<tr><td valign="top">
		<%If ViewDraft = "draft" Then
    logLink = "blogedit.asp?id="&webLogArr(0, PageCount)
    logIcon = "<a href=""blogedit.asp?id="&webLogArr(0, PageCount)&""" title=""编辑草稿""><img border=""0"" alt=""编辑草稿"" src=""images/drafts.gif"" style=""margin:0px 4px -2px 0px""/></a>"
Else
    logLink = aUrl
    logIcon = "<a href=""default.asp?cateID="&webLogArr(1, PageCount)&""" ><img border=""0"" alt=""查看 "&getCate.cate_Name&" 的日志"" src="""&getCate.cate_icon&""" style=""margin:0px 2px -3px 0px""/></a>"
End If

If webLogArr(9, PageCount) Then
%><b><%end If%>
		<%=logIcon%>
		<%If CanRead Then%>
			<a href="<%=logLink%>" title="作者:<%=webLogArr(2,PageCount)%> 日期:<%=DateToStr(webLogArr(4,PageCount),"Y-m-d")%>"><%=HtmlEncode(webLogArr(3,PageCount))%></a>
		<%Else%>
			<a href="<%=logLink%>"><%if webLogArr(11,PageCount)=False then%><%=HtmlEncode(webLogArr(3,PageCount))%><%ElseIf Trim(webLogArr(10,PageCount)) <> "" Then%>[加密日志]<%Else%>[私密日志]<%End If%></a>
		<%End If

If webLogArr(5, PageCount) = False Or getCate.cate_Secret Then
%>
		<%If Trim(webLogArr(10,PageCount)) <> "" Then%><img src="images/icon_lock2.gif" style="margin:0px 0px -3px 2px;" alt="加密日志"/><%Else%><img src="images/icon_lock1.gif" style="margin:0px 0px -3px 2px;" alt="私密日志"/><%End If%>

		<%end if%>
		</td>
		<%If webLogArr(9,PageCount) Then %></b><%end If%>
		<%If not ViewDraft="draft" then %>
			<td valign="top" width="60"><nobr><a href="<%=aUrl%>#comm_top" title="评论"><%=webLogArr(6,PageCount)%></a> | <span title="引用通告"><%=webLogArr(7,PageCount)%></span> | <span title="阅读次数"><%=webLogArr(8,PageCount)%></span></nobr></td>
		<%else%>
		    <td valign="top" width="60"><nobr><%=webLogArr(2,PageCount)%></span></nobr></td>
		<%end if%>
		</tr>
<%end function%>
