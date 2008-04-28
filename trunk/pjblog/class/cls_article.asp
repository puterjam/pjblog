<%
'==================================
'  日志类文件
'    更新时间: 2006-1-22
'==================================

'SQL="SELECT top 1 log_ID,log_CateID,log_title,Log_IsShow,log_ViewNums,log_Author,log_comorder,log_DisComment,log_Content,log_PostTime,log_edittype,log_ubbFlags,log_CommNums,log_QuoteNums,log_weather,log_level,log_Modify,log_FromUrl,log_From,log_tag FROM blog_Content WHERE log_ID="&id&" and log_IsDraft=false"
'row序号:          0     ,1         ,2         ,3        ,4           ,5         ,6           ,7             ,8          ,9           ,10          ,11          ,12           ,13         ,14         ,15       ,16        ,17         ,18       ,19

'*******************************************
'  显示日志内容
'*******************************************

	  sub updateViewNums(logID,vNums)
	   if not blog_postFile then exit sub
	   dim LoadArticle,splitStr,getA,i,tempStr
	   splitStr="<"&"%ST(A)%"&">"
	   tempStr=""
	   LoadArticle=LoadFromFile("cache/"&LogID&".asp")
	   if LoadArticle(0)=0 then
		   getA=split(LoadArticle(1),splitStr)
		   getA(2)=vNums
		   for i=1 to ubound(getA)
		     tempStr=tempStr&splitStr&getA(i)
		   next
		   call SaveToFile (tempStr,"cache/" & LogID & ".asp")	
   		end if
	  end sub


 sub ShowArticle(LogID)
         If (log_ViewArr(5,0)=memName And log_ViewArr(3,0)=False) Or stat_Admin or log_ViewArr(3,0)=true then 
           else
           showmsg "错误信息","该日志为隐藏日志，没有权限查看该日志！<br/><a href=""default.asp"">单击返回</a>","ErrorIcon",""
	     End if
	     If (Not getCate.cate_Secret) Or (log_ViewArr(5,0)=memName And getCate.cate_Secret) Or stat_Admin Or (getCate.cate_Secret and stat_ShowHiddenCate) Then
           else
           showmsg "错误信息","该日志分类为保密类型，无法查看该日志！<br/><a href=""default.asp"">单击返回</a>","ErrorIcon",""
	     end if 
	               
		 if log_ViewArr(6,0) then comDesc="Desc" else comDesc="Asc" end If
	          
	     '从文件读取日志
	     if blog_postFile then 
	       dim LoadArticle,TempStr,TempArticle
	       LoadArticle=LoadFromFile("post/"&LogID&".asp")
	       
	       if LoadArticle(0)=0 then
	            TempArticle=LoadArticle(1)
		        TempStr=""
		        if stat_EditAll or (stat_Edit and memName=log_ViewArr(5,0)) then 
			       TempStr=TempStr&"<a href=""blogedit.asp?id="&LogID&""" title=""编辑该日志"" accesskey=""E""><img src=""images/icon_edit.gif"" alt="""" border=""0"" style=""margin-bottom:-2px""/></a> "
			    end if
			   
			    if stat_DelAll or (stat_Del and memName=log_ViewArr(5,0)) then 
	    		   TempStr=TempStr&"<a href=""blogedit.asp?action=del&amp;id="&LogID&""" onclick=""if (!window.confirm('是否要删除该日志')) return false"" title=""删除该日志"" accesskey=""K""><img src=""images/icon_del.gif"" alt="""" border=""0"" style=""margin-bottom:-2px""/></a>"
			    end if
			    
		        TempArticle=Replace(TempArticle,"<"&"%ST(A)%"&">","")
		        TempArticle=Replace(TempArticle,"<$EditAndDel$>",TempStr)
		        TempArticle=Replace(TempArticle,"<$log_ViewNums$>",log_ViewArr(4,0))
		        
			    response.write TempArticle
	            ShowComm LogID,comDesc,log_ViewArr(7,0)
		        call updateViewNums(id,log_ViewArr(4,0))
	       else
			    response.write "读取日志出错.<br/>" & LoadArticle(0) & " : " &  LoadArticle(1)
	       end if
           exit sub
        end If
        
	     '从数据库读取日志
	     'on error resume Next
		set preLog=Conn.Execute("SELECT TOP 1 log_Title,log_ID FROM blog_Content WHERE log_PostTime<#"&DateToStr(log_ViewArr(9,0),"Y-m-d H:I:S")&"# and log_IsShow=true and log_IsDraft=false ORDER BY log_PostTime DESC")
		set nextLog=Conn.Execute("SELECT TOP 1 log_Title,log_ID FROM blog_Content WHERE log_PostTime>#"&DateToStr(log_ViewArr(9,0),"Y-m-d H:I:S")&"# and log_IsShow=true and log_IsDraft=false ORDER BY log_PostTime ASC")
		SQLQueryNums=SQLQueryNums+2

 %>
					   <div id="Content_ContentList" class="content-width"><a name="body" accesskey="B" href="#body"></a>
					   <div class="pageContent">
						   <div style="float:right;width:180px !important;width:auto">
						   <%
								 if not preLog.eof then
							       response.write ("<a href=""?id="&preLog("log_ID")&""" title=""上一篇日志: "&preLog("log_Title")&""" accesskey="",""><img border=""0"" src=""images/Cprevious.gif"" alt=""""/> 上一篇</a>")
							      else
							       response.write ("<img border=""0"" src=""images/Cprevious1.gif"" alt=""这是最新一篇日志""/>上一篇")
							    end if
							    if not nextLog.eof then
							       response.write (" | <a href=""?id="&nextLog("log_ID")&""" title=""下一篇日志: "&nextLog("log_Title")&""" accesskey="".""><img border=""0"" src=""images/Cnext.gif"" alt=""""/> 下一篇</a>")
							      else
							       response.write (" | <img border=""0"" src=""images/Cnext1.gif"" alt=""这是最后一篇日志""/>下一篇")
							    end if
							    preLog.close
							    nextLog.close
							    set preLog=nothing
							    set nextLog=nothing
						   %>
						   </div>
 						   <img src="<%=getCate.cate_icon%>" style="margin:0px 2px -4px 0px" alt=""/> <strong><a href="default.asp?cateID=<%=log_ViewArr(1,0)%>" title="查看所有<%=getCate.cate_Name%>的日志"><%=getCate.cate_Name%></a></strong> <a href="feed.asp?cateID=<%=log_ViewArr(1,0)%>" target="_blank" title="订阅所有<%=getCate.cate_Name%>的日志" accesskey="O"><img border="0" src="images/rss.png" alt="订阅所有<%=getCate.cate_Name%>的日志" style="margin-bottom:-1px"/></a>
					   </div>
					   <div class="Content">
					   <div class="Content-top"><div class="ContentLeft"></div><div class="ContentRight"></div>
					     <h1 class="ContentTitle"><strong><%=HtmlEncode(log_ViewArr(2,0))%></strong></h1>
					     <h2 class="ContentAuthor">作者:<%=log_ViewArr(5,0)%> 日期:<%=DateToStr(log_ViewArr(9,0),"Y-m-d")%></h2>
					   </div>
					    <div class="Content-Info">
						  <div class="InfoOther">字体大小: <a href="javascript:SetFont('12px')" accesskey="1">小</a> <a href="javascript:SetFont('14px')" accesskey="2">中</a> <a href="javascript:SetFont('16px')" accesskey="3">大</a></div>
						  <div class="InfoAuthor"><img src="images/weather/hn2_<%=log_ViewArr(14,0)%>.gif" style="margin:0px 2px -6px 0px" alt=""/><img src="images/weather/hn2_t_<%=log_ViewArr(14,0)%>.gif" alt=""/> <img src="images/<%=log_ViewArr(15,0)%>.gif" style="margin:0px 2px -1px 0px" alt=""/>
						    <%if stat_EditAll or (stat_Edit and log_ViewArr(5,0)=memName) then %>　<a href="blogedit.asp?id=<%=log_ViewArr(0,0)%>" title="编辑该日志" accesskey="E"><img src="images/icon_edit.gif" alt="" border="0" style="margin-bottom:-2px"/></a><%end if%>
					        <%if stat_DelAll or (stat_Del and log_ViewArr(5,0)=memName)  then %>　<a href="blogedit.asp?action=del&amp;id=<%=log_ViewArr(0,0)%>" onclick="if (!window.confirm('是否要删除该日志')) return false" accesskey="K"><img src="images/icon_del.gif" alt="" border="0" style="margin-bottom:-2px"/></a><%end if%>
						  </div>
						</div>
					  <div id="logPanel" class="Content-body">
						<%
					    keyword=CheckStr(Request.QueryString("keyword"))
						if log_ViewArr(10,0)=1 then
						 response.write (highlight(UnCheckStr(UBBCode(HtmlEncode(log_ViewArr(8,0)),mid(log_ViewArr(11,0),1,1),mid(log_ViewArr(11,0),2,1),mid(log_ViewArr(11,0),3,1),mid(log_ViewArr(11,0),4,1),mid(log_ViewArr(11,0),5,1))),keyword))
						else
						 response.write (highlight(UnCheckStr(log_ViewArr(8,0)),keyword))
						end if	%>
					   <br/><br/>

					   </div>
					   <div class="Content-body">
					    <%if len(log_ViewArr(16,0))>0 then response.write (log_ViewArr(16,0)&"<br/>")%>
						<img src="images/From.gif" style="margin:4px 2px -4px 0px" alt=""/><strong>文章来自:</strong> <a href="<%=log_ViewArr(17,0)%>" target="_blank"><%=log_ViewArr(18,0)%></a><br/>
						<img src="images/icon_trackback.gif" style="margin:4px 2px -4px 0px" alt=""/><strong>引用通告:</strong> <a href="<%="trackback.asp?tbID="&id&"&amp;action=view"%>" target="_blank">查看所有引用</a> | <a href="javascript:;" title="获得引用文章的链接" onclick="getTrackbackURL(<%=id%>)">我要引用此文章</a><br/>
					   	<%dim getTag
						  set getTag=new tag
						%>
						 <img src="images/tag.gif" style="margin:4px 2px -4px 0px" alt=""/><strong>Tags:</strong> <%=getTag.filterHTML(log_ViewArr(19,0))%><br/>
					   </div>
					   <div class="Content-bottom"><div class="ContentBLeft"></div><div class="ContentBRight"></div>评论: <%=log_ViewArr(12,0)%> | 引用: <%=log_ViewArr(13,0)%> | 查看次数: <%=log_ViewArr(4,0)%>
					   </div></div>
					   </div>
<%	                   set getTag=nothing
 	  ShowComm LogID,comDesc,log_ViewArr(7,0) '显示评论内容
end sub


'*******************************************
'  显示日志评论内容
'*******************************************
Sub ShowComm(LogID,comDesc,DisComment)
	   response.write ("<a name=""comm_top"" href=""#comm_top"" accesskey=""C""></a>")
	   dim blog_Comment,Pcount,comm_Num,blog_CommID,blog_CommAuthor,blog_CommContent,Url_Add,commArr,commArrLen
	   Set blog_Comment=Server.CreateObject("Adodb.RecordSet")
	   Pcount=0
	   SQL="SELECT comm_ID,comm_Content,comm_Author,comm_PostTime,comm_DisSM,comm_DisUBB,comm_DisIMG,comm_AutoURL,comm_PostIP,comm_AutoKEY FROM blog_Comment WHERE blog_ID="&LogID&" UNION ALL SELECT 0,tb_Intro,tb_Title,tb_PostTime,tb_URL,tb_Site,tb_ID,0,'127.0.0.1',0 FROM blog_Trackback WHERE blog_ID="&LogID&" ORDER BY comm_PostTime "&comDesc
	   blog_Comment.Open SQL,Conn,1,1
	   SQLQueryNums=SQLQueryNums+1
	  IF blog_Comment.EOF AND blog_Comment.BOF Then
      else
	   blog_Comment.PageSize=blogcommpage
	   blog_Comment.AbsolutePage=CurPage
	   comm_Num=blog_Comment.RecordCount
	   
	   commArr=blog_Comment.GetRows(comm_Num)
       blog_Comment.close
       set blog_Comment = nothing
       commArrLen=Ubound(commArr,2)
		   
	   Url_Add="?id="&LogID&"&"%>
       <div class="pageContent"><%=MultiPage(comm_Num,blogcommpage,CurPage,Url_Add,"#comm_top","float:right")%></div>
	   <%
	   Do Until Pcount = commArrLen + 1 OR Pcount=blogcommpage
	   blog_CommID=commArr(0,Pcount)
	   blog_CommAuthor=commArr(2,Pcount)
	   blog_CommContent=commArr(1,Pcount)
	   %>
	  <div class="comment">
	  <%IF blog_CommID=0 Then%>
	    <div class="commenttop"><img src="images/icon_trackback.gif" alt="" style="margin:0px 4px -3px 0px"/><strong><%=("<a href="""&commArr(4,Pcount)&""">"&commArr(5,Pcount)&"</a>")%></strong> <span class="commentinfo">[<%=DateToStr(commArr(3,Pcount),"Y-m-d H:I A")%><%if stat_Admin=true then response.write (" | <a href=""trackback.asp?action=deltb&amp;tbID="&commArr(6,Pcount)&"&amp;logID="&LogID&""" onclick=""if (!window.confirm('是否删除该引用?')) {return false}""><img src=""images/del1.gif"" alt=""删除该引用"" border=""0""/></a>") end if%>]</span></div>
	    <div class="commentcontent">
		<b>标题:</b> <%=blog_CommAuthor%><br/>
		<b>链接:</b> <%=("<a href="""&commArr(4,Pcount)&""" target=""_blank"">"&commArr(4,Pcount)&"</a>")%><br/>
		<b>摘要:</b> <%=checkURL(HTMLDecode(blog_CommContent))%><br/>
<br/>
		</div>
	  <%else%>
	    <div class="commenttop"><a name="comm_<%=blog_CommID%>" href="javascript:addQuote('<%=blog_CommAuthor%>','commcontent_<%=blog_CommID%>')"><img border="0" src="images/<%if memName=blog_CommAuthor then response.write ("icon_quote_author.gif") else response.write ("icon_quote.gif") end if%>" alt="" style="margin:0px 4px -3px 0px"/></a><a href="member.asp?action=view&memName=<%=Server.URLEncode(blog_CommAuthor)%>"><strong><%=blog_CommAuthor%></strong></a> <span class="commentinfo">[<%=DateToStr(commArr(3,Pcount),"Y-m-d H:I A")%> <%if stat_Admin then response.write (" | "&commArr(8,Pcount)) end if%><%if stat_Admin=true or (stat_CommentDel=true and memName=blog_CommAuthor) then response.write (" | <a href=""blogcomm.asp?action=del&amp;commID="&blog_CommID&""" onclick=""if (!window.confirm('是否删除该评论?')) {return false}""><img src=""images/del1.gif"" alt=""删除该评论"" border=""0""/></a>") end if%>]</span></div>
	    <div class="commentcontent" id="commcontent_<%=blog_CommID%>"><%=UBBCode(HtmlEncode(blog_CommContent),commArr(4,Pcount),blog_commUBB,blog_commIMG,commArr(7,Pcount),commArr(9,Pcount))%></div>
	  <%end if%>
	   </div>
	  <%
	   Pcount=Pcount+1
	   loop
       %>
       <div class="pageContent"><%=MultiPage(comm_Num,blogcommpage,CurPage,Url_Add,"#comm_top","float:right")%></div>
       <%
       end if
	   if not DisComment then
	  %>
	  <div id="MsgContent" style="width:94%;">
      <div id="MsgHead">发表评论</div>
      <div id="MsgBody">
      <%
       if not stat_CommentAdd then 
  	    response.write ("你没有权限发表留言！")
  	    response.write ("</div></div>")
  	    exit sub
       end if
      %>
      <script type="text/javascript">
      		function checkCommentPost(){
      			if (!CheckPost) return false
				// 备用方法
      			return true
      		}
      </script>
      <form name="frm" action="blogcomm.asp" method="post" onsubmit="return checkCommentPost()" style="margin:0px;">	  
	  <table width="100%" cellpadding="0" cellspacing="0">	  
	  <tr><td align="right" width="70"><strong>昵　称:</strong></td><td align="left" style="padding:3px;"><input name="username" type="text" size="18" class="userpass" maxlength="24" <%if not memName=empty then response.write ("value="""&memName&""" readonly=""readonly""")%>/></td></tr>
      <%if memName=empty then%><tr><td align="right" width="70"><strong>密　码:</strong></td><td align="left" style="padding:3px;"><input name="password" type="password" size="18" class="userpass" maxlength="24"/> 游客发言不需要密码.</td></tr><%end if%>
	  <%if memName=empty or blog_validate=true then%><tr><td align="right" width="70"><strong>验证码:</strong></td><td align="left" style="padding:3px;"><input name="validate" type="text" size="4" class="userpass" maxlength="4"/> <%=getcode()%></td></tr><%end if%>
	  <tr><td align="right" width="70" valign="top"><strong>内　容:</strong><br/>
	  </td><td style="padding:2px;"><%
	   UBB_TextArea_Height="150px;"
	   UBB_Tools_Items="bold,italic,underline"
	   UBB_Tools_Items=UBB_Tools_Items&"||image,link,mail,quote,smiley"
	   UBBeditor("Message")
	  %></td></tr>
	  <tr><td align="right" width="70" valign="top"><strong>选　项:</strong></td><td align="left" style="padding:3px;">
             <label for="label5"><input name="log_DisSM" type="checkbox" id="label5" value="1" />禁止表情转换</label>
             <label for="label6"><input name="log_DisURL" type="checkbox" id="label6" value="1" />禁止自动转换链接</label>
             <label for="label7"><input name="log_DisKey" type="checkbox" id="label7" value="1" />禁止自动转换关键字</label>
	  </td></tr>
          <tr>
            <td colspan="2" align="center" style="padding:3px;">
			  <input name="logID" type="hidden" value="<%=LogID%>"/>
              <input name="action" type="hidden" value="post"/>
			  <input name="submit2" type="submit" class="userbutton" value="发表评论" accesskey="S"/>
              <input name="button" type="reset" class="userbutton" value="重写"/></td>
          </tr>
          <tr>
            <td colspan="2" align="right" >
			 <%if memName=empty then%>虽然发表评论不用注册，但是为了保护您的发言权，建议您<a href="register.asp">注册帐号</a>. <br/><%end if%> 
	  字数限制 <b><%=blog_commLength%> 字</b> | 
	  UBB代码 <b><%if (blog_commUBB=0) then response.write ("开启") else response.write ("关闭") %></b> | 
	  [img]标签 <b><%if (blog_commIMG=0) then response.write ("开启") else response.write ("关闭") %></b>

			</td>
          </tr>		  
	  </table></form>
	   <%
  	   response.write ("</div></div>")
	  end if
end sub
%>