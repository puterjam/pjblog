<!--#include file="BlogCommon.asp" -->
<!--#include file="header.asp" -->
<!--#include file="common/ModSet.asp" -->
<!--#include file="common/UBBconfig.asp" -->
<!--#include file="class/cls_article.asp" -->
<!--#include file="common/sha1.asp" -->
<!--#include file="plugins.asp" -->

<%
'==================================
'  显示日志
'    更新时间: 2007-5-22
'==================================
'处理日志信息

      dim id,tKey
	  If CheckStr(Request.QueryString("id"))<>Empty Then
			id=CheckStr(Request.QueryString("id"))
	  End If
      Dim log_View,log_ViewArr,keyword,preLog,nextLog,blog_Cate,blog_CateArray,comDesc
      Dim getCate
      set getCate=new Category
      if IsInteger(id) then
		  Set log_View=Server.CreateObject("ADODB.RecordSet")
		  if blog_postFile then
		    SQL="SELECT top 1 log_ID,log_CateID,log_title,Log_IsShow,log_ViewNums,log_Author,log_comorder,log_DisComment FROM blog_Content WHERE log_ID="&id&" and log_IsDraft=false"
		   else
		    SQL="SELECT top 1 log_ID,log_CateID,log_title,Log_IsShow,log_ViewNums,log_Author,log_comorder,log_DisComment,log_Content,log_PostTime,log_edittype,log_ubbFlags,log_CommNums,log_QuoteNums,log_weather,log_level,log_Modify,log_FromUrl,log_From,log_tag FROM blog_Content WHERE log_ID="&id&" and log_IsDraft=false"
		  end if
		  
		  log_View.Open SQL,Conn,1,3
		  SQLQueryNums=SQLQueryNums+1
		  if log_View.eof or log_View.bof then log_View.close:showmsg "错误信息","不存在当前日志！<br/><a href=""default.asp"">单击返回</a>","ErrorIcon",""
	      log_View("log_ViewNums")=log_View("log_ViewNums")+1
		  log_View.UPDATE
		  log_ViewArr=log_View.GetRows
		  log_View.Close
		  set log_View = nothing
		  getCate.load(int(log_ViewArr(1,0))) '获取分类信息
		  
		  if log_ViewArr(3,0) and not getCate.cate_Secret then 
		   		BlogTitle=log_ViewArr(2,0) & " - " & siteName
		  end if
      else
		showmsg "错误信息","非法操作","ErrorIcon",""
	  end if
	  getBlogHead BlogTitle,getCate.cate_Name,getCate.cate_ID
	  tKey = getTempKey
%>
 <!--内容-->
  <div id="Tbody">
  <div id="mainContent">
   <div id="innermainContent">
       <div id="mainContent-topimg"></div>
	   <%=content_html_Top%>
	   <%
	   if id<>"" and IsInteger(id)=False then 
	      response.write ("非法操作！！")
	    else
	      ShowArticle id '显示日志
	    end if  
	    set getCate=nothing
	  %>
	   <%=content_html_Bottom%>
       <div id="mainContent-bottomimg"></div>
   </div>
   </div>
   <%Side_Module_Replace '处理系统侧栏模块信息%>
   <div id="sidebar">
	   <div id="innersidebar">
		     <div id="sidebar-topimg"><!--工具条顶部图象--></div>
			  <%=side_html%>
		     <div id="sidebar-bottomimg"></div>
	   </div>
   </div>
   <div style="clear: both;height:1px;overflow:hidden;margin-top:-1px;"></div>
  </div>
  <!--#include file="footer.asp" -->
