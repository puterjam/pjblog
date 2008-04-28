<!--#include file="BlogCommon.asp" -->
<!--#include file="header.asp" -->
<!--#include file="common/ModSet.asp" -->
<!--#include file="plugins.asp" -->
<%
'==================================
'  查询页面
'    更新时间: 2005-10-28
'==================================

'处理面板
Dim searchType
IF CheckStr(Request.QueryString("searchType"))<>Empty Then
  searchType=CheckStr(Request.QueryString("searchType"))
Else
  searchType="title"
End if
function getSType(searchType)
  select case searchType
      case "Content"
      getSType="日志内容"
      case "Comments"
      getSType="日志评论"
      case "trackback"
      getSType="引用通告"
      Case Else 
      getSType="日志标题"
     end select
end function
%>
 <!--内容-->
  <div id="Tbody">
   <div id="mainContent">
   <div id="innermainContent">
       <div id="mainContent-topimg"></div>
       <div id="Content_ContentList" class="content-width">
       	<div class="Content">
	     <div class="Content-top"><div class="ContentLeft"></div><div class="ContentRight"></div>
		      <h1 class="ContentTitle"><img src="images/search.gif" alt="" style="margin:0px 4px -3px 0px" class="CateIcon"/><strong>查询结果</strong></h1>
              <h2 class="ContentAuthor">search result</h2>
          </div>
          <%
          dim blog_Search,SearchContent
          SearchContent=CheckStr(Request.QueryString("SearchContent"))
          Url_Add="?SearchContent="&Server.URLEncode(SearchContent)&"&amp;searchType="&searchType&"&amp;"
          	Set blog_Search=Server.CreateObject("ADODB.RecordSet")
          	select case searchType
			  case "Content"
          	    if len(SearchContent)>0 then
			      SQL="SELECT log_ID,log_Title,log_PostTime,log_CommNums FROM blog_Content WHERE log_IsShow=true AND (log_Title LIKE '%"&SearchContent&"%' OR log_Content LIKE '%"&SearchContent&"%')"
          	     else
          	      Response.Redirect("default.asp")
          	    end if
          	 case "Comments"
			   if len(SearchContent)>0 then
			    SQL="SELECT T.log_ID,T.log_Title,C.comm_ID,C.comm_Author,C.comm_Content,C.comm_DisSM,C.comm_AutoURL,C.comm_AutoKEY,C.comm_PostTime FROM blog_Comment C,blog_Content T WHERE C.blog_ID=T.log_ID AND T.log_IsShow=true AND C.comm_Content LIKE '%"&SearchContent&"%'"
          	   else
			    SQL="SELECT T.log_ID,T.log_Title,C.comm_ID,C.comm_Author,C.comm_Content,C.comm_DisSM,C.comm_AutoURL,C.comm_AutoKEY,C.comm_PostTime FROM blog_Comment C,blog_Content T WHERE C.blog_ID=T.log_ID AND T.log_IsShow=true order by comm_ID desc"
          	   end if
          	  case "trackback"
			  if len(SearchContent)>0 then
          	    SQL="SELECT T.blog_ID,T.tb_Title,T.tb_URL,T.tb_Intro,T.tb_Site,T.tb_PostTime,C.log_Title FROM blog_Content C,blog_Trackback T WHERE T.blog_ID=C.log_ID AND C.log_IsShow=true AND (tb_Title LIKE '%"&SearchContent&"%' or tb_Intro LIKE '%"&SearchContent&"%')"
          	   else
          	    SQL="SELECT T.blog_ID,T.tb_Title,T.tb_URL,T.tb_Intro,T.tb_Site,T.tb_PostTime,C.log_Title FROM blog_Content C,blog_Trackback T WHERE T.blog_ID=C.log_ID AND C.log_IsShow=true order by tb_ID desc"
          	   end if
          	  case else
          	    if len(SearchContent)>0 then
          	      SQL="SELECT log_ID,log_Title,log_PostTime,log_CommNums FROM blog_Content WHERE log_IsShow=true AND log_Title LIKE '%"&SearchContent&"%'"
          	     else
          	      Response.Redirect("default.asp")
          	    end if
          	 end select
          	blog_Search.Open SQL,Conn,1,1
          	%>
      <div class="Content-Info">
		 <div class="InfoAuthor"> 查询类型 <b><%=getSType(searchType)%></b> | 关键词 <b><%=HTMLEncode(SearchContent)%></b> | 共找到 <b><%=blog_Search.RecordCount%></b> 条记录</div>
     </div>
           
<%		
  IF blog_Search.EOF AND blog_Search.BOF Then
   response.write ("<b>没有找到符合条件的数据</b><br/><br/><a href=""default.asp"">返回</a><br/><br/>")
  else
	blog_Search.PageSize=10
	blog_Search.AbsolutePage=CurPage
	Dim Search_Nums,SearchArr,SearchArrLen
	Search_Nums=blog_Search.RecordCount
	SearchArr=blog_Search.GetRows(Search_Nums)
    blog_Search.close
    Set blog_Search=Nothing

	Dim MultiPages,PageCount
	SQLQueryNums=SQLQueryNums+1
    SearchArrLen=Ubound(SearchArr,2)
    
    response.write "<div class=""pageContent"">"&MultiPage(Search_Nums,10,CurPage,Url_Add,"","float:left")&"</div><div class=""Content-body"" style=""line-height:200%"">"
    
    Do Until PageCount = SearchArrLen + 1 OR PageCount=10
          	select case searchType
			  case "Content"
	          	Response.Write("<a href=""article.asp?id="&SearchArr(0,PageCount)&"&keyword="&Server.URLEncode(SearchContent)&""" target=""_blank""> "&highlight(SearchArr(1,PageCount),SearchContent)&" [ 日期: "&SearchArr(2,PageCount)&" | 评论数:"&SearchArr(3,PageCount)&"]</a><br/>")
			  case "Comments"
			  %>
	                <div class="commenttop"><img border="0" src="images/icon_quote.gif" alt="" style="margin:0px 4px -3px 0px"/></a><strong><%=SearchArr(3,PageCount)%></strong> <span class="commentinfo">[<%=DateToStr(SearchArr(8,PageCount),"Y-m-d H:I A")%>]</span></div>
	                <div class="commentcontent"><%=highlight(UBBCode(HtmlEncode(SearchArr(4,PageCount)),SearchArr(5,PageCount),blog_commUBB,blog_commIMG,SearchArr(6,PageCount),SearchArr(7,PageCount)),SearchContent)%></div>
	         <%
	         	Response.Write("<div class=""Content-bottom"">相关日志: <a href=""article.asp?id="&SearchArr(0,PageCount)&"#comm_"&SearchArr(2,PageCount)&"""><b>"&SearchArr(1,PageCount)&"</b></a></div>")
          	  case "trackback"
          	  %>
	             <div class="commenttop"><img src="images/icon_trackback.gif" alt="" style="margin:0px 4px -3px 0px"/><strong><%=("<a href="""&SearchArr(2,PageCount)&""">"&SearchArr(4,PageCount)&"</a>")%></strong> <span class="commentinfo">[<%=DateToStr(SearchArr(5,PageCount),"Y-m-d H:I A")%>]</span></div>
	             <div class="commentcontent">
	            	<b>标题:</b> <%=highlight(SearchArr(1,PageCount),SearchContent)%><br/>
		            <b>链接:</b> <%=("<a href="""&SearchArr(2,PageCount)&""" target=""_blank"">"&SearchArr(2,PageCount)&"</a>")%><br/>
	             	<b>摘要:</b> <%=highlight(checkURL(HTMLDecode(SearchArr(3,PageCount))),SearchContent)%><br/>
		         </div>
		   <%
	         	Response.Write("<div class=""Content-bottom"">相关日志: <a href=""article.asp?id="&SearchArr(0,PageCount)&""" target=""_blank""><b>"&SearchArr(6,PageCount)&"</b></a></div>")
          	  case else
	        	Response.Write("<a href=""article.asp?id="&SearchArr(0,PageCount)&"""> "&highlight(SearchArr(1,PageCount),SearchContent)&" [ 日期: "&SearchArr(02,PageCount)&" | 评论数:"&SearchArr(3,PageCount)&"]</a><br/>")
          	end select
	    PageCount=PageCount+1
	loop
    response.write "</div><div class=""pageContent"">"&MultiPage(Search_Nums,10,CurPage,Url_Add,"","float:left")&"</div>"
  end if
%>         
         </div>
        </div>
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