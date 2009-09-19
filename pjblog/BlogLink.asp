<!--#include file="BlogCommon.asp" -->
<!--#include file="header.asp" -->
<!--#include file="common/ModSet.asp" -->
<!--#include file="plugins.asp" -->
<!--内容-->
<%
'==================================
'  友情连接
'    更新时间: 2005-12-11
'==================================
%>
 <!--内容-->
  <div id="Tbody">
   <div id="mainContent">
     <div id="innermainContent">
       <div id="mainContent-topimg"></div>
       	 <div id="Content_ContentList" class="content-width">
       	 
         <%
If request.Form("action") = "postLink" Then
    Dim link_Name, link_URL, link_Image, linkCount, linkDB, linkvalidate
    link_Name = checkURL(checkstr(request.Form("link_Name")))
    link_URL = checkURL(checkstr(request.Form("link_URL")))
    link_Image = checkURL(checkstr(request.Form("link_Image")))
    linkvalidate = checkURL(checkstr(request.Form("link_validate")))
    If CStr(LCase(Session("GetCode")))<>CStr(LCase(linkvalidate)) And Not stat_Admin Then
        showmsg lang.Tip.Link.Err, "<b>" & lang.Err.info(5) & "</b><br/><a href=""javascript:history.go(-1);"">" & lang.Tip.SysTem(6) & "</a>", "ErrorIcon", ""
    End If
    If Len(link_Name)<1 Then
        showmsg lang.Tip.Link.Err, "<b>" & lang.Err.info(6) & "</b><br/><a href=""javascript:history.go(-1);"">" & lang.Tip.SysTem(6) & "</a>", "ErrorIcon", ""
    End If
    If Len(link_URL)<1 Then
        showmsg lang.Tip.Link.Err, "<b>" & lang.Err.info(7) & "</b><br/><a href=""javascript:history.go(-1);"">" & lang.Tip.SysTem(6) & "</a>", "ErrorIcon", ""
    End If

    linkCount = Int(conn.Execute("select count(*) from blog_links")(0))
    Set linkDB = Server.CreateObject("ADODB.RecordSet")
    linkDB.Open "blog_links", Conn, 1, 2
    linkDB.addNew
    linkDB("link_Name") = link_Name
    linkDB("link_URL") = link_URL
    linkDB("link_Image") = link_Image
    linkDB("link_Order") = linkCount
    linkDB("link_IsShow") = False
    linkDB.update
    linkDB.Close
    Set linkDB = Nothing
    showmsg lang.Tip.Link.Suc, "<b>" & lang.Tip.Link.Aduit & "</b><br/><a href=""BlogLink.asp"">" & lang.Tip.SysTem(6) & "</a>", "MessageIcon", ""
End If
On Error Resume Next
server.Execute("post/link.html")
If Err Then
    Err.Clear
    Dim blog_Links, ImgLink, TextLink
    Set blog_Links = conn.Execute("select * from blog_Links where link_IsShow=true order by link_Order asc")
    SQLQueryNums = SQLQueryNums + 1
    Do Until blog_Links.EOF
        If Len(blog_Links("link_Image"))>0 Then
            ImgLink = ImgLink&"<a href="""&blog_Links("link_URL")&""" target=""_blank""><img src="""&blog_Links("link_Image")&""" alt="""&blog_Links("link_Name")&""" border=""0"" style=""margin:3px;width:88px;height:31px""/></a>"
        Else
            TextLink = TextLink&"<div class=""link"" style=""width:108px;float:left;overflow:hidden;margin-right:8px;height:24px;line-height:180%""><a href="""&blog_Links("link_URL")&""" target=""_blank"" title="""&blog_Links("link_Name")&""">"&blog_Links("link_Name")&"</a></div>"
        End If
        blog_Links.movenext
    Loop

%>
               <div class="Content">
                 <div class="Content-top"><div class="ContentLeft"></div><div class="ContentRight"></div>
                   <h1 class="ContentTitle"><img src="images/image.gif" alt="" style="margin:0px 2px -3px 0px" class="CateIcon"/><b><%=lang.Tip.Link.ImgLink(1)%></b></h1>
                   <h2 class="ContentAuthor"><%=lang.Tip.Link.ImgLink(2)%></h2>
                 </div>
                 <div class="Content-body"><%=ImgLink%></div>
               </div>
               <div class="Content">
                 <div class="Content-top"><div class="ContentLeft"></div><div class="ContentRight"></div>
                   <h1 class="ContentTitle"><img src="images/html.gif" alt="" style="margin:0px 2px -3px 0px" class="CateIcon"/><b><%=lang.Tip.Link.TextLink(1)%></b></h1>
                   <h2 class="ContentAuthor"><%=lang.Tip.Link.TextLink(2)%></h2>
                 </div>
                 <div class="Content-body"><%=TextLink%></div>
               </div>
         <%end If%>
         
           <div id="MsgContent" style="width:94%;">
                <div id="MsgHead"><%=lang.Tip.Link.Form(1)%></div>
                <div id="MsgBody">
                <form name="frm" action="BlogLink.asp" method="post" onsubmit="return CheckPost()" style="margin:0px;">	  
          	  <table width="100%" cellpadding="0" cellspacing="0">
          	  <tr><td align="right" width="70"><strong><%=lang.Tip.Link.Form(2)%>:</strong></td><td align="left" style="padding:3px;"><input name="link_Name" type="text" size="35" class="userpass" maxlength="40"/> <span style="color:#f00">*</span></td></tr>
          	  <tr><td align="right" width="70"><strong><%=lang.Tip.Link.Form(3)%>:</strong></td><td align="left" style="padding:3px;"><input name="link_URL" type="text" size="50" class="userpass"/> <span style="color:#f00">*</span></td></tr>
          	  <tr><td align="right" width="70"><strong><%=lang.Tip.Link.Form(4)%>:</strong></td><td align="left" style="padding:3px;"><input name="link_Image" type="text" size="50" class="userpass"/></td></tr>
          	  <tr><td align="right" width="70"><strong><%=lang.Tip.Link.Form(5)%>:</strong></td><td align="left" style="padding:3px;"><input name="link_validate" type="text" size="4" class="userpass" maxlength="4" onfocus="this.select()"/> <%=getcode()%></td></tr>
          	  <tr><td align="right" width="70"></td><td align="left"><%=lang.Tip.Link.Form(6)%></td></tr>
                    <tr>
                      <td colspan="2" align="center" style="padding:3px;">
                        <input name="action" type="hidden" value="postLink"/>
          			    <input type="submit" class="userbutton" value="<%=lang.Action.Submit%>"/>
                        <input name="button" type="reset" class="userbutton" value="<%=lang.Action.ReSet%>"/>
                        </td>
                    </tr>
          	  </table></form>
          </div></div>

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
