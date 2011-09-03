<!--#include file="../../commond.asp" -->
<!--#include file="../../header.asp" -->
<!--#include file="../../plugins.asp" -->
<!--#include file="../../common/ModSet.asp" -->
<!--#include file="../../common/UBBconfig.asp" -->

 <!--内容-->
 <style>
  .CFace:link,.CFace:visited{padding:1px;padding-left:0px;padding-right:0px;border:1px solid #164283;background:#c9dbf5;border-bottom-width:3px;}
  .CFace:hover{padding:1px;padding-left:0px;padding-right:0px;border:1px solid #164283;background:#9bbbec;border-bottom-width:3px}
  .LFace:link,.LFace:visited{padding:1px;}
  .LFace:hover{padding-left:0px;padding-right:0px;border:1px solid #164283;background:#c9dbf5}
  </style>
  <div id="Tbody">
   <div id="mainContent">
     <div id="innermainContent">
       <div id="mainContent-topimg"></div>
	   <%=content_html_Top%>
	<%'在这里写代码，为了保证页面完整性，其他地方尽量不要修改 %>
	    <div id="Content_Contentlist" class="content-width">
	    
	    <%	     
	    	dim GBSet,Getplugins,action,replyMsg,MsgID,MsgReplyContent
	        Set GBSet=New ModSet
	        replyMsg=false
	        action = Request.QueryString("action")
	        MsgID = CheckStr(Request.QueryString("id"))
	        if action="Reply" then
	         replyMsg=true
	        end if
	         if replyMsg then
	          if not (memName<>empty and stat_Admin) then
                showmsg "错误信息","你没有权限回复留言<br/><a href=""javascript:history.go(-1)"">单击返回</a>","WarningIcon",""
	          end if
	          If MsgID=Empty then 
	            showmsg "错误信息","非法操作！<br/><a href=""javascript:history.go(-1)"">单击返回</a>","WarningIcon","" 
	          end if
	          If IsInteger(MsgID)=False then 
	            showmsg "错误信息","非法操作！<br/><a href=""javascript:history.go(-1)"">单击返回</a>","WarningIcon","" 
	          end if
	         end if
	         
	        Getplugins=CheckStr(Request.QueryString("plugins"))
	        GBSet.open(Getplugins)
            if not cBool(GBSet.getKeyValue("OpenState")) then
              showmsg "错误信息","留言本暂时关闭！<br/><a href=""default.asp"">单击返回首页</a>","WarningIcon",""
            end if
            
	        Dim GuestDB,GuestNum,PageCount
	        set GuestDB=Server.CreateObject("Adodb.Recordset")
	        if replyMsg then
	          SQL="select * from blog_book where book_id=" & MsgID
	         else
	          SQL="select * from blog_book order by book_ID desc"
	         end if
	        SQLQueryNums=SQLQueryNums+1
	        GuestDB.Open SQL,CONN,1,1
	        Url_Add=Url_Add&"plugins=GuestBookForPJBlog&"
	        if GuestDB.eof and GuestDB.bof then
               response.write "<div style=""margin:10px 0px 10px 0px""><strong>抱歉，没有找到任何留言！</strong></div>"
               if replyMsg then 
                showmsg "错误信息","没有找到可以回复的留言！<br/><a href=""javascript:history.go(-1)"">单击返回</a>","ErrorIcon","" 
               end if
	        else
	            dim bookPage
	            bookPage=int(GBSet.getKeyValue("bookPage"))
                GuestDB.PageSize=bookPage
                GuestDB.AbsolutePage=CurPage
                GuestNum=GuestDB.RecordCount
                  %>
       	     <%if not replyMsg then%><div class="pageContent"><%=MultiPage(GuestNum,bookPage,CurPage,Url_Add,"","float:right","")%></div><%end if%>
       	     <div class="Content">
       	       <%
                 Do Until GuestDB.eof or PageCount=bookPage
                 if replyMsg then MsgReplyContent=GuestDB("book_reply") else MsgReplyContent=""
       	       %>
       	       <div class="comment" style="margin-bottom:20px;
<%
			If blog_GravatarOpen Then
				' Gravatar 头像基本设置
					Dim NewGravatar, GravatarImg
					Set NewGravatar = new Gravatar
					If not isblank(GuestDB("email")) Then
						NewGravatar.Gravatar_EmailMd5 = Trim(MD5(Trim(GuestDB("email"))))
						GravatarImg = lcase(NewGravatar.outPut())
					Else
						NewGravatar.Gravatar_EmailMd5 = ""
						GravatarImg = "images/gravatar.gif"
					End If
					Set NewGravatar = nothing
%>
				text-align:right
<%
			End If
%>
               ">
<%
			If blog_GravatarOpen Then
%>
                <div class="commentleft Gravatar" style="float:left" id="Gravatar_<%=MsgID%>">
					<img src="<%=GravatarImg%>" alt="<%=GuestDB("book_Messager")%>" border="0" />
				    </div><div class="commentright" style="text-align:left">
<%                    
			End If
%>

       	         <div class="commenttop"><a name="book_<%=GuestDB("book_ID")%>"></a>
       	         <%if not GuestDB("book_HiddenReply") then%>
       	             <img src="images/icon_quote.gif" alt="公开" border="0" style="margin:0px 1px -3px 0px"/>
       	           <%else%>
       	             <img src="images/icon_lock.gif" alt="隐藏" border="0" style="margin:0px 1px -3px 0px"/>
       	         <%end if%>
       	         <a href="member.asp?action=view&memName=<%=Server.URLEncode(GuestDB("book_Messager"))%>"><b><%=GuestDB("book_Messager")%></b></a> 
<%
if trim(GuestDB("email"))<>"" then
    response.write " <a href=mailto:"&trim(GuestDB("email"))&" target=_blank><img src=images/email1.gif border=0></a>"
else
    response.write " <img src=images/noemail1.gif>"
end if
if trim(GuestDB("siteurl"))<>"" and trim(GuestDB("siteurl"))<>"http://" then
    response.write " <a href="&trim(GuestDB("siteurl"))&" target=_blank><img src=images/url1.gif border=0></a>"
else
    response.write " <img src=images/nourl1.gif>"
end if
%>
                 
                 <span class="commentinfo">[<%=DateToStr(GuestDB("book_PostTime"),"Y-m-d H:I:S")%>
       	         <%if stat_Admin then response.write " | " & GuestDB("book_IP")%>
       	         <%if stat_Admin then%> | <a href="LoadMod.asp?plugins=GuestBookForPJBlog&amp;action=Reply&amp;id=<%=GuestDB("book_ID")%>"><img src="Plugins/<%=GBSet.GetPath%>/reply.gif" alt="回复" border="0" style="margin-bottom:-3px"/></a><%end if%>
       	         <%if (memName<>empty and stat_Admin) or (cbool(GBSet.getKeyValue("canDel")) and GuestDB("book_Messager")=memName) then%> | <a href="Plugins/<%=GBSet.GetPath%>/bookaction.asp?action=del&amp;id=<%=GuestDB("book_ID")%>" onclick="if (!confirm('确定删除该留言信息吗？')) return false "><img src="Plugins/<%=GBSet.GetPath%>/del.gif" alt="删除" border="0" style="margin-bottom:-3px"/></a><%end if%>
       	         
       	         ]</span></div>
                 
        	      <%if (GuestDB("book_HiddenReply") and Lcase(GuestDB("book_Messager"))=Lcase(memName)) Or stat_Admin or (not GuestDB("book_HiddenReply")) then %>
            	      <div class="commentcontent"> 
                	       <table border="0" cellspacing="0" cellpadding="0">
                    	       <tr>
                    	         <td width="60" valign="top" align="center"><img src="<%response.write GBSet.getKeyValue(GuestDB("book_face"))%>" alt="" border="0" style="margin-left:-4px"/><div style="height:6px"></div></td>
                    	         <td valign="top"><%=UBBCode(HtmlEncode(GuestDB("book_Content")),0,GBSet.getKeyValue("ubbCode"),GBSet.getKeyValue("ImgCode"),GBSet.getKeyValue("autoURL"),1)%></td>
                    	       </tr>
                    	   </table>
                	   </div>
            	       <%if len(GuestDB("book_reply"))>0 and not replyMsg then %>
               	         <div class="commenttop"><img src="images/reply.gif" alt="" border="0" style="margin:0px 3px -3px 0px"/><a href="member.asp?action=view&memName=<%=Server.URLEncode(GuestDB("book_replyAuthor"))%>"><b><%=GuestDB("book_replyAuthor")%></b></a> <span class="commentinfo">[<%=DateToStr(GuestDB("book_replyTime"),"Y-m-d H:I:S")%>]</span></div>
                	     <div class="commentcontent"><%=UBBCode(HtmlEncode(GuestDB("book_reply")),0,0,0,1,1)%></div>
            	       <%end if%>
        	      <%else%>   
            	       <div class="commentcontent"><b>该留言为隐藏留言! 只有管理员和留言者可以查看.</b></div>
        	      <%end if%>
<%
			If blog_GravatarOpen Then
%>
				</div>
<%
			End If
%>
        	     </div>
       	      <%
       	        GuestDB.movenext
       	        PageCount=PageCount+1
       	        loop
	            response.write "</div>"
	            if not replyMsg then response.write "<div class=""pageContent"">"&MultiPage(GuestNum,bookPage,CurPage,Url_Add,"","float:right","")&"</div>" end if

	        end if 
	       GuestDB.close
	       set GuestDB=nothing
	      %>
	    </div>	  
          <div id="MsgContent" style="width:94%;">
                <div id="MsgHead"><%if not replyMsg then%>发表留言<%else%>回复留言<%end if%></div>
                <div id="MsgBody">
                <form name="frm" action="Plugins/<%=GBSet.GetPath%>/bookaction.asp" method="post" onsubmit="return CheckPost()" style="margin:0px;">	  
          	  <table width="100%" cellpadding="0" cellspacing="0">
          <%if not replyMsg then%>	  
          	  <tr><td align="right" width="70"><strong>昵　称:</strong></td><td align="left" style="padding:3px;"><input name="username" type="text" size="18" class="userpass" maxlength="24" <%if not memName=empty then response.write ("value="""&memName&""" readonly=""readonly""")%>/></td></tr>
          	  <%if memName=empty then%><tr><td align="right" width="70"><strong>密　码:</strong></td><td align="left" style="padding:3px;"><input name="password" type="password" size="18" class="userpass" maxlength="24"/></td></tr><%end if%>
              <%if memName=empty then%><tr><td align="right" width="70"><strong>邮　箱:</strong></td><td align="left" style="padding:3px;"><input name="myblogemail" type="text" size="18" class="userpass" maxlength="24"/> 请填写您的邮箱.</td></tr><%end if%>
              <%if memName=empty then%><tr><td align="right" width="70"><strong>网　址:</strong></td><td align="left" style="padding:3px;"><input name="myblogsiteurl" type="text" class="userpass" value="http://" size="18" maxlength="24"/> 请填写您的网址.</td></tr><%end if%>

          	  <tr><td align="right" width="70"><strong>验证码:</strong></td><td align="left" style="padding:3px;"><input name="validate" type="text" size="4" class="userpass" maxlength="4" onFocus="get_checkcode();this.onfocus=null;" onKeyUp="ajaxcheckcode('isok_checkcode',this);"/> <span id="checkcode"><label style="cursor:pointer;" onClick="get_checkcode();">点击获取验证码</label></span> <span id="isok_checkcode"></span></td></tr>
          	  <tr><td align="right" width="70" valign="top" height="56"><strong>头　像:</strong></td>
          	  <td style="padding:2px;" align="left">
          	   <input type="hidden" name="book_face" value="face1"/>
          	    <a href="javascript:void(0)" class="CFace" id="face1" onclick="S('face1')"><img src="<%response.write GBSet.getKeyValue("face1")%>" alt="" border="0"/></a>
          	    <a href="javascript:void(0)" class="LFace" id="face2" onclick="S('face2')"><img src="<%response.write GBSet.getKeyValue("face2")%>" alt="" border="0"/></a>
          	    <a href="javascript:void(0)" class="LFace" id="face3" onclick="S('face3')"><img src="<%response.write GBSet.getKeyValue("face3")%>" alt="" border="0"/></a>
          	    <a href="javascript:void(0)" class="LFace" id="face4" onclick="S('face4')"><img src="<%response.write GBSet.getKeyValue("face4")%>" alt="" border="0"/></a>
          	    <a href="javascript:void(0)" class="LFace" id="face5" onclick="S('face5')"><img src="<%response.write GBSet.getKeyValue("face5")%>" alt="" border="0"/></a>
          	    <a href="javascript:void(0)" class="LFace" id="face6" onclick="S('face6')"><img src="<%response.write GBSet.getKeyValue("face6")%>" alt="" border="0"/></a>
          	    <a href="javascript:void(0)" class="LFace" id="face7" onclick="S('face7')"><img src="<%response.write GBSet.getKeyValue("face7")%>" alt="" border="0"/></a>
          	    <script type="text/javascript">
          	     var LastA=document.getElementById("face1")
          	     function S(Face){
          	      LastA.className="LFace"
          	      document.getElementById(Face).className="CFace"
          	      LastA=document.getElementById(Face)
          	      document.getElementById(Face).blur()
          	      document.forms["frm"].book_face.value=Face
          	     }
          	    </script>
          	  </td></tr>
         <%end if%>
          	  <tr><%if not replyMsg then%><td align="right" width="70" valign="top"><strong>内　容:</strong><br/><%end if%>
          	  </td><td style="padding:2px;">
             	  <%UBB_TextArea_Height="150px;"
             	   UBB_Tools_Items="bold,italic,underline"
             	   UBB_Tools_Items=UBB_Tools_Items&"||image,link,mail,quote,smiley"
   	               UBB_Msg_Value=UBBFilter(UnCheckStr(MsgReplyContent))
             	   UBBeditor("Message")
             	   %>
          	  </td></tr>
         <%if not replyMsg then%>
          	  <tr><td align="right" width="70"><strong>选　项:</strong></td><td align="left" style="padding:3px;">
                       <label for="label5"><input name="hiddenMsg" type="checkbox" id="label5" value="1" />隐藏留言</label>
         <%end if%>
          	  </td></tr>
                    <tr>
                      <td colspan="2" align="center" style="padding:3px;">
                        <%if replyMsg then%><input name="MsgID" type="hidden" value="<%=MsgID%>"/><%end if%>
                        <input name="action" type="hidden" value="<%if not replyMsg then%>post<%else%>reply<%end if%>"/>
          			    <input name="submit" type="submit" class="userbutton" value="<%if not replyMsg then%>发表留言<%else%>回复留言<%end if%>"/>
                        <input name="button" type="reset" class="userbutton" value="重写"/>
                        <%if replyMsg then%><input name="button" type="button" class="userbutton" value="返回" onclick="history.go(-1)"/><%end if%>
                        </td>
                    </tr>
         <%if not replyMsg then%>
                    <tr>
                      <td colspan="2" align="right" >
                       	 <%if memName=empty then%>发表留言不用注册，但是为了保护您的发言权，建议您<a href="register.asp">注册帐号</a>. <br/><%end if%> 
                       	  字数限制 <b><%=GBSet.getKeyValue("charCount")%> 字</b> | 
                       	  UBB代码 <b><%if not cBool(GBSet.getKeyValue("ubbCode")) then response.write ("开启") else response.write ("关闭") %></b> | 
                       	  [img]标签 <b><%if not cBool(GBSet.getKeyValue("ImgCode")) then response.write ("开启") else response.write ("关闭") %></b>
          
          			</td>
                    </tr>	
         <%end if%>
          	  </table></form>
          </div></div>
	   <%=content_html_Bottom%>
       <div id="mainContent-bottomimg"></div>
   </div>
   </div>
   
   <div id="sidebar">
    <div id="innersidebar">
     <div id="sidebar-topimg"><!--工具条顶部图象--></div>
	  <%=side_html%>
     <div id="sidebar-bottomimg"></div>
   </div>
  </div>
 </div>
 <div style="font: 0px/0px sans-serif;clear: both;display: block"></div>
 <!--#include file="../../footer.asp" -->