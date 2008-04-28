<!--#include file="commond.asp" -->
<!--#include file="header.asp" -->
<!--#include file="common/UBBconfig.asp" -->
<!--#include file="FCKeditor/fckeditor.asp" -->
<!--#include file="common/ModSet.asp" -->
<!--#include file="class/cls_logAction.asp" -->
<div id="Tbody">
    <div style="text-align:center;"><br/>
<%
'==================================
'  日志编辑或删除
'    更新时间: 2006-5-29
'==================================
dim logid
dim preLog,nextLog,getTag
dim loadTag,loadTags

logid=Trim(CheckStr(Request("id")))
If ChkPost() then
  if logid=empty or IsInteger(logid)=False then 
   %>
        <div id="MsgContent" style="width:350px">
         <div id="MsgHead">出错信息</div>
         <div id="MsgBody">
   		 <div class="ErrorIcon"></div>
           <div class="MessageText"><b>非法操作</b><br/>
           <a href="javascript:history.go(-1)">单击返回上一页</a> 
   		 </div>
   	 </div>
   	</div>
   <%else
   dim lArticle,EditLog,DeleteLog
   set lArticle=new logArticle
   editLog = lArticle.getLog(logid)

   if stat_EditAll or (stat_Edit and lArticle.logAuthor = memName) Then%>
    <!--内容-->
   <%IF Request.Form("action")="post" Then
         lArticle.categoryID = request.form("log_CateID")
         lArticle.logTitle = request.form("title")
         lArticle.logAuthor = memName
         lArticle.logEditType = request.form("log_editType")
         lArticle.logIntroCustom = request.form("log_IntroC")
         lArticle.logIntro = request.form("log_Intro")
         lArticle.logWeather = request.form("log_weather")
         lArticle.logLevel = request.form("log_Level")
         lArticle.logCommentOrder = request.form("log_comorder")
         lArticle.logDisableComment = request.form("log_DisComment")
         lArticle.logIsShow = request.form("log_IsShow")
         lArticle.logIsTop = request.form("log_IsTop")
         lArticle.logIsDraft = request.form("log_IsDraft")
         lArticle.logFrom = request.form("log_From")
         lArticle.logFromURL = request.form("log_FromURL")
         lArticle.logDisableImage = request.form("log_disImg")
         lArticle.logDisableSmile = request.form("log_DisSM")
         lArticle.logDisableURL = request.form("log_DisURL")
         lArticle.logDisableKeyWord = request.form("log_DisKey")
         lArticle.logMessage = request.form("Message")
         lArticle.logTrackback = request.form("log_Quote")
         lArticle.logTags = request.form("tags")
         lArticle.logPubTime = request.form("PubTime")
	 	 lArticle.logPublishTimeType = request.form("PubTimeType")
	 	 EditLog = lArticle.editLog(request.form("id"))
    set lArticle=nothing
		  %>
		      <div id="MsgContent" style="width:300px">
		        <div id="MsgHead">反馈信息</div>
		        <div id="MsgBody">
		  		 <div class="<%if EditLog(0)<0 then response.write "ErrorIcon" else response.write "MessageIcon"%>"></div>
		          <div class="MessageText"><%=EditLog(1)%><br/><a href="default.asp">点击返回首页</a><br/>
		  		 <%if EditLog(0)>=0 then %>
			      	 <a href="default.asp?id=<%=EditLog(2)%>">返回你所编辑的日志</a><br/>
		  		     <meta http-equiv="refresh" content="3;url=default.asp?logID=<%=EditLog(2)%>"/>
			     <%end if%>
		  	  </div>
		  	</div>
		    </div>
		    <%
	  Elseif Request.QueryString("action")="del" Or Request.form("action")="del" then
		DeleteLog=lArticle.deleteLog(request("id"))
        set lArticle=nothing
		  %>
		      <div id="MsgContent" style="width:300px">
		        <div id="MsgHead">反馈信息</div>
		        <div id="MsgBody">
		  		 <div class="<%if DeleteLog(0)<0 then response.write "ErrorIcon" else response.write "MessageIcon"%>"></div>
		          <div class="MessageText"><%=DeleteLog(1)%><br/><a href="default.asp">点击返回首页</a><br/>
		  	  </div>
		  	</div>
		    </div>
		    <%
	else

   if editLog(0)<0 then
   %>
        <div id="MsgContent" style="width:350px">
         <div id="MsgHead">出错信息</div>
         <div id="MsgBody">
   		 <div class="ErrorIcon"></div>
           <div class="MessageText"><b><%=editLog(1)%></b><br/>
           <a href="default.asp">返回首页</a> 
   		 </div>
   	 </div>
   	</div>
   <%
   else
   	dim log_editType,editTs
   	log_editType=lArticle.logEditType
   %>
   
   <!--第二步-->
     <form name="frm" action="blogedit.asp" method="post" onsubmit="return CheckPost()" style="margin:0px">
                <input name="id" type="hidden" id="id" value="<%=logid%>"/>
                <input name="log_editType" type="hidden" id="log_editType" value="<%=log_editType%>"/>
   				<input id="action" name="action" type="hidden" value="post"/>
                <input name="log_IsDraft" type="hidden" id="log_IsDraft" value="<%=lArticle.logIsDraft%>"/>
   	<div id="MsgContent" style="width:700px">
         <div id="MsgHead">编辑日志</div>
         <div id="MsgBody">        
           <table width="100%" border="0" cellpadding="2" cellspacing="0">
             <tr>
               <td width="72" height="24" align="right" valign="top"><span style="font-weight: bold">标题:</span></td>
               <td align="left"><input name="title" type="text" class="inputBox" id="title" size="50" maxlength="50" value="<%=lArticle.logTitle%>"/>
          &nbsp;&nbsp;移动到 <select name="log_CateID" id="select2">
           <%
	          outCate
	          sub outCate
		            Dim Arr_Category,Category_Len,i
		            Arr_Category=Application(CookieName&"_blog_Category")
		            if ubound(Arr_Category,1)=0 then exit sub
				    Category_Len=ubound(Arr_Category,2)

				  	For i=0 to Category_Len
				   		if not Arr_Category(4,i) then
					  		 if cbool(Arr_Category(10,i)) then
					  		    if stat_ShowHiddenCate and stat_Admin then 
							   		 Response.Write("<option value='"&Arr_Category(0,i)&"'")
							   		 if lArticle.categoryID=int(Arr_Category(0,i)) then Response.Write (" selected")
							   		 Response.Write(">"&Arr_Category(1,i)&"</option>")
							    end if
					  		 else
							   		 Response.Write("<option value='"&Arr_Category(0,i)&"'")
							   		 if lArticle.categoryID=int(Arr_Category(0,i)) then Response.Write (" selected")
							   		 Response.Write(">"&Arr_Category(1,i)&"</option>")
					  		 end if
				   		end if
				  	Next
			  end sub
		  %>
         </select>
   	</td>
               <td width="120" rowspan="3" align="center">
                 <div class="LDialog" style="display:none">
                   <div class="LHead">魔法表情</div>
                   <div class="LBody"><a href="javascript:alert('测试版，该功能暂时不可用')">选择魔法表情</a></div>
               </div></td>
             </tr>
             <tr>
               <td height="24" align="right" valign="top"><span style="font-weight: bold">参数:</span></td>
               <td width="517" align="left">
                 <select name="log_IsShow">
                   <option value="1" <%if lArticle.logIsShow then response.write ("selected=""selected""")%>>公开日志</option>
                   <option value="0" <%if not lArticle.logIsShow then response.write ("selected=""selected""")%>>隐藏日志</option>
                 </select>
                 <select name="log_weather">
                   <option value="sunny" <%if lArticle.logWeather="sunny" then response.write ("selected=""selected""")%>>晴天 </option>
                   <option value="cloudy" <%if lArticle.logWeather="cloudy" then response.write ("selected=""selected""")%>>多云 </option>
                   <option value="flurries" <%if lArticle.logWeather="flurries" then response.write ("selected=""selected""")%>>疾风 </option>
                   <option value="ice" <%if lArticle.logWeather="ice" then response.write ("selected=""selected""")%>>冰雹 </option>
                   <option value="ptcl" <%if lArticle.logWeather="ptcl" then response.write ("selected=""selected""")%>>阴天 </option>
                   <option value="rain" <%if lArticle.logWeather="rain" then response.write ("selected=""selected""")%>>下雨 </option>
                   <option value="showers" <%if lArticle.logWeather="showers" then response.write ("selected=""selected""")%>>阵雨 </option>
                   <option value="snow" <%if lArticle.logWeather="snow" then response.write ("selected=""selected""")%>>下雪 </option>
                 </select>
                 <select name="log_Level">
                   <option value="level1" <%if lArticle.logLevel="level1" then response.write ("selected=""selected""")%>>马马虎虎 </option>
                   <option value="level2" <%if lArticle.logLevel="level2" then response.write ("selected=""selected""")%>>普通 </option>
                   <option value="level3" <%if lArticle.logLevel="level3" then response.write ("selected=""selected""")%>>内容不错 </option>
                   <option value="level4" <%if lArticle.logLevel="level4" then response.write ("selected=""selected""")%>>好东西 </option>
                   <option value="level5" <%if lArticle.logLevel="level5" then response.write ("selected=""selected""")%>>非看不可 </option>
                 </select>
                 <label for="label">
                 <input id="label" name="log_comorder" type="checkbox" value="1" <%if lArticle.logCommentOrder then response.write ("checked=""checked""")%>/>
         评论倒序</label>
                 <label for="label2">
                 <input name="log_DisComment" type="checkbox" id="label2" value="1" <%if lArticle.logDisableComment then response.write ("checked=""checked""")%>/>
         禁止评论</label>
                 <label for="label3">
                 <input name="log_IsTop" type="checkbox" id="label3" value="1" <%if lArticle.logIsTop then response.write ("checked=""checked""")%>/>
         日志置顶</label><br/>
   	            </td>
             </tr>
             <tr>
               <td height="24" align="right" valign="top">&nbsp;</td>
               <td align="left"><span style="font-weight: bold">来自:</span>
                   <input name="log_From" type="text" id="log_From" size="12" class="inputBox" value="<%=lArticle.logFrom%>" />
                   <span style="font-weight: bold">网址:</span>
                   <input name="log_FromURL" type="text" id="log_FromURL" size="38" class="inputBox" value="<%=lArticle.logFromURL%>"/>
                 </td>
             </tr>
            <tr>
              <td height="24" align="right" valign="top"><span style="font-weight: bold">发表时间:</span></td>
              <td align="left">
                 <%if lArticle.logIsDraft then%>
	                  <label for="P1"><input name="PubTimeType" type="radio" id="P1" value="now" size="12" checked/>当前时间</label> 
	                  <label for="P2"><input name="PubTimeType" type="radio" id="P2" value="com" size="12" />自定义日期:</label>
	                  <input name="PubTime" type="text" value="<%=DateToStr(lArticle.logPubTime,"Y-m-d H:I:S")%>" size="21" class="inputBox" /> (格式:yyyy-mm-dd hh:mm:ss)
                 <%else%>
	                  <label for="P2"><input name="PubTimeType" type="radio" id="P2" value="com" size="12" checked/>自定义日期:</label>
	                  <input name="PubTime" type="text" value="<%=DateToStr(lArticle.logPubTime,"Y-m-d H:I:S")%>" size="21" class="inputBox" /> (格式:yyyy-mm-dd hh:mm:ss)
                 <%end if%>
                </td>
            </tr>
            <tr>
              <td height="24" align="right" valign="top"><span style="font-weight: bold">Tags:</span></td>
              <td align="left">
                      <input name="tags" type="text" value="<%=lArticle.logTags%>" size="50" class="inputBox" /> <img src="images/insert.gif" alt="插入已经使用的Tag" onclick="popnew('getTags.asp','tag','250','324')" style="cursor:pointer"/> (tag之间用英文的逗号分割)
               </td>
            </tr>
            <tr>
               <td  align="right" valign="top"><span style="font-weight: bold">内容:</span></td>
               <td colspan="2" align="center"><%
   	if log_editType=0 then 
          Dim sBasePath
          sBasePath = "fckeditor/"
          Dim oFCKeditor
          Set oFCKeditor = New FCKeditor
          oFCKeditor.BasePath	= sBasePath
          oFCKeditor.Config("AutoDetectLanguage") = False
          oFCKeditor.Config("DefaultLanguage")    = "zh-cn"
          oFCKeditor.Value	= UnCheckStr(lArticle.logMessage)
          oFCKeditor.Create "Message"
   	 else
	   	 UBB_TextArea_Height="200px;"
	  	 UBB_AutoHidden=False
	   	 UBB_Msg_Value=UBBFilter(UnCheckStr(lArticle.logMessage))
	   	 UBBeditor("Message")
   	end if
   	%></td>
             </tr>
                       <%if log_editType<>0 then %>          <tr>
               <td align="right" valign="top">&nbsp;</td>
                <td colspan="2" align="left"><label for="label4">
               <label for="label4"><input id="label4" name="log_disImg" type="checkbox" value="1" <%if lArticle.logDisableImage=1 then response.write ("checked")%>/>
   禁止显示图片</label>
                 <label for="label5">
                 <input name="log_DisSM" type="checkbox" id="label5" value="1" <%if lArticle.logDisableSmile=1 then response.write ("checked")%>/>
   禁止表情转换</label>
                 <label for="label6">
                 <input name="log_DisURL" type="checkbox" id="label6" value="1" <%if lArticle.logDisableURL=0 then response.write ("checked")%>/>
   禁止自动转换链接</label>
                <label for="label7">
                 <input name="log_DisKey" type="checkbox" id="label7" value="1" <%if lArticle.logDisableKeyWord=0 then response.write ("checked")%>/>
   禁止自动转换关键字</label></td>
             </tr><%end if%>
             <%
	             Dim UseIntro
	             UseIntro=false
	             UseIntro=CBool(lArticle.logIntroCustom)
             %>
           <tr>
               <td  align="right" valign="top"><span style="font-weight: bold">内容摘要:</span></td>
               <td colspan="2" align="left"><div><label for="shC"><input id="shC" name="log_IntroC" type="checkbox" value="1" onclick="document.getElementById('Div_Intro').style.display=(this.checked)?'block':'none'" <%if not UseIntro then response.write("checked=""checked""")%>/>编辑内容摘要</label></div>
               <div id="Div_Intro" <%if UseIntro then response.write("style=""display:none""")%>>
               <%
               if log_editType=0 then 
                  Dim oFCKeditor1
                  Set oFCKeditor1 = New FCKeditor
                  oFCKeditor1.BasePath	= sBasePath
                  oFCKeditor1.Height="150"
                  oFCKeditor1.ToolbarSet="Basic"
                  oFCKeditor1.Config("AutoDetectLanguage") = False
                  oFCKeditor1.Config("DefaultLanguage")    = "zh-cn"
                  oFCKeditor1.Value	= UnCheckStr(lArticle.logIntro)
                  oFCKeditor1.Create "log_Intro"
               else
   	         %>
   	         <textarea name="log_Intro" class="editTextarea" style="width:99%;height:120px;"><%=UBBFilter(HTMLDecode(UnCheckStr(lArticle.logIntro)))%></textarea>
   	         <%
               end if
               %></div>
               </td>
           </tr>
           <tr>
               <td align="right" valign="top" nowrap><span style="font-weight: bold">附件上传:</span></td>
               <td colspan="2" align="left"><iframe src="attachment.asp" width="100%" height="24" frameborder="0" scrolling="no" border="0" frameborder="0"></iframe></td>
             </tr>
            <tr>
              <td align="right" valign="top"><span style="font-weight: bold">引用通告:</span></td>
              <td colspan="2" align="left"><input name="log_Quote" type="text" size="80" class="inputBox" /><br>请输入网络日志项的引用通告URL。可以用逗号分隔多个引用通告地址.          </td>
            </tr>
            <tr>
               <td colspan="3" align="center">
                 <input name="SaveArticle" type="submit" class="userbutton" value="保存日志" accesskey="S"/>
                 <%if lArticle.logIsDraft then%>
                    <input name="SaveDraft" type="submit" class="userbutton" value="保存并取消草稿" accesskey="D" onclick="document.getElementById('log_IsDraft').value='False'"/>
                 <%end if%>
                 <input name="DelArticle" type="button" class="userbutton" value="删除该日志" accesskey="K" onclick="if (window.confirm('是否要删除该日志')) {document.getElementById('action').value='del';document.forms['frm'].submit()}"/>
                 <input name="CancelEdit" type="button" class="userbutton" value="取消编辑" accesskey="Q" onClick="location='default.asp?id=<%=logid%>'"/>
               </td>
             </tr>
                 <%if lArticle.logIsDraft then%>
	             <tr>
	                <td colspan="3" align="right">
	                友情提示:保存草稿后，日志不会在日志列表中出现。只有再次编辑，<b>取消草稿</b>后才显示出来。</td>
	             </tr>
                 <%end if%>
           </table>
         </div>
   	</div>
   </form>
   <%
	   set lArticle=nothing
	   end if
    end if
    else%>
        <div id="MsgContent" style="width:350px">
         <div id="MsgHead">出错信息</div>
         <div id="MsgBody">
   		 <div class="ErrorIcon"></div>
           <div class="MessageText"><b>你没有没有权限修改日志</b><br/>
           <a href="default.asp">单击返回首页</a> 
   		 </div>
   	 </div>
   	</div>
   <%end if
   end if
 else%>
   <div style="text-align:center;">
    <div id="MsgContent" style="width:300px">
      <div id="MsgHead">日志发表错误</div>
      <div id="MsgBody">
		 <div class="ErrorIcon"></div>
        <div class="MessageText">不允许外部链接提交数据<br/><a href="default.asp">点击返回首页</a>
		 <meta http-equiv="refresh" content="3;url=default.asp"/></div>
	  </div>
	</div>
  </div> 
  <%end if%>
  <br/></div> 
 </div>
<!--#include file="plugins.asp" -->
<!--#include file="footer.asp" -->