<!--#include file="commond.asp" -->
<!--#include file="header.asp" -->
<!--#include file="common/UBBconfig.asp" -->
<!--#include file="FCKeditor/fckeditor.asp" -->
<!--#include file="common/ModSet.asp" -->
<!--#include file="class/cls_logAction.asp" -->
<div id="Tbody">
  <div style="text-align:center;">
  <br/>
  <%
 '==================================
'  发表日志页面
'    更新时间: 2006-1-22
'==================================
dim preLog,nextLog
 if ChkPost() then
  IF stat_AddAll<>true and stat_Add<>true Then%>
      <div id="MsgContent" style="width:350px">
        <div id="MsgHead">出错信息</div>
        <div id="MsgBody">
  		 <div class="ErrorIcon"></div>
          <div class="MessageText"><b>你没有没有权限发表新日志</b><br/>
          <a href="default.asp">单击返回首页</a><%if memName=Empty then %> | <a href="login.asp">登录</a><%end if%>	 
  		 </div>
  	 </div>
  	</div>
  <%else%>
   <!--内容-->
  <%IF Request.Form("action")="post" Then
      dim lArticle,postLog
      set lArticle=new logArticle
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
	 	 
	 	 postLog = lArticle.postLog
    set lArticle=nothing

		  %>
		      <div id="MsgContent" style="width:300px">
		        <div id="MsgHead">反馈信息</div>
		        <div id="MsgBody">
		  		 <div class="<%if postLog(0)<0 then response.write "ErrorIcon" else response.write "MessageIcon"%>"></div>
		          <div class="MessageText"><%=postLog(1)%><br/><a href="default.asp">点击返回首页</a><br/>
		  		 <%if postLog(0)>=0 then %>
			  		 <a href="default.asp?id=<%=postLog(2)%>">返回你所发表的日志</a><br/>
			  		 <meta http-equiv="refresh" content="3;url=default.asp?logID=<%=postLog(2)%>"/>
			     <%end if%>
		  	  </div>
		  	</div>
		    </div>
		    <%
  else
   IF Request.Form("log_CateID")=Empty Then%>
   <!--第一步-->
   <script language="javascript">
    function chkFrm(){
     if (document.forms["frm"].log_CateID.value=="") {
      alert("请选择日志分类")
  	return false
     }
     return true
    }
   </script>
    <form name="frm" action="blogpost.asp" method="post" onsubmit="return chkFrm()">
      <div id="MsgContent" style="width:350px">
        <div id="MsgHead">发表日志 - 选择分类</div>
        <div id="MsgBody">
          <table width="100%" border="0" cellpadding="3" cellspacing="0">
    <tr>
      <td width="100" align="right"><span style="font-weight: bold">选择日志分类:</span></td>
      <td align="left"><span style="font-weight: bold">
        <select name="log_CateID" id="select2">
          <option value="" selected="selected" style="color:#333">请选择分类</option>
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
					  		if stat_ShowHiddenCate and stat_Admin then Response.Write("<option value='"&Arr_Category(0,i)&"'>&nbsp;&nbsp;&nbsp;"&Arr_Category(1,i)&"&nbsp;["&Arr_Category(7,i)&"]&nbsp;&nbsp;</option>")
		  	          else
						  	Response.Write("<option value='"&Arr_Category(0,i)&"'>&nbsp;&nbsp;&nbsp;"&Arr_Category(1,i)&"&nbsp;["&Arr_Category(7,i)&"]&nbsp;&nbsp;</option>")
		  	         end if
		  		end if
		  	Next
		  end sub
		  %>
        </select>
      </span></td>
    </tr>
    <tr>
      <td align="right"><span style="font-weight: bold">选择编辑类型:</span></td>
      <td align="left"><label title="UBB编辑器" for="ET1" accesskey="U"><input type="radio" id="ET1" name="log_editType" value="1" checked="checked" />UBBeditor</label>         <label title="FCK在线编辑器" for="ET2" accesskey="K">
          <input name="log_editType" type="radio" id="ET2" value="0" />         
          FCKeditor</label></td>
    </tr>
    <tr>
    <td colspan="2" align="center"><input name="submit" type="submit" class="userbutton" value="下一步" accesskey="N"/> <input name="button" type="button" class="userbutton" value="返回主页" onclick="location='default.asp'" accesskey="Q"/></td>
    </tr>
  </table>
  
  	  </div>
  	</div>
  </form>
  <%else
  	dim log_editType,editTs
  	log_editType=Request.Form("log_editType")
  %>
  <!--第二步-->
    <form name="frm" action="blogpost.asp" method="post" onsubmit="return CheckPost()">
      		    <input name="log_CateID" type="hidden" id="log_CateID" value="<%=Request.Form("log_CateID")%>"/>
                <input name="log_editType" type="hidden" id="log_editType" value="<%=log_editType%>"/>
  				<input name="action" type="hidden" value="post"/>
                <input name="log_IsDraft" type="hidden" id="log_IsDraft" value="False"/>
  	<div id="MsgContent" style="width:700px">
        <div id="MsgHead">在 【<%=Conn.ExeCute("SELECT cate_Name FROM blog_Category WHERE cate_ID="&Request.Form("log_CateID")&"")(0)%>】 发表日志</div>
        <div id="MsgBody">        
          <table width="100%" border="0" cellpadding="2" cellspacing="0">
            <tr>
              <td width="72" height="24" align="right" valign="top"><span style="font-weight: bold">标题:</span></td>
              <td align="left"><input name="title" type="text" class="inputBox" id="title" size="50" maxlength="50"/>
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
                  <option value="1" selected="selected">公开日志</option>
                  <option value="0">隐藏日志</option>
                </select>
                <select name="log_weather">
                  <option value="sunny" selected="selected">晴天 </option>
                  <option value="cloudy">多云 </option>
                  <option value="flurries">疾风 </option>
                  <option value="ice">冰雹 </option>
                  <option value="ptcl">阴天 </option>
                  <option value="rain">下雨 </option>
                  <option value="showers">阵雨 </option>
                  <option value="snow">下雪 </option>
                </select>
                <select name="log_Level">
                  <option value="level1">马马虎虎 </option>
                  <option value="level2">普通 </option>
                  <option value="level3" selected="selected">内容不错 </option>
                  <option value="level4">好东西 </option>
                  <option value="level5">非看不可 </option>
                </select>
                <label for="label">
                <input id="label" name="log_comorder" type="checkbox" value="1" checked="checked" />
        评论倒序</label>
                <label for="label2">
                <input name="log_DisComment" type="checkbox" id="label2" value="1" />
        禁止评论</label>
                <label for="label3">
                <input name="log_IsTop" type="checkbox" id="label3" value="1" />
        日志置顶</label>
              </td>
            </tr>
            <tr>
              <td height="24" align="right" valign="top">&nbsp;</td>
              <td align="left"><span style="font-weight: bold">来自:</span>
                  <input name="log_From" type="text" id="log_From" value="本站原创" size="12" class="inputBox" />
                  <span style="font-weight: bold">网址:</span>
                  <input name="log_FromURL" type="text" id="log_FromURL" value="<%=siteURL%>" size="38" class="inputBox" />
                </td>
            </tr>
            <tr>
              <td height="24" align="right" valign="top"><span style="font-weight: bold">发表时间:</span></td>
              <td align="left">
                  <label for="P1"><input name="PubTimeType" type="radio" id="P1" value="now" size="12" checked/>当前时间</label> 
                  <label for="P2"><input name="PubTimeType" type="radio" id="P2" value="com" size="12" />自定义日期:</label>
                  <input name="PubTime" type="text" value="<%=DateToStr(now(),"Y-m-d H:I:S")%>" size="21" class="inputBox" /> (格式:yyyy-mm-dd hh:mm:ss)
                </td>
            </tr>
            <tr>
              <td height="24" align="right" valign="top"><span style="font-weight: bold">Tags:</span></td>
              <td align="left">
                      <input name="tags" type="text" value="" size="50" class="inputBox" /> <img src="images/insert.gif" alt="插入已经使用的Tag" onclick="popnew('getTags.asp','tag','250','324')" style="cursor:pointer"/> (tag之间用英文的逗号分割)
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
         oFCKeditor.Value	= ""
         oFCKeditor.Create "Message"
  	 else
	  	UBB_TextArea_Height="200px;"
	  	UBB_AutoHidden=False
		UBBeditor("Message")
  	end if
  	%></td>
            </tr>
  <%if log_editType<>0 then %>          <tr>
              <td align="right" valign="top">&nbsp;</td>
               <td colspan="2" align="left"><label for="label4">
              <label for="label4"><input id="label4" name="log_disImg" type="checkbox" value="1" />
  禁止显示图片</label>
                <label for="label5">
                <input name="log_DisSM" type="checkbox" id="label5" value="1" />
  禁止表情转换</label>
                <label for="label6">
                <input name="log_DisURL" type="checkbox" id="label6" value="1" />
  禁止自动转换链接</label>
               <label for="label7">
                <input name="log_DisKey" type="checkbox" id="label7" value="1" />
  禁止自动转换关键字</label></td>
            </tr><%end if%>
          <tr>
              <td  align="right" valign="top"><span style="font-weight: bold">内容摘要:</span></td>
              <td colspan="2" align="left"><div><label for="shC"><input id="shC" name="log_IntroC" type="checkbox" value="1" onclick="document.getElementById('Div_Intro').style.display=(this.checked)?'block':'none'"/>编辑内容摘要</label></div>
              <div id="Div_Intro" style="display:none">
              <%
              if log_editType=0 then 
                 Dim oFCKeditor1
                 Set oFCKeditor1 = New FCKeditor
                 oFCKeditor1.BasePath	= sBasePath
                 oFCKeditor1.Height="150"
                 oFCKeditor1.ToolbarSet="Basic"
                 oFCKeditor1.Config("AutoDetectLanguage") = False
                 oFCKeditor1.Config("DefaultLanguage")    = "zh-cn"
                 oFCKeditor1.Value	= ""
                 oFCKeditor1.Create "log_Intro"
              else
  	         %>
  	         <textarea name="log_Intro" class="editTextarea" style="width:99%;height:120px;"></textarea>
  	         <%
              end if
              %></div>
              </td>
          </tr>          <tr>
              <td align="right" valign="top" nowrap><span style="font-weight: bold">附件上传:</span></td>
              <td colspan="2" align="left"><iframe src="attachment.asp" width="100%" height="24" frameborder="0" scrolling="no" border="0" frameborder="0"></iframe></td>
            </tr>
            <tr>
              <td align="right" valign="top"><span style="font-weight: bold">引用通告:</span></td>
              <td colspan="2" align="left"><input name="log_Quote" type="text" size="80" class="inputBox" /><br>请输入网络日志项的引用通告URL。可以用逗号分隔多个引用通告地址.          </td>
            </tr>
            <tr>
              <td colspan="3" align="center">
                <input name="SaveArticle" type="submit" class="userbutton" value="提交日志" accesskey="S"/>
                <input name="SaveDraft" type="submit" class="userbutton" value="保存为草稿" accesskey="D" onclick="document.getElementById('log_IsDraft').value='True'"/>
                <input name="ReturnButton" type="button" class="userbutton" value="返回" accesskey="Q" onClick="history.go(-1)"/></td>
            </tr>
            <tr>
              <td colspan="3" align="right">
                友情提示:保存草稿后，日志不会在日志列表中出现。只有再次编辑，<b>取消草稿</b>后才显示出来。</td>
            </tr>
           
           </table>
        </div>
  	</div>
  </form>
  <%
    end if
   end if
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
  <%end if%><br/> 
 </div> 
</div>
<!--#include file="plugins.asp" -->
<!--#include file="footer.asp" -->