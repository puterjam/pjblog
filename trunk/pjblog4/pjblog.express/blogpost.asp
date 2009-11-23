<!--#include file = "../include.asp" -->
<!--#include file = "../FCKeditor/fckeditor.asp" -->
<!--#include file = "../pjblog.model/cls_ubbconfig.asp" -->
<!--#include file = "../pjblog.model/cls_fso.asp" -->
<!--#include file = "../pjblog.model/cls_Stream.asp" -->
<!--#include file = "../pjblog.data/cls_webconfig.asp" -->
<!--#include file = "../pjblog.model/cls_template.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="UTF-8">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
	<meta http-equiv="Content-Language" content="UTF-8" />
	<meta name="robots" content="all" />
	<meta name="author" content="<%=blog_email%>, <%=blog_master%>" />
	<meta name="Copyright" content="PJBlog3 CopyRight 2008" />
	<meta name="keywords" content="<%=blog_KeyWords%>" />
	<meta name="description" content="<%=blog_Description%>" />
	<meta name="generator" content="PJBlog3" />
	<link rel="service.post" type="application/atom+xml" title="<%=SiteName%> - Atom" href="<%=SiteURL%>xmlrpc.asp" />
	<link rel="EditURI" type="application/rsd+xml" title="RSD" href="<%=SiteURL%>rsd.asp" />
	<title><%=SiteName%> - <%=blog_Title%></title>
	<link rel="alternate" type="application/rss+xml" href="<%=SiteURL%>feed.asp" title="订阅 I Miss You 所有文章(rss2)" />
	<link rel="alternate" type="application/atom+xml" href="<%=SiteURL%>atom.asp" title="订阅 I Miss You 所有文章(atom)" />
	<link rel="stylesheet" rev="stylesheet" href="../pjblog.template/<%=blog_DefaultSkin%>/style/global.css" type="text/css" media="all" /><!--全局样式表-->
	<link rel="stylesheet" rev="stylesheet" href="../pjblog.template/<%=blog_DefaultSkin%>/style/layout.css" type="text/css" media="all" /><!--层次样式表-->
	<link rel="stylesheet" rev="stylesheet" href="../pjblog.template/<%=blog_DefaultSkin%>/style/typography.css" type="text/css" media="all" /><!--局部样式表-->
	<link rel="stylesheet" rev="stylesheet" href="../pjblog.template/<%=blog_DefaultSkin%>/style/link.css" type="text/css" media="all" /><!--超链接样式表-->
	<link rel="stylesheet" rev="stylesheet" href="../pjblog.template/<%=blog_DefaultSkin%>/UBB/editor.css" type="text/css" media="all" /><!--UBB编辑器代码-->
    <link rel="stylesheet" rev="stylesheet" href="../pjblog.common/common.css" type="text/css" media="all" /><!--系统样式-->
	<link rel="stylesheet" rev="stylesheet" href="../FCKeditor/editor/css/Dphighlighter.css" type="text/css" media="all" /><!--FCK块引用&代码样式-->
	<link rel="icon" href="favicon.ico" type="image/x-icon" />
	<link rel="shortcut icon" href="../favicon.ico" type="image/x-icon" />
    <script language="javascript" type="text/javascript" src="../pjblog.common/language.js"></script>
	<script language="javascript" type="text/javascript" src="../pjblog.common/common.js"></script>
    <script language="javascript" type="text/javascript" src="../pjblog.common/Ajax.js"></script>
	<script language="javascript" type="text/javascript" src="../pjblog.common/upload.js"></script>
    <script language="javascript" type="text/javascript" src="../pjblog.common/pinyin.js"></script>
</head>
<body>
<a href="default.asp" accesskey="i"></a>
<a href="javascript:history.go(-1)" accesskey="z"></a>
<div id="container">
 <!--顶部-->
    <div id="header">
        <div id="blogname"><%=SiteName%><div id="blogTitle"><%=blog_Title%></div></div>
        <div id="menu">
            <div id="Left"></div><div id="Right"></div>
            <%Init.HeadNav%>
        </div>
 	</div>
    
 	<div id="Tbody">
  		<div style="text-align:center;">
  		<br/>
  <%
'==================================
'  发表日志页面
'    更新时间: 2006-1-22
'==================================
Dim preLog, nextLog
If Asp.ChkPost() Then
    If stat_AddAll <> True And stat_Add <> True Then
%>
		<div id="MsgContent" style="width:350px">
			<div id="MsgHead"><%=lang.Set.Asp(6)%></div>
			<div id="MsgBody">
				<div class="ErrorIcon"></div>
				<div class="MessageText"><b><%=lang.Set.Asp(25)%></b><br/>
				<a href="default.asp"><%=lang.Set.Asp(15)%></a><%if memName=Empty Then %> | <a href="login.asp"><%=lang.Set.Asp(21)%></a><%end if%></div>
  	 		</div>
		</div>
<%
	else
	If Request.QueryString("action") = "complete" Then
%>
		 <div id="MsgContent" style="width:400px; height:120px;">
		     <div id="MsgHead"><%=lang.Set.Asp(26)%></div>
		     <div id="MsgBody">
		  		 <div class="<%if Request.QueryString("suc") = "false" Then response.write "ErrorIcon" else response.write "MessageIcon"%>"></div>
		         <div class="MessageText" style="float:right; display:block; text-align:left; padding-top:10px;">
				 	<%
						If Request.QueryString("suc") = "true" Then
							Response.Write(lang.Set.Asp(95))
						Else
							Response.Write(Request.QueryString("info"))
						End If
					%>
                    <br/>
		  		 	<%If Request.QueryString("suc") = "true" Then %>
			  		 <input type="button" value="<%=lang.Set.Asp(31)%>" onClick="location='?action=Static&id=<%=Request.QueryString("id")%>'" class="userbutton" /><br/>
			     <%end if%>
                 </div>
		  	  	</div>
			</div>
		</div>
<%
	Else
		If Request.Form("log_CateID") = Empty Then
%>
   <!--第一步-->
   <script language="javascript">
    function chkFrm(){
	   if (document.forms["frm"].log_CateID.value == ""){
		  alert(lang.Set.Js(1));
		  return false;
	   }
	   return true;
    }
   </script>
    <form name="frm" action="blogpost.asp" method="post" onSubmit="return chkFrm()">
      <div id="MsgContent" style="width:350px">
        	<div id="MsgHead"><%=lang.Set.Asp(27)%> - <%=lang.Set.Js(1)%></div>
        	<div id="MsgBody">
                <table width="100%" border="0" cellpadding="3" cellspacing="0">
                    <tr><td colspan="2" align="left" style=" padding:2px 2px 2px 2px">
                    <ol>
                        <li>1</li>
                        <li>2</li>
                        <li>3</li>
                    </ol>
                    </td></tr>
                    <tr>
                        <td width="100" align="right"><span style="font-weight: bold"><%=lang.Set.Js(1)%>:</span></td>
                        <td align="left"><span style="font-weight: bold">
                        <select name="log_CateID" id="select2">
                            <option value="" selected="selected" style="color:#333"><%=lang.Set.Js(1)%></option>
<%
	Dim CateArr, Arr, Arri
	CateArr = Cache.CategoryCache(1)
	If UBound(CateArr, 1) = 0 Then
		Arr = ""
	Else
		Arr = ""
		For Arri = 0 To UBound(CateArr, 2)
			If Not CateArr(4, Arri) Then Arr = Arr & "<option value=""" & CateArr(0, Arri) & """>" & CateArr(1, Arri) & "</option>"
		Next
	End If
	Response.Write(Arr)
%>
                        </select>
                        </span></td>
                    </tr>
                    <tr>
                        <td align="right"><span style="font-weight: bold"><%=lang.Set.Asp(28)%>:</span></td>
                        <td align="left">
                        <label title="<%=lang.Set.Asp(29)%>" for="ET1" accesskey="U"><input type="radio" id="ET1" name="log_editType" value="1" checked="checked" /><%=lang.Set.Asp(29)%></label>
                        <label title="<%=lang.Set.Asp(30)%>" for="ET2" accesskey="K"><input name="log_editType" type="radio" id="ET2" value="0" /><%=lang.Set.Asp(30)%></label></td>
                    </tr>
                    <tr>
                        <td colspan="2" align="center"><input name="submit" type="submit" class="userbutton" value="<%=lang.Set.Asp(31)%>" accesskey="N"/> <input name="button" type="button" class="userbutton" value="<%=lang.Set.Asp(15)%>" onClick="location='default.asp'" accesskey="Q"/></td>
                    </tr>
                </table>
  	  		</div>
  	  </div>
  	</form>
  <%Else
    Dim log_editType, editTs, CateRequestName
    log_editType = Request.Form("log_editType")
	CateRequestName = Conn.ExeCute("SELECT cate_Name FROM blog_Category WHERE cate_ID=" & Request.Form("log_CateID"))(0)
%>
  <!--第二步-->
    <form name="frm" action="../pjblog.logic/log_Article.asp?action=add" method="post">
      		    <input name="log_CateID" type="hidden" id="log_CateID" value="<%=Request.Form("log_CateID")%>"/>
                <input name="log_editType" type="hidden" id="log_editType" value="<%=log_editType%>"/>
                <input name="log_IsDraft" type="hidden" id="log_IsDraft" value="False"/>
  	<div id="MsgContent" style="width:700px">
        <div id="MsgHead"><%=lang.Set.Asp(32)(CateRequestName)%></div>
        <div id="MsgBody">
          <table width="100%" border="0" cellpadding="2" cellspacing="0">
            <tr>
              <td width="76" height="24" align="right" valign="top"><span style="font-weight: bold"><%=lang.Set.Asp(33)%>:</span></td>
              <td align="left"><input name="log_Title" type="text" class="inputBox" id="log_Title" size="50" maxlength="50"/>
              </td>
            </tr>
			<tr>
              <td height="24" align="right" valign="top"><span style="font-weight: bold"><%=lang.Set.Asp(34)%>:</span></td>
              <td align="left">
			  <input name="log_cname" type="text" class="inputBox" id="log_cname" size="30" maxlength="50"/>
			   <span> . </span>
			  <select name="log_ctype" onBlur="getAlias()" id="log_ctype">
			  <%
			  		Dim blog_html_option, blog_html_option_i
					blog_html_option = Split(blog_html, ",")
					For blog_html_option_i = 0 To Ubound(blog_html_option)
						Response.Write("<option value=""" & blog_html_option(blog_html_option_i) & """>" & blog_html_option(blog_html_option_i) & "</option>")
					Next
			  %>
			  </select> <span id="CheckAlias"></span>&nbsp;&nbsp;<span><a href="javascript:void(0)" onClick="$('log_cname').value= pinyin.go($('log_Title').value)"><%=lang.Set.Asp(36)%></a> &nbsp;&nbsp;<a href="javascript:void(0)" onClick="$('log_cname').value= pinyin.go($('log_Title').value,1)"><%=lang.Set.Asp(37)%></a></span>
              </td>
            </tr>
            <tr>
              <td align="right" valign="top"><span style="font-weight: bold"><%=lang.Set.Asp(35)%>:</span></td>
              <td align="left">
              	<select name="log_IsShow">
                  <option value="1" selected="selected"><%=lang.Set.Asp(93)%></option>
                  <option value="0"><%=lang.Set.Asp(94)%></option>
                </select>
                <select name="log_weather" id="logweather">
                  <option value="sunny" selected="selected"><%=lang.Set.Asp(38)%> </option>
                  <option value="cloudy"><%=lang.Set.Asp(39)%> </option>
                  <option value="flurries"><%=lang.Set.Asp(40)%> </option>
                  <option value="ice"><%=lang.Set.Asp(41)%> </option>
                  <option value="ptcl"><%=lang.Set.Asp(42)%> </option>
                  <option value="rain"><%=lang.Set.Asp(43)%> </option>
                  <option value="showers"><%=lang.Set.Asp(44)%> </option>
                  <option value="snow"><%=lang.Set.Asp(45)%> </option>
                </select>
                <select name="log_Level" id="logLevel">
                  <option value="level1">★</option>
                  <option value="level2">★★</option>
                  <option value="level3" selected="selected">★★★</option>
                  <option value="level4">★★★★</option>
                  <option value="level5">★★★★★</option>
                </select>
                <label for="label"><input id="label" name="log_comorder" type="checkbox" value="1" checked="checked" /><%=lang.Set.Asp(46)%></label>
                <label for="label2"><input name="log_DisComment" type="checkbox" id="label2" value="1" /><%=lang.Set.Asp(47)%></label>
                <label for="label3"><input name="log_IsTop" type="checkbox" id="label3" value="1" /><%=lang.Set.Asp(48)%></label>
              </td>
            </tr>
			<tr>
               <td align="right" valign="top"><span style="font-weight: bold"><%=lang.Set.Asp(49)%>:</span></td>
               <td align="left">
               		<div>
	 					<label for="Meta"><input id="log_Meta" name="log_Meta" type="checkbox" value="1" onClick="document.getElementById('Div_Meta').style.display=(this.checked)?'block':'none'" /><%=lang.Set.Asp(51)%></label>
                    </div>
                    
	                <div id="Div_Meta" style="display:none;" class="tips_body">
      	 				- <%=lang.Set.Asp(58)%><br/>
		                <span style="font-weight: bold"><%=lang.Set.Asp(59)%>&nbsp;&nbsp;:</span>
						<input name="log_KeyWords" type="text" class="inputBox" id="log_KeyWords" title="<%=lang.Set.Asp(60)%>" size="80" maxlength="254" />
						<br />
						<span style="font-weight: bold"><%=lang.Set.Asp(61)%>:</span>
						<input name="log_Description" type="text" class="inputBox" id="log_Description" title="<%=lang.Set.Asp(62)%>" size="80" maxlength="254" />
	                </div>
				  </td>
             </tr>
            <tr>
              <td height="24" align="right" valign="top"><b><%=lang.Set.Asp(63)%>:</b></td>
              <td align="left"><span style="font-weight: bold"></span>
                  <input name="log_From" type="text" id="log_From" value="<%=lang.Set.Asp(64)%>" size="12" class="inputBox" />
                  <span style="font-weight: bold"><%=lang.Set.Asp(65)%>:</span>
                  <input name="log_FromURL" type="text" id="log_FromURL" value="<%=siteURL%>" size="38" class="inputBox" />
                </td>
            </tr>
            <tr>
              <td height="24" align="right" valign="top"><span style="font-weight: bold"><%=lang.Set.Asp(66)%>:</span></td>
              <td align="left">
                  <label for="P1"><input name="PubTimeType" type="radio" id="P1" value="now" size="12" checked/><%=lang.Set.Asp(67)%></label> 
                  <label for="P2"><input name="PubTimeType" type="radio" id="P2" value="com" size="12" /><%=lang.Set.Asp(68)%>:</label>
                  <input onFocus="this.select();$('P2').checked='checked'" name="PubTime" type="text" value="<%=Asp.DateToStr(now(),"Y-m-d H:I:S")%>" size="21" class="inputBox" /> (<%=lang.Set.Asp(69)%>)
                </td>
            </tr>
            <tr>
              <td height="24" align="right" valign="top"><span style="font-weight: bold"><%=lang.Set.Asp(70)%>:</span></td>
              <td align="left">
                      <input name="log_tag" type="text" value="" size="50" class="inputBox" id="log_tag" /> <img src="../images/insert.gif" alt="<%=lang.Set.Asp(71)%>" onClick="popnew('getTags.asp','tag','250','324')" style="cursor:pointer"/> (<%=lang.Set.Asp(72)%>)
               </td>
            </tr>
             <tr>
              <td  align="right" valign="top"><span style="font-weight: bold"><%=lang.Set.Asp(73)%>:</span></td>
              <td align="center"><%
If log_editType = 0 Then
    Dim sBasePath
    sBasePath = "../FCKeditor/"
    Dim oFCKeditor
    Set oFCKeditor = New FCKeditor
    oFCKeditor.BasePath = sBasePath
    oFCKeditor.Config("AutoDetectLanguage") = False
    oFCKeditor.Config("DefaultLanguage") = "zh-cn"
    oFCKeditor.Value = ""
    oFCKeditor.Height = "350"
    oFCKeditor.Create "Message"
Else
    UBB_TextArea_Height = "200px;"
    UBB_AutoHidden = False
    UBBeditor("Message")
End If

%></td>
            </tr>
            <tr>
              <td align="right" valign="top">&nbsp;</td>
               <td align="left">
 			<%if log_editType<>0 then %>
                <label for="label4"><input id="label4" name="log_disImg" type="checkbox" value="1" /><%=lang.Set.Asp(74)%></label>
                <label for="label5"><input name="log_DisSM" type="checkbox" id="label5" value="1" /><%=lang.Set.Asp(75)%></label>
                <label for="label6"><input name="log_DisURL" type="checkbox" id="label6" value="1" /><%=lang.Set.Asp(76)%></label>
                <label for="label7"><input name="log_DisKey" type="checkbox" id="label7" value="1" /><%=lang.Set.Asp(77)%></label>
  			<%else%>
                <strong>[&nbsp;&nbsp;<a herf="#" onClick="GetLength();" style="cursor:pointer"><%=lang.Set.Asp(78)%></a>&nbsp;&nbsp;|&nbsp;&nbsp;<a herf="#" onClick="SetContents();" style="cursor:pointer"><%=lang.Set.Asp(79)%></a>&nbsp;&nbsp;]</strong>
  			<%end if%>
            	</td></tr>
          <tr>
              <td align="right" valign="top"><span style="font-weight: bold"><%=lang.Set.Asp(80)%>:</span></td>
              <td align="left"><div><label for="shC"><input id="shC" name="log_IntroC" type="checkbox" value="1" onClick="document.getElementById('Div_Intro').style.display=(this.checked)?'block':'none'"/><%=lang.Set.Asp(81)%></label></div>
              <div id="Div_Intro" style="display:none">
              <%
If log_editType = 0 Then
    Dim oFCKeditor1
    Set oFCKeditor1 = New FCKeditor
    oFCKeditor1.BasePath = sBasePath
    oFCKeditor1.Height = "150"
    oFCKeditor1.ToolbarSet = "Basic"
    oFCKeditor1.Config("AutoDetectLanguage") = False
    oFCKeditor1.Config("DefaultLanguage") = "zh-cn"
    oFCKeditor1.Value = ""
    oFCKeditor1.Create "log_Intro"
Else

%>
  	         <textarea name="log_Intro" class="editTextarea" style="width:99%;height:120px;"></textarea>
  	         <%
End If

%></div>
              </td>
          </tr>          <tr>
              <td align="right" valign="top" nowrap><span style="font-weight: bold"><%=lang.Set.Asp(82)%>:</span></td>
              <script language="javascript">
			  	temp = 0;
			  </script>
              <td align="left"><input type="hidden" value="" id="upload" name="UploadFiles"/><label for="memberCanDown"><input type="checkbox" id="memberCanDown" onClick="if (this.checked){temp = 1;}else{temp = 0;}" />只允许会员下载</label>&nbsp;&nbsp;<a href="javascript:void(0);" onClick="Upload.open(this, 'Message', 0, '../upload/<%="month_" & Asp.DateToStr(Now(), "ym")%>', 10)" onFocus="this.blur();">点击上传附件</a><span style=" margin-left:10px; color:#ccc">支持多附件批量上传, 附件统一管理...</span></td>
            </tr>
            <tr>
              <td align="right" valign="top"><span style="font-weight: bold"><%=lang.Set.Asp(83)%>:</span></td>
              <td align="left"><input name="log_Quote" type="text" size="80" class="inputBox" id="log_Quote"/><br><%=lang.Set.Asp(84)%></td>
            </tr>
            <tr>
              <td colspan="2" align="center">
                <input name="SaveArticle" type="submit" class="userbutton" value="<%=lang.Set.Asp(86)%>" accesskey="S"/>
                <input name="SaveDraft" type="submit" class="userbutton" value="<%=lang.Set.Asp(88)%>" accesskey="D" onClick="document.getElementById('log_IsDraft').value='True'"/>
                <input name="ReturnButton" type="button" class="userbutton" value="<%=lang.Set.Asp(15)%>" accesskey="Q" onClick="history.go(-1)"/></td>
            </tr>
            <tr>
              <td colspan="2" align="right">
                 <%=lang.Set.Asp(90)%></td>
            </tr>
           
           </table>
        </div>
  	</div>
  </form>
  <%
		End If
	End If
End If
ElseIf Request.QueryString("action") = "Static" Then
	Call web.default
	Dim Rs
	Set Rs = Conn.Execute("Select T.cate_ID, T.cate_Folder From blog_Category AS T, blog_Content AS B Where B.log_ID=" & Asp.CheckStr(Request.QueryString("id")) & " And B.log_CateID=T.cate_ID")
	If Not (Rs.Bof And Rs.Eof) Then
		Call web.category(Rs(0).value, Rs(1).value)
	End If
	Set Rs = Nothing
%>
	 <div id="MsgContent" style="width:500px;">
		     <div id="MsgHead"><%=lang.Set.Asp(97)%></div>
		     <div id="MsgBody">
		  		 <div class="MessageIcon"></div>
		         <div class="MessageText">
                 <%=lang.Set.Asp(98)%>
                    <br/>
		  		 	<a href="../default.asp"><%=lang.Set.Asp(15)%></a> | <a href="../default.asp?id=<%=Asp.CheckStr(Request.QueryString("id"))%>"><%=lang.Set.Asp(92)%></a>
                    <meta http-equiv="refresh" content="3;url=default.asp"/>
                 </div>
		  	  	</div>
			</div>
		</div>
<%
Else
%>
   <div style="text-align:center;">
    <div id="MsgContent" style="width:300px">
      <div id="MsgHead"><%=lang.Set.Asp(26)%></div>
      <div id="MsgBody">
		 <div class="ErrorIcon"></div>
        <div class="MessageText"><%=lang.Set.Asp(96)%><br/><a href="default.asp"><%=lang.Set.Asp(15)%></a>
		 <meta http-equiv="refresh" content="3;url=default.asp"/></div>
	  </div>
	</div>
  </div> 
  <%end if%><br/>
 </div> 
</div>


<!--底部-->
  <div id="foot">
    <p>Powered By <a href="http://www.pjhome.net" target="_blank"><strong>PJBlog3</strong></a> <a href="http://www.pjhome.net" target="_blank"><strong>V<%=Sys.version%></strong></a> CopyRight 2005 - 2009, <strong><%=SiteName%></strong> <a href="http://validator.w3.org/check/referer" target="_blank">xhtml</a> | <a href="http://jigsaw.w3.org/css-validator/validator-uri.html">css</a></p>
    <p style="font-size:11px;">Processed in <b><%=FormatNumber(Timer()-StartTime,6,-1)%></b> second(s) , <b><%'=SQLQueryNums%></b> queries<%'=SkinInfo%>
    <br/><a href="http://www.miibeian.gov.cn" style="font-size:12px" target="_blank"><b><%=blogabout%></b></a>
    </p>
   </div>
</div>
</body>
</html><%Sys.Close%>