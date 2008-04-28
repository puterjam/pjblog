<!--#include file="commond.asp" -->
<!--#include file="header.asp" -->
<!--#include file="common/sha1.asp" -->

<!--内容-->
 <div id="Tbody">
<%
'==================================
'  会员页面
'    更新时间: 2006-1-9
'==================================
Dim blog_Mem
IF Request.QueryString("action")="edit" then
  if memName=empty then Response.Redirect("member.asp")
 %><br/><br/>
   <div style="text-align:center;">
    <div id="MsgContent" style="width:520px">
      <div id="MsgHead">修改用户信息</div>
      <div id="MsgBody">
       <%set blog_Mem=conn.execute("select * from blog_Member where mem_Name='"&CheckStr(memName)&"'")
        if blog_Mem.eof or blog_Mem.bof then
          %>
	     <div class="ErrorIcon"></div>
         <div class="MessageText"  align="center">无法找到该用户信息！！<br/><a href="javascript:history.go(-1)">单击返回</a></div>
         <script>
           document.getElementById("MsgContent").style.width="300px"
           document.getElementById("MsgHead").innerText="错误信息"
         </script>
       <%else%>
	  <table width="100%" cellpadding="0" cellspacing="0">
	  <form name="frm" action="member.asp" method="post" onsubmit="if (this.Oldpassword.value.length<1){alert('请输入你的登录密码!');this.Oldpassword.focus();return false}">
	  <input name="UID" type="hidden" value="<%=blog_Mem("mem_ID")%>"/>
	  <tr><td align="right" width="85"><strong>　昵　称:</strong></td><td align="left" style="padding:3px;"><%=blog_Mem("mem_Name")%></td></tr>
	  <tr><td align="right" width="85"><strong>　旧密码:</strong></td><td align="left" style="padding:3px;"><input name="Oldpassword" type="password" size="18" class="userpass" maxlength="16"/><font color="#FF0000">&nbsp;*</font> 输入你的旧密码.下面的密码输入框为空则不修改密码</td></tr>
	  <tr><td align="right" width="85"><strong>　密　码:</strong></td><td align="left" style="padding:3px;"><input name="password" type="password" size="18" class="userpass" maxlength="16"/> 密码必须是6到16个字符，建议使用英文和符号混合</td></tr>
	  <tr><td align="right" width="85"><strong>密码重复:</strong></td><td align="left" style="padding:3px;"><input name="Confirmpassword" type="password" size="18" class="userpass" maxlength="16"/> 必须和上面的密码一样</td></tr>
	  <tr><td align="right" width="85"><strong>　性　别:</strong></td><td align="left" style="padding:3px;"><input name="Gender" type="radio" value="0" <%if blog_Mem("mem_Sex")=0 then response.write "checked"%>/> 保密 <input name="Gender" type="radio" value="1" <%if blog_Mem("mem_Sex")=1 then response.write "checked"%>/>男 <input name="Gender" type="radio" value="2" <%if blog_Mem("mem_Sex")=2 then response.write "checked"%>/>女</td></tr>
	  <tr><td align="right" width="85"><strong>电子邮件:</strong></td><td align="left" style="padding:3px;"><input name="email" type="text" size="38" class="userpass" maxlength="255" value="<%=blog_Mem("mem_Email")%>"/> <input id="hiddenEmail" name="hiddenEmail" type="checkbox" value="1" <%if blog_Mem("mem_HideEmail") then response.write "checked"%>/> <label for="hiddenEmail">不公开我的电子邮件</label></td></tr>
	  <tr><td align="right" width="85"><strong>个人主页:</strong></td><td align="left" style="padding:3px;"><input name="homepage" type="text" size="38" class="userpass" maxlength="255" value="<%=blog_Mem("mem_HomePage")%>"/></td></tr>
	  <tr><td align="right" width="85"><strong>QQ号码:</strong></td><td align="left" style="padding:3px;"><input name="QQ" type="text" size="10" class="userpass" value="<%=blog_Mem("mem_QQ")%>"/></td></tr>

          <tr>
            <td colspan="2" align="center" style="padding:3px;">
              <input name="action" type="hidden" value="save"/>
			  <input name="submit2" type="submit" class="userbutton" value="修改资料"/>
              <input name="button" type="reset" class="userbutton" value="重写"/></td>
          </tr>
		  </form>
	<%   
        end if
       blog_Mem.Close
       Set blog_Mem=Nothing%>
	  </table>
		</div>
	  </div>
	</div>
<br/><br/>
 <%
ElseIF Request.QueryString("action")="view" then
%>
<br/><br/>
   <div style="text-align:center;">
    <div id="MsgContent" style="width:420px">
      <div id="MsgHead">用户信息</div>
      <div id="MsgBody">
	  <table width="100%" cellpadding="0" cellspacing="0">
	  <%
	   if CheckStr(Request.QueryString("memName"))=Empty then 
         %>
	     <div class="ErrorIcon"></div>
         <div class="MessageText"  align="center">非法操作！！无法完成你的请求！<br/><a href="javascript:history.go(-1)">单击返回</a></div>
         <script>
           document.getElementById("MsgContent").style.width="300px"
           document.getElementById("MsgHead").innerText="错误信息"
         </script>
         <%
       else
        set blog_Mem=conn.execute("select * from blog_Member where mem_Name='"&CheckStr(Request.QueryString("memName"))&"'")
        if blog_Mem.eof or blog_Mem.bof then
          %>
	     <div class="ErrorIcon"></div>
         <div class="MessageText"  align="center">无法找到该用户信息！！<br/><a href="javascript:history.go(-1)">单击返回</a></div>
         <script>
           document.getElementById("MsgContent").style.width="300px"
           document.getElementById("MsgHead").innerText="错误信息"
         </script>
       <%else%>
	  <tr><td align="right" width="85"><strong>　昵　称:</strong></td><td align="left" style="padding:3px;"><%=blog_Mem("mem_Name")%></td></tr>
	  <tr><td align="right" width="85"><strong>　性　别:</strong></td><td align="left" style="padding:3px;"><%
	  select case int(blog_Mem("mem_Sex"))
	    case 1
	     response.write "我是GG"
	    case 2
	     response.write "我是MM"
	    case else
	     response.write "保密"
	  end select
	  %></td></tr>
	  <tr><td align="right" width="85"><strong>电子邮件:</strong></td><td align="left" style="padding:3px;"><%if (blog_Mem("mem_HideEmail") and (not stat_Admin)) or len(blog_Mem("mem_Email"))<1 or isnull(blog_Mem("mem_Email")) then response.write "该用户没有或不公开电子邮件" else response.write blog_Mem("mem_Email") end if%></td></tr>
	  <tr><td align="right" width="85"><strong>个人主页:</strong></td><td align="left" style="padding:3px;"><a href="<%=blog_Mem("mem_HomePage")%>" target="_blank"><%=blog_Mem("mem_HomePage")%></a></td></tr>
	  <tr><td align="right" width="85"><strong>　QQ号码:</strong></td><td align="left" style="padding:3px;"><%=blog_Mem("mem_QQ")%></td></tr>
	  <tr><td align="right" width="85"><strong>统计:</strong></td><td align="left" style="padding:3px;">日志共 <b><%=blog_Mem("mem_PostLogs")%></b> 篇，评论共 <b><%=blog_Mem("mem_PostComms")%></b> 篇，留言共 <b><%=blog_Mem("mem_PostMessageNums")%></b> 个。</td></tr>
          <tr>
            <td colspan="2" align="center" style="padding:3px;">
			  <input type="button" class="userbutton" value="返回" onclick="history.go(-1)"/>
          </tr>
        <%   
        end if
       blog_Mem.Close
       Set blog_Mem=Nothing
    end if
	  %>
	  </table>
		</div>
	  </div>
	</div>
<br/><br/>
<%
ElseIF Request.form("action")="save" then
dim reg
reg=SaveMem
%>
<br/><br/>
   <div style="text-align:center;">
    <div id="MsgContent" style="width:300px">
      <div id="MsgHead"><%=reg(0)%></div>
      <div id="MsgBody">
	   <div class="<%=reg(2)%>"></div>
       <div class="MessageText"><%=reg(1)%></div>
	  </div>
	</div>
  </div><br/><br/>
<%   	
Else
Dim searchType
    Dim PageCount,BM
    Set blog_Mem=Server.CreateObject("ADODB.RecordSet")
    SQL="SELECT * FROM blog_Member order by mem_RegTime desc"
	blog_Mem.Open SQL,Conn,1,1
	SQLQueryNums=SQLQueryNums+1
	blog_Mem.PageSize=20
	blog_Mem.AbsolutePage=CurPage
%><br/><br/>
   <div style="text-align:center;">
    <div id="MsgContent" style="width:480px">
      <div id="MsgHead">用户列表</div>
      <div id="MsgBody" style="padding:0px">
<% if blog_Mem.eof or blog_Mem.bof then
    response.write "没找到任何注册用户!"
   else %>
	  <table width="100%" border="0" cellspacing="0"><tr><th width="80">用户名</th><th>邮件</th><th>主页</th><th>ＱＱ</th><th width="34">日志</th><th width="34">评论</th><th width="34">留言</th><th>注册时间</th></tr>
 <%
 Do Until blog_Mem.EOF OR PageCount=blog_Mem.PageSize
 if blog_Mem("mem_HideEmail") or len(blog_Mem("mem_Email"))<1 or isnull(blog_Mem("mem_Email")) then
   BM="<td><img border=""0"" src=""images/noemail.gif"" alt=""该用户没有或不公开电子邮件""/></td>"
  else
   BM="<td><img border=""0"" src=""images/email.gif"" alt="""&blog_Mem("mem_Email")&"""/></td>"
 end if
 if len(blog_Mem("mem_HomePage"))<1 or isnull(blog_Mem("mem_HomePage")) then
   BM=BM&"<td><img border=""0"" src=""images/nourl.gif"" alt=""该用户没有个人主页""/></td>"
  else
   BM=BM&"<td><a href="""&blog_Mem("mem_HomePage")&""" target=""_blank""><img border=""0"" src=""images/url.gif"" alt="""&blog_Mem("mem_HomePage")&"""/></a></td>"
 end if
 if len(blog_Mem("mem_QQ"))<1 or isnull(blog_Mem("mem_QQ")) then
   BM=BM&"<td><img border=""0"" src=""images/nooicq.gif"" alt=""该用户没有QQ号码""/></td>"
  else
   BM=BM&"<td><a href=""http://wpa.qq.com/msgrd?V=1&Uin="&blog_Mem("mem_QQ")&"&Site="&SiteName&""" target=""_blank""><img border=""0"" src=""images/oicq.gif"" alt=""给该用户发QQ信息""/></a></td>"
 end if   
  response.write "<tr><td align=""left""><a href=""member.asp?action=view&memName="&Server.URLEncode(blog_Mem("mem_Name"))&""">"&blog_Mem("mem_Name")&"</a></td>"&BM&"<td>"&blog_Mem("mem_PostLogs")&"</td><td>"&blog_Mem("mem_PostComms")&"</td><td>"&blog_Mem("mem_PostMessageNums")&"</td><td>"&DateToStr(blog_Mem("mem_RegTime"),"Y-m-d H:I A")&"</td>"
  PageCount=PageCount+1
  blog_Mem.movenext
 loop
 response.write "</table>"
 response.write "<div class=""pageContent"">"&MultiPage(blog_Mem.RecordCount,20,CurPage,"?","","float:left")&"</div>"
 end if
    blog_Mem.Close
    Set blog_Mem=Nothing

%>
	   </div>
	</div>
  </div><br/><br/>
 
 <%End if%>
 </div>
  <!--#include file="footer.asp" -->
  <%
function SaveMem
 dim ReInfo
 dim UID,username,Oldpassword,password,Confirmpassword,Gender,email,homepage,QQ,HideEmail,checkUser
 UID=clng(trim(CheckStr(request.form("UID"))))
 ReInfo=Array("错误信息","","MessageIcon")
 Oldpassword=trim(CheckStr(request.form("Oldpassword")))
 password=trim(CheckStr(request.form("password")))
 Confirmpassword=trim(CheckStr(request.form("Confirmpassword")))
 Gender=CheckStr(request.form("Gender"))
 email=trim(CheckStr(request.form("email")))
 homepage=trim(checkURL(CheckStr(request.form("homepage"))))
 QQ=CheckStr(request.form("QQ"))
 if request.form("hiddenEmail")=1 then 
   HideEmail=true 
  else 
   HideEmail=false 
 end if
 
 if IsInteger(Gender)=false then
	 ReInfo(0)="错误信息"
	 ReInfo(1)="<b>非法操作!</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
	 ReInfo(2)="ErrorIcon"
	 SaveMem=ReInfo
	 exit function
 end if

 set checkUser=conn.execute("select top 1 * from blog_Member where mem_id="&UID&" and mem_Name='"&CheckStr(memName)&"'")
 if checkUser.eof then
	 ReInfo(0)="错误信息"
	 ReInfo(1)="<b>不存在此用户<br/>操作失败！</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
	 ReInfo(2)="ErrorIcon"
	 SaveMem=ReInfo
	 exit function
 end if
 if len(password)>0 then
  if len(password)<6 or len(password)>16 then
	 ReInfo(0)="错误信息"
	 ReInfo(1)="<b>请输入6到16位密码！</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
	 ReInfo(2)="WarningIcon"
	 SaveMem=ReInfo
	 exit function
  end if 
  if password<>Confirmpassword then
	 ReInfo(0)="错误信息"
	 ReInfo(1)="<b>密码验证失败！请重新输入。</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
	 ReInfo(2)="ErrorIcon"
	 SaveMem=ReInfo
	 exit function
  end if  
 end if
 
 if len(QQ)>0 and IsInteger(QQ)=false then
	 ReInfo(0)="错误信息"
	 ReInfo(1)="<b>非法QQ号</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
	 ReInfo(2)="ErrorIcon"
	 SaveMem=ReInfo
	 exit function
 end if
 
 if len(email)>0 and IsValidEmail(email)=false then
	 ReInfo(0)="错误信息"
	 ReInfo(1)="<b>错误的电子邮件地址。</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
	 ReInfo(2)="ErrorIcon"
	 SaveMem=ReInfo
	 exit function
 end if 
   
 set checkUser=conn.execute("select top 1 * from blog_Member where mem_id="&UID&" and mem_Name='"&CheckStr(memName)&"'")
 if checkUser("mem_Password")<>SHA1(Oldpassword&checkUser("mem_salt")) then 
	 ReInfo(0)="错误信息"
  	 ReInfo(1)="<b>用户名与密码错误</b><br/><a href=""javascript:history.go(-1);"">请返回重新输入</a>"
	 ReInfo(2)="ErrorIcon"
	 SaveMem=ReInfo
	 exit function
 end if
 
	Conn.Execute("update blog_member set mem_Sex="&Gender&",mem_Email='"&email&"',mem_HideEmail="&HideEmail&",mem_HomePage='"&homepage&"',mem_QQ='"&QQ&"' where mem_id="&UID&" and mem_Name='"&CheckStr(memName)&"'") 
	SQLQueryNums=SQLQueryNums+1
	if len(password)>0 then
      dim strSalt
      strSalt=randomStr(6)
      password=SHA1(password&strSalt)
	  Conn.Execute("update blog_member set mem_Password='"&password&"',mem_salt='"&strSalt&"' where mem_id="&UID&" and mem_Name='"&CheckStr(memName)&"'")
      SQLQueryNums=SQLQueryNums+1
      logout(true)
	  ReInfo(0)="用户修改成功"
	  ReInfo(1)="<b>你的资料已经修改成功</b><br/>由于你更改了密码所以必须 <a href=""login.asp"">重新登录</a>"
      ReInfo(2)="MessageIcon"
      SaveMem=ReInfo
 	  Session(CookieName&"_LastDo")="EditUser"
      exit function
   end if
	getInfo(2)
	ReInfo(0)="用户修改成功"
	ReInfo(1)="<b>你的资料已经修改成功</b><br/><a href=""default.asp"">返回首页</a>"
	ReInfo(2)="MessageIcon"
    SaveMem=ReInfo	 
 	Session(CookieName&"_LastDo")="EditUser"
  end function	%>