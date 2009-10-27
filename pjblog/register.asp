<!--#include file="commond.asp" -->
<!--#include file="header.asp" -->
<!--#include file="common/sha1.asp" -->
<!--内容-->
<%
'==================================
'  用户注册页面
'    更新时间: 2006-5-29
'==================================
If blog_Disregister Then showmsg lang.Tip.SysTem(1), "站点不允许注册新用户<br/><a href=""default.asp"">" & lang.Tip.SysTem(2) & "</a>", "ErrorIcon", ""
%>

 <div id="Tbody">
<%
dim Referer_Url
If Request.QueryString("action") = "agree" Then
    logout(True)

    Referer_Url = Session(CookieName & "_Register_Referer_Url")
    If len(Referer_Url) < 8 then Referer_Url = Cstr(Request.ServerVariables("HTTP_REFERER"))
    If len(Referer_Url) < 8 then Referer_Url = "http://" & Request.ServerVariables("HTTP_HOST")
%><br/><br/>
   <div style="text-align:center;">
    <div id="MsgContent" style="width:520px">
      <div id="MsgHead"><%=lang.Action.Register%></div>
      <div id="MsgBody">
	  <table width="100%" cellpadding="0" cellspacing="0">
	  <form name="frm" action="register.asp" method="post">
	  <tr><td align="right" width="85"><strong>　<%=lang.MemBer.EditForm(4)%>:</strong></td><td align="left" style="padding:3px;"><input name="username" type="text" size="18" class="userpass" maxlength="24" onblur="if (this.value.length != 0) {CheckName();}" onfocus="if (this.value.length != 0) {CheckName();}" onclick="if (this.value.length != 0) {CheckName();}" id="vs"/><input id="PostBack_UserName" value="False|$|False" type="hidden"><font color="#FF0000">&nbsp;*</font> <%=lang.MemBer.EditForm(48)%> <span id="CheckName"></span></td></tr>
	  <tr><td align="right" width="85"><strong>　<%=lang.Action.LoginForm(3)%>:</strong></td><td align="left" style="padding:3px;"><input name="password" type="password" size="18" class="userpass" maxlength="16" id="cpassword" onkeyup="istrong()"/><font color="#FF0000">&nbsp;*</font> <%=lang.MemBer.EditForm(7)%></td></tr>
      <tr><td align="right" width="85"><strong><%=lang.MemBer.EditForm(8)%>:</strong></td><td align="left">&nbsp;<img src="images/0.gif" id="strong"></td></tr>
	  <tr><td align="right" width="85"><strong><%=lang.MemBer.EditForm(9)%>:</strong></td><td align="left" style="padding:3px;"><input name="Confirmpassword" type="password" size="18" class="userpass" maxlength="16" onblur="if (this.value.length != 0) {CheckPwd();}" onfocus="if (this.value.length != 0) {CheckPwd();}" onclick="if (this.value.length != 0) {CheckPwd();}" id="cConfirmpassword"/><font color="#FF0000">&nbsp;*</font> <%=lang.MemBer.EditForm(10)%><span id="CheckPwds"><span></td></tr>
      <%
	  	If blog_PasswordProtection Then
	  %>
      <tr><td align="right" width="85"><strong><%=lang.MemBer.EditForm(49)%>:</strong></td><td align="left">&nbsp;<input name="Question" type="text" class="userpass" value="" size="40" maxlength="50"/></td></tr>
      <tr><td align="right" width="85"><strong><%=lang.MemBer.EditForm(50)%>:</strong></td><td align="left">&nbsp;<input name="Answer" type="text" class="userpass" value="" size="40" maxlength="50"></td></tr>
      <%
	  	End If
	  %>
	  <tr><td align="right" width="85"><strong>　<%=lang.MemBer.EditForm(13)%>:</strong></td><td align="left" style="padding:3px;"><input name="Gender" type="radio" value="0" checked/> <%=lang.MemBer.EditForm(14)%> <input name="Gender" type="radio" value="1"/><%=lang.MemBer.EditForm(15)%> <input name="Gender" type="radio" value="2"/><%=lang.MemBer.EditForm(16)%></td></tr>
	  <tr><td align="right" width="85"><strong><%=lang.MemBer.EditForm(17)%>:</strong></td><td align="left" style="padding:3px;"><input name="email" type="text" size="38" class="userpass" maxlength="255"/> <input id="hiddenEmail" name="hiddenEmail" type="checkbox" value="1" checked/> <label for="hiddenEmail"><%=lang.MemBer.EditForm(18)%></label></td></tr>
	  <tr><td align="right" width="85"><strong><%=lang.MemBer.EditForm(19)%>:</strong></td><td align="left" style="padding:3px;"><input name="homepage" type="text" size="38" class="userpass" maxlength="255" value=""/></td></tr>
	  <tr><td align="right" width="85"><strong><%=lang.Action.LoginForm(4)%>:</strong></td><td align="left" style="padding:3px;"><input name="validate" type="text" size="4" class="userpass" maxlength="4" onFocus="get_checkcode();this.onfocus=null;" onKeyUp="ajaxcheckcode('isok_checkcode',this);"/> <span id="checkcode"><label style="cursor:pointer;" onClick="get_checkcode();"><%=lang.Action.LoginForm(5)%></label></span> <span id="isok_checkcode"></span></td></tr>

          <tr>
            <td colspan="2" align="center" style="padding:3px;">
              <input name="action" type="hidden" value="save"/>
			  <input name="submit2" type="button" class="userbutton" value="<%=lang.Action.Register%>" onclick="IsPost()"/>
              <input name="button" type="reset" class="userbutton" value="<%=lang.Action.ReSet%>"/></td>
          </tr>
		  </form>
	  </table>
		</div>
	  </div>
	</div>
<br/><br/>
 <%
ElseIf Request.Form("action") = "save" Then
    Dim reg
    reg = register
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
Function register
    Dim ReInfo
    Dim username, password, Confirmpassword, Gender, email, homepage, validate, HideEmail, checkUser, Question, Answer

    ReInfo = Array(lang.Tip.SysTem(1), "", "MessageIcon")
    username = Trim(CheckStr(request.Form("username")))
    password = Trim(CheckStr(request.Form("password")))
    Confirmpassword = Trim(CheckStr(request.Form("Confirmpassword")))
    Gender = CheckStr(request.Form("Gender"))
    email = Trim(CheckStr(request.Form("email")))
    homepage = Trim(checkURL(CheckStr(request.Form("homepage"))))
    validate = CheckStr(request.Form("validate"))
	
	If blog_PasswordProtection Then
		Question = Trim(CheckStr(request.Form("Question")))
		Answer = Trim(CheckStr(request.Form("Answer")))
	End If

    If request.Form("hiddenEmail") = 1 Then
        HideEmail = True
    Else
        HideEmail = False
    End If

    If Len(username) = 0 Then
        ReInfo(0) = lang.Tip.SysTem(1)
        ReInfo(1) = "<b>" & lang.MemBer.EditForm(51) & "</b><br/><a href=""javascript:history.go(-1);"">" & lang.Tip.SysTem(2) & "</a>"
        ReInfo(2) = "WarningIcon"
        register = ReInfo
        Exit Function
    End If

    If Len(username)<2 Or Len(username)>24 Then
        ReInfo(0) = lang.Tip.SysTem(1)
        ReInfo(1) = "<b>" & lang.MemBer.EditForm(52) & "</b><br/><a href=""javascript:history.go(-1);"">" & lang.Tip.SysTem(2) & "</a>"
        ReInfo(2) = "ErrorIcon"
        register = ReInfo
        Exit Function
    End If

    If IsValidUserName(username) = False Then
        ReInfo(0) = lang.Tip.SysTem(1)
        ReInfo(1) = lang.MemBer.EditForm(53) & "<a href=""javascript:history.go(-1);"">" & lang.Tip.SysTem(2) & "</a>"
        ReInfo(2) = "ErrorIcon"
        register = ReInfo
        Exit Function
    End If

    Set checkUser = conn.Execute("select top 1 mem_id from blog_Member where mem_Name='"&username&"'")
    If Not checkUser.EOF Then
        ReInfo(0) = lang.Tip.SysTem(1)
        ReInfo(1) = lang.MemBer.EditForm(54) & "<a href=""javascript:history.go(-1);"">" & lang.Tip.SysTem(2) & "</a>"
        ReInfo(2) = "ErrorIcon"
        register = ReInfo
        Exit Function
    End If

    If Len(password) = 0 Or (Len(password)<6 Or Len(password)>16) Then
        ReInfo(0) = lang.Tip.SysTem(1)
        ReInfo(1) = lang.MemBer.EditForm(55) & "<a href=""javascript:history.go(-1);"">" & lang.Tip.SysTem(2) & "</a>"
        ReInfo(2) = "WarningIcon"
        register = ReInfo
        Exit Function
    End If

    If password<>Confirmpassword Then
        ReInfo(0) = lang.Tip.SysTem(1)
        ReInfo(1) = lang.MemBer.EditForm(56) & "<a href=""javascript:history.go(-1);"">" & lang.Tip.SysTem(2) & "</a>"
        ReInfo(2) = "ErrorIcon"
        register = ReInfo
        Exit Function
    End If
	
	If blog_PasswordProtection Then
		If (len(Question) > 0 and len(Answer) = 0) or (len(Question) = 0 and len(Answer) > 0) Then
			ReInfo(0) = lang.Tip.SysTem(1)
        	ReInfo(1) = lang.MemBer.EditForm(57) & "<a href=""javascript:history.go(-1);"">" & lang.Tip.SysTem(2) & "</a>"
        	ReInfo(2) = "ErrorIcon"
        	register = ReInfo
        	Exit Function
		End If
	End If

    If Len(email)>0 And IsValidEmail(email) = False Then
        ReInfo(0) = lang.Tip.SysTem(1)
        ReInfo(1) = lang.MemBer.EditForm(58) & "<a href=""javascript:history.go(-1);"">" & lang.Tip.SysTem(2) & "</a>"
        ReInfo(2) = "ErrorIcon"
        register = ReInfo
        Exit Function
    End If

    If CStr(LCase(Session("GetCode")))<>CStr(LCase(validate)) Then
        ReInfo(0) = lang.Tip.SysTem(1)
        ReInfo(1) = lang.MemBer.EditForm(59) & "<a href=""javascript:history.go(-1);"">" & lang.Tip.SysTem(2) & "</a>"
        ReInfo(2) = "ErrorIcon"
        register = ReInfo
        Exit Function
    End If

    Dim strSalt, AddUser, hashkey
    hashkey = SHA1(randomStr(6)&Now())
    strSalt = randomStr(6)
    password = SHA1(password&strSalt)
	If blog_PasswordProtection Then
    	AddUser = Array(Array("mem_Name", username), Array("mem_Password", password), Array("mem_Sex", Gender), Array("mem_salt", strSalt), Array("mem_Email", email), Array("mem_HideEmail", Int(HideEmail)), Array("mem_HomePage", homepage), Array("mem_LastIP", getIP), Array("mem_lastVisit", Now()), Array("mem_hashKey", hashkey), Array("mem_Question", Question), Array("mem_Answer", Answer))
	Else
		AddUser = Array(Array("mem_Name", username), Array("mem_Password", password), Array("mem_Sex", Gender), Array("mem_salt", strSalt), Array("mem_Email", email), Array("mem_HideEmail", Int(HideEmail)), Array("mem_HomePage", homepage), Array("mem_LastIP", getIP), Array("mem_lastVisit", Now()), Array("mem_hashKey", hashkey))
	End If
    DBQuest "blog_member", AddUser, "insert"
    'Conn.Execute("INSERT INTO blog_member(mem_Name,mem_Password,mem_Sex,mem_salt,mem_Email,mem_HideEmail,mem_HomePage,mem_LastIP) Values ('"&username&"','"&password&"',"&Gender&",'"&strSalt&"','"&email&"',"&HideEmail&",'"&homepage&"','"&getIP&"')")
    Conn.Execute("UPDATE blog_Info SET blog_MemNums=blog_MemNums+1")
    getInfo(2)
    SQLQueryNums = SQLQueryNums + 2
    ReInfo(0) = lang.MemBer.EditForm(60)
	Referer_Url = Session(CookieName & "_Register_Referer_Url")
    If len(Referer_Url) < 8 Then
    	ReInfo(1) = lang.MemBer.EditForm(61) & "<meta http-equiv=""refresh"" content=""3;url=default.asp""/>"
    Else
    	ReInfo(1) = "<b>" & lang.MemBer.EditForm(62) & "</b><br/><a href="""&Referer_Url&""">" & lang.MemBer.EditForm(63) & "</a>&nbsp;|&nbsp;<a href=""default.asp"">" & lang.Tip.SysTem(4) & "</a><br/>" & lang.MemBer.EditForm(64) & "<meta http-equiv=""Refresh"" content=""3;url="&Referer_Url&"""/>"
    End If
    ReInfo(2) = "MessageIcon"
    register = ReInfo
    Response.Cookies(CookieName)("memName") = username
    Response.Cookies(CookieName)("memHashKey") = hashkey
    Response.Cookies(CookieName).Expires = Date+365
    Session(CookieName&"_LastDo") = "RegisterUser"
End Function

Else
    Referer_Url = Cstr(Request.ServerVariables("HTTP_REFERER"))
    If len(Referer_Url) < 8 then Referer_Url= "http://" & Request.ServerVariables("HTTP_HOST")
    Session(CookieName & "_Register_Referer_Url") = Referer_Url
%><br/><br/>
   <div style="text-align:center;">
  <form name="aform" action="login.asp" method="post">
    <div id="MsgContent">
      <div id="MsgHead"><%=lang.Action.Register%></div>
      <div id="MsgBody">
	  <div style="text-align:left;line-height:120%;"><%=lang.MemBer.EditForm(65)%></div>
	   <input type="button" name="agreesubmit" value="<%=lang.MemBer.EditForm(66)%>" class="userbutton" onclick="location='register.asp?action=agree'"/>
	   <input type="button" name="noagreesubmit" value="<%=lang.MemBer.EditForm(67)%>" class="userbutton" onclick="location='default.asp'"/>
	   </div>
	</div>
  </form>
  </div><br/><br/>
 <script language="javascript">
var secs = 5;
var wait = secs * 1000;
var agreetext="<%=lang.MemBer.EditForm(68)%>";
document.aform.agreesubmit.value = agreetext+"(" + secs + ") ";
document.aform.agreesubmit.disabled = true;
document.aform.noagreesubmit.disabled = false;
for(i = 1; i <= secs; i++) {
  window.setTimeout("update(" + i + ")", i * 1000);
}
window.setTimeout("timer()", wait);
function update(num, value) {
  if(num == (wait/1000)) {
    document.aform.agreesubmit.value = agreetext;
  } else {
    printnr = (wait / 1000)-num;
    document.aform.agreesubmit.value = agreetext+"(" + printnr + ") ";
  }
}
function timer() {
  document.aform.agreesubmit.disabled = false;
  document.aform.noagreesubmit.disabled = false;
  document.aform.agreesubmit.value = "<%=lang.MemBer.EditForm(66)%>";
}
</script>
 
 <%End if%>
 </div>
  <!--#include file="footer.asp" -->