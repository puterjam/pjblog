<!--#include file="commond.asp" -->
<!--#include file="header.asp" -->
<!--#include file="common/sha1.asp" -->
<!--内容-->
<%
'==================================
'  用户注册页面
'    更新时间: 2006-5-29
'==================================
If blog_Disregister Then showmsg "错误信息", "站点不允许注册新用户<br/><a href=""default.asp"">单击返回</a>", "ErrorIcon", ""
%>

 <div id="Tbody">
<%
If Request.QueryString("action") = "agree" Then
    logout(True)

%><br/><br/>
   <div style="text-align:center;">
    <div id="MsgContent" style="width:520px">
      <div id="MsgHead">用户注册</div>
      <div id="MsgBody">
	  <table width="100%" cellpadding="0" cellspacing="0">
	  <form name="frm" action="register.asp" method="post">
	  <tr><td align="right" width="85"><strong>　昵　称:</strong></td><td align="left" style="padding:3px;"><input name="username" type="text" size="18" class="userpass" maxlength="24" onblur="if (this.value.length != 0) {CheckName();}" onfocus="if (this.value.length != 0) {CheckName();}" onclick="if (this.value.length != 0) {CheckName();}" id="vs"/><input id="PostBack_UserName" value="False|$|False" type="hidden"><font color="#FF0000">&nbsp;*</font> 昵称由2到24个字符组成 <span id="CheckName"></span></td></tr>
	  <tr><td align="right" width="85"><strong>　密　码:</strong></td><td align="left" style="padding:3px;"><input name="password" type="password" size="18" class="userpass" maxlength="16" id="cpassword"/><font color="#FF0000">&nbsp;*</font> 密码必须是6到16个字符，建议使用英文和符号混合</td></tr>
	  <tr><td align="right" width="85"><strong>密码重复:</strong></td><td align="left" style="padding:3px;"><input name="Confirmpassword" type="password" size="18" class="userpass" maxlength="16" onblur="if (this.value.length != 0) {CheckPwd();}" onfocus="if (this.value.length != 0) {CheckPwd();}" onclick="if (this.value.length != 0) {CheckPwd();}" id="cConfirmpassword"/><font color="#FF0000">&nbsp;*</font> 必须和上面的密码一样<span id="CheckPwds"><span></td></tr>
	  <tr><td align="right" width="85"><strong>　性　别:</strong></td><td align="left" style="padding:3px;"><input name="Gender" type="radio" value="0" checked/> 保密 <input name="Gender" type="radio" value="1"/>男 <input name="Gender" type="radio" value="2"/>女</td></tr>
	  <tr><td align="right" width="85"><strong>电子邮件:</strong></td><td align="left" style="padding:3px;"><input name="email" type="text" size="38" class="userpass" maxlength="255"/> <input id="hiddenEmail" name="hiddenEmail" type="checkbox" value="1" checked/> <label for="hiddenEmail">不公开我的电子邮件</label></td></tr>
	  <tr><td align="right" width="85"><strong>个人主页:</strong></td><td align="left" style="padding:3px;"><input name="homepage" type="text" size="38" class="userpass" maxlength="255" value=""/></td></tr>
	  <tr><td align="right" width="85"><strong>验证码:</strong></td><td align="left" style="padding:3px;"><input name="validate" type="text" size="4" class="userpass" maxlength="4"/> <%=getcode()%> <font color="#FF0000">&nbsp;*</font> 请输入验证码</td></tr>

          <tr>
            <td colspan="2" align="center" style="padding:3px;">
              <input name="action" type="hidden" value="save"/>
			  <input name="submit2" type="button" class="userbutton" value="注册新用户" onclick="IsPost()"/>
              <input name="button" type="reset" class="userbutton" value="重写"/></td>
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
    Dim username, password, Confirmpassword, Gender, email, homepage, validate, HideEmail, checkUser

    ReInfo = Array("错误信息", "", "MessageIcon")
    username = Trim(CheckStr(request.Form("username")))
    password = Trim(CheckStr(request.Form("password")))
    Confirmpassword = Trim(CheckStr(request.Form("Confirmpassword")))
    Gender = CheckStr(request.Form("Gender"))
    email = Trim(CheckStr(request.Form("email")))
    homepage = Trim(checkURL(CheckStr(request.Form("homepage"))))
    validate = CheckStr(request.Form("validate"))

    If request.Form("hiddenEmail") = 1 Then
        HideEmail = True
    Else
        HideEmail = False
    End If

    If Len(username) = 0 Then
        ReInfo(0) = "错误信息"
        ReInfo(1) = "<b>请输入用户名(昵称)!</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
        ReInfo(2) = "WarningIcon"
        register = ReInfo
        Exit Function
    End If

    If Len(username)<2 Or Len(username)>24 Then
        ReInfo(0) = "错误信息"
        ReInfo(1) = "<b>用户名(昵称)不能小于2或<br/>大于24个字符！</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
        ReInfo(2) = "ErrorIcon"
        register = ReInfo
        Exit Function
    End If

    If IsValidUserName(username) = False Then
        ReInfo(0) = "错误信息"
        ReInfo(1) = "<b>非法用户名！<br/>请尝试使用其他用户名！</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
        ReInfo(2) = "ErrorIcon"
        register = ReInfo
        Exit Function
    End If

    Set checkUser = conn.Execute("select top 1 mem_id from blog_Member where mem_Name='"&username&"'")
    If Not checkUser.EOF Then
        ReInfo(0) = "错误信息"
        ReInfo(1) = "<b>用户名已经被注册！<br/>请尝试使用其他用户名！</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
        ReInfo(2) = "ErrorIcon"
        register = ReInfo
        Exit Function
    End If

    If Len(password) = 0 Or (Len(password)<6 Or Len(password)>16) Then
        ReInfo(0) = "错误信息"
        ReInfo(1) = "<b>请输入6到16位密码！</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
        ReInfo(2) = "WarningIcon"
        register = ReInfo
        Exit Function
    End If

    If password<>Confirmpassword Then
        ReInfo(0) = "错误信息"
        ReInfo(1) = "<b>两次密码输入不一致！请重新输入。</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
        ReInfo(2) = "ErrorIcon"
        register = ReInfo
        Exit Function
    End If

    If Len(email)>0 And IsValidEmail(email) = False Then
        ReInfo(0) = "错误信息"
        ReInfo(1) = "<b>错误的电子邮件地址。</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
        ReInfo(2) = "ErrorIcon"
        register = ReInfo
        Exit Function
    End If

    If CStr(LCase(Session("GetCode")))<>CStr(LCase(validate)) Then
        ReInfo(0) = "错误信息"
        ReInfo(1) = "<b>验证码有误，请返回重新输入</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
        ReInfo(2) = "ErrorIcon"
        register = ReInfo
        Exit Function
    End If

    Dim strSalt, AddUser, hashkey
    hashkey = SHA1(randomStr(6)&Now())
    strSalt = randomStr(6)
    password = SHA1(password&strSalt)
    AddUser = Array(Array("mem_Name", username), Array("mem_Password", password), Array("mem_Sex", Gender), Array("mem_salt", strSalt), Array("mem_Email", email), Array("mem_HideEmail", Int(HideEmail)), Array("mem_HomePage", homepage), Array("mem_LastIP", getIP), Array("mem_lastVisit", Now()), Array("mem_hashKey", hashkey))
    DBQuest "blog_member", AddUser, "insert"
    'Conn.Execute("INSERT INTO blog_member(mem_Name,mem_Password,mem_Sex,mem_salt,mem_Email,mem_HideEmail,mem_HomePage,mem_LastIP) Values ('"&username&"','"&password&"',"&Gender&",'"&strSalt&"','"&email&"',"&HideEmail&",'"&homepage&"','"&getIP&"')")
    Conn.Execute("UPDATE blog_Info SET blog_MemNums=blog_MemNums+1")
    getInfo(2)
    SQLQueryNums = SQLQueryNums + 2
    ReInfo(0) = "用户注册成功"
    ReInfo(1) = "<b>注册并自动登录成功，三秒钟返回首页！</b><br/><a href=""default.asp"">如果您的浏览器没有自动跳转，请点击这里</a><meta http-equiv=""refresh"" content=""3;url=default.asp""/>"
    ReInfo(2) = "MessageIcon"
    register = ReInfo
    Response.Cookies(CookieName)("memName") = username
    Response.Cookies(CookieName)("memHashKey") = hashkey
    Response.Cookies(CookieName).Expires = Date+365
    Session(CookieName&"_LastDo") = "RegisterUser"
End Function

Else
%><br/><br/>
   <div style="text-align:center;">
  <form name="aform" action="login.asp" method="post">
    <div id="MsgContent">
      <div id="MsgHead">用户注册</div>
      <div id="MsgBody">
	  <div style="text-align:left;line-height:120%;">为维护网上公共秩序和社会稳定，请您自觉遵守以下条款： <br/><br/>

　 一、不得利用本站危害国家安全、泄露国家秘密，不得侵犯国家社会集体的和公民的合法权益，不得利用本站制作、复制和传播下列信息：<br/> 
　　 （一）煽动抗拒、破坏宪法和法律、行政法规实施的； <br/>
　　 （二）煽动颠覆国家政权，推翻社会主义制度的； <br/>
　　 （三）煽动分裂国家、破坏国家统一的； <br/>
　　 （四）煽动民族仇恨、民族歧视，破坏民族团结的； <br/>
　　 （五）捏造或者歪曲事实，散布谣言，扰乱社会秩序的； <br/>
　　 （六）宣扬封建迷信、淫秽、色情、赌博、暴力、凶杀、恐怖、教唆犯罪的； <br/>
　　 （七）公然侮辱他人或者捏造事实诽谤他人的，或者进行其他恶意攻击的；<br/> 
　　 （八）损害国家机关信誉的； <br/>
　　 （九）其他违反宪法和法律行政法规的； <br/>
　　 （十）进行商业广告行为的。 <br/>
　 二、互相尊重，对自己的言论和行为负责。 <br/><br/></div>
	   <input type="button" name="agreesubmit" value="我同意" class="userbutton" onclick="location='register.asp?action=agree'"/>
	   <input type="button" name="noagreesubmit" value="不同意" class="userbutton" onclick="location='default.asp'"/>
	   </div>
	</div>
  </form>
  </div><br/><br/>
 <script language="javascript">
var secs = 3;
var wait = secs * 1000;
var agreetext="请仔细阅读以上条款";
document.aform.agreesubmit.value = agreetext+"(" + secs + ") ";
document.aform.agreesubmit.disabled = true;
document.aform.noagreesubmit.disabled = true;
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
  document.aform.agreesubmit.value = "我同意";
}
</script>
 
 <%End if%>
 </div>
  <!--#include file="footer.asp" -->
