<!--#include file="commond.asp" -->
<!--#include file="header.asp" -->
<!--内容-->
 <div id="Tbody">
<%
'==================================
'  用户登录页面
'    更新时间: 2006-5-29
'==================================
dim Referer_Url
If Request.QueryString("action") = "logout" Then
    logout(True)
    Referer_Url = Cstr(Request.ServerVariables("HTTP_REFERER"))
    If len(Referer_Url) < 8 Then Referer_Url= "http://" & Request.ServerVariables("HTTP_HOST")
%><br/><br/>
   <div style="text-align:center;">
    <div id="MsgContent" style="width:300px">
      <div id="MsgHead"><%=lang.Action.Logout%></div>
      <div id="MsgBody">
		 <div class="MessageIcon"></div>
        <div class="MessageText"><b><%=lang.Action.LoginSuc(1)%></b><br/>
         <a href="default.asp"><%=lang.Tip.SysTem(4)%></a>&nbsp;|&nbsp;<a href="<%=Referer_Url%>"><%=lang.Tip.SysTem(8)%></a>
         <br/><%=lang.Action.LoginSuc(2)%></div>
         <meta http-equiv="refresh" content="3;url=<%=Referer_Url%>"/>
		</div>
	  </div>
	</div>
<br/><br/>
 <%
ElseIf Request.Form("action") = "login" Then
    Dim loginUser
    Referer_Url = Session(CookieName & "_Login_Referer_Url")
    If len(Referer_Url) < 8 Then Referer_Url = Cstr(Request.ServerVariables("HTTP_REFERER"))
    If len(Referer_Url) < 8 Then Referer_Url = "http://" & Request.ServerVariables("HTTP_HOST")
    loginUser = login(Request.Form("UserName"), Request.Form("Password"))
%><br/><br/>
   <div style="text-align:center;">
    <div id="MsgContent" style="width:300px">
      <div id="MsgHead"><%=loginUser(0)%></div>
      <div id="MsgBody">
	   <div class="<%=loginUser(2)%>"></div>
       <div class="MessageText"><%=Replace(Replace(loginUser(1),"default.asp",Referer_Url), lang.Tip.SysTem(4) & "</a>", lang.Tip.SysTem(9) & "</a>&nbsp;|&nbsp;<a href=""default.asp"">" & lang.Tip.SysTem(4) & "</a><br/>" & lang.Action.LoginSuc(2))%></div>
	  </div>
	</div>
  </div><br/><br/>
<%
Else
	Referer_Url = Cstr(Request.ServerVariables("HTTP_REFERER"))
    If len(Referer_Url) < 8 then Referer_Url= "http://" & Request.ServerVariables("HTTP_HOST")
    Session(CookieName & "_Login_Referer_Url") = Referer_Url
%><br/><br/>
   <div style="text-align:center;">
  <form name="checkUser" action="login.asp" method="post">
    <div id="MsgContent">
      <div id="MsgHead"><%=lang.Action.LoginForm(1)%></div>
      <div id="MsgBody">
	  <input name="action" type="hidden" value="login"/>
	   <label><%=lang.Action.LoginForm(2)%>：<input name="username" type="text" size="18" class="userpass" maxlength="24"/></label><br/>
	   <label><%=lang.Action.LoginForm(3)%>：<input name="password" type="password" size="18" class="userpass" onfocus="this.select()"/></label><br/>
	   <label><%=lang.Action.LoginForm(4)%>：<input name="validate" type="text" size="4" class="userpass" maxlength="4" onFocus="get_checkcode();this.onfocus=null;" onKeyUp="ajaxcheckcode('isok_checkcode',this);"/> <span id="checkcode"><label style="cursor:pointer;" onClick="get_checkcode();"><%=lang.Action.LoginForm(5)%></label></span> <span id="isok_checkcode"></span></label><br/>
	   　　<label><input name="KeepLogin" type="checkbox" value="1"/><%=lang.Action.LoginForm(6)%></label><br/>
	   <input type="submit" value="<%=lang.Action.Login%>" class="userbutton"/>　<input type="button" value="<%=lang.Action.Register%>" class="userbutton" onclick="location='register.asp'"/>
	   </div>
	</div>
  </form>
  </div><br/><br/>
 <%End if%>
 </div>
<!--#include file="footer.asp" -->