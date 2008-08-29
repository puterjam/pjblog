<!--#include file="commond.asp" -->
<!--#include file="header.asp" -->
<!--内容-->
 <div id="Tbody">
<%
'==================================
'  用户登录页面
'    更新时间: 2006-5-29
'==================================
If Request.QueryString("action") = "logout" Then
    logout(True)

%><br/><br/>
   <div style="text-align:center;">
    <div id="MsgContent" style="width:300px">
      <div id="MsgHead">退出系统</div>
      <div id="MsgBody">
		 <div class="MessageIcon"></div>
        <div class="MessageText"><b>退出登录成功</b><br/>
         <a href="default.asp">单击返回首页</a></div>
         <meta http-equiv="refresh" content="3;url=default.asp"/>
		</div>
	  </div>
	</div>
<br/><br/>
 <%
ElseIf Request.Form("action") = "login" Then
    Dim loginUser
    loginUser = login(Request.Form("UserName"), Request.Form("Password"))
%><br/><br/>
   <div style="text-align:center;">
    <div id="MsgContent" style="width:300px">
      <div id="MsgHead"><%=loginUser(0)%></div>
      <div id="MsgBody">
	   <div class="<%=loginUser(2)%>"></div>
       <div class="MessageText"><%=loginUser(1)%></div>
	  </div>
	</div>
  </div><br/><br/>
<%
Else
%><br/><br/>
   <div style="text-align:center;">
  <form name="checkUser" action="login.asp" method="post">
    <div id="MsgContent">
      <div id="MsgHead">用户登录</div>
      <div id="MsgBody">
	  <input name="action" type="hidden" value="login"/>
	   <label>用户名：<input name="username" type="text" size="18" class="userpass" maxlength="24"/></label><br/>
	   <label>密　码：<input name="password" type="password" size="18" class="userpass"/></label><br/>
	   <label>验证码：<input name="validate" type="text" size="4" class="userpass" maxlength="4" onfocus="this.select()"/> <%=getcode()%></label><br/>
	   　　<label><input name="KeepLogin" type="checkbox" value="1"/>记住我的登录信息</label><br/>
	   <input type="submit" value="登　陆" class="userbutton"/>　<input type="button" value="用户注册" class="userbutton" onclick="location='register.asp'"/>
	   </div>
	</div>
  </form>
  </div><br/><br/>
 <%End if%>
 </div>
  <!--#include file="footer.asp" -->
