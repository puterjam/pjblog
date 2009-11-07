<!--#include file = "../include.asp" -->
<script language="javascript" type="text/javascript" src="../pjblog.common/common.js"></script>
<script language="javascript" type="text/javascript" src="../pjblog.common/Ajax.js"></script>
<script language="javascript" type="text/javascript" src="../pjblog.common/language.js"></script>
<script language="javascript" type="text/javascript" src="../pjblog.common/checkform.js"></script>
<form id="checkUser" action="../pjblog.logic/log_User.asp?action=login" method="post">
    <div id="MsgContent">
      <div id="MsgHead"><%=lang.Set.Asp(17)%></div>
      <div id="MsgBody">
	   <label><input name="username" type="text" size="18" class="userpass" maxlength="24"/> <%=lang.Set.Asp(18)%> </label><br/>
	   <label><input name="password" type="password" size="18" class="userpass" onfocus="this.select()"/> <%=lang.Set.Asp(19)%> </label><br/>
	   <label><input name="validate" type="text" size="4" class="userpass" maxlength="4" id="validate" onfocus="CheckForm.User.GetCode('../pjblog.common/GetCode.asp', 'checkcode')" onkeyup="CheckForm.User.CheckCode('validate', 'isok_checkcode')"/> <span id="checkcode"><label style="cursor:pointer;" onClick="CheckForm.User.GetCode('../pjblog.common/GetCode.asp', 'checkcode')"><%=lang.Set.Asp(4)%></label></span> <span id="isok_checkcode"></span></label><br/>
	   　　<label><input name="KeepLogin" type="checkbox" value="1"/><%=lang.Set.Asp(20)%></label><br/>
	   <input type="submit" value="<%=lang.Set.Asp(21)%>" class="userbutton"/>　<input type="button" value="<%=lang.Set.Asp(22)%>" class="userbutton" onclick="location='register.asp'"/>
	   </div>
	</div>
  </form>
<%
Sys.Close
%>