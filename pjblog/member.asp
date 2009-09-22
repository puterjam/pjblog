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
Dim blog_Mem, Referer_Url
If Request.QueryString("action") = "edit" Then
    If memName = Empty Then RedirectUrl("member.asp")
    Referer_Url = Cstr(Request.ServerVariables("HTTP_REFERER"))
    If len(Referer_Url) < 8 then Referer_Url= "http://" & Request.ServerVariables("HTTP_HOST")
    Session(CookieName & "_Member_Referer_Url") = Referer_Url

%><br/><br/>
   <div style="text-align:center;">
    <div id="MsgContent" style="width:520px">
      <div id="MsgHead"><%=lang.Tip.MemBer.EditForm(1)%></div>
      <div id="MsgBody">
       <%Set blog_Mem = conn.Execute("select * from blog_Member where mem_Name='"&CheckStr(memName)&"'")
If blog_Mem.EOF Or blog_Mem.bof Then

%>
	     <div class="ErrorIcon"></div>
         <div class="MessageText"  align="center"><%=lang.Tip.MemBer.EditForm(2)%><br/><a href="javascript:history.go(-1)"><%=lang.Tip.SysTem(2)%></a></div>
         <script>
           document.getElementById("MsgContent").style.width="300px"
           document.getElementById("MsgHead").innerText = "<%=lang.Tip.SysTem(1)%>"
         </script>
       <%else%>
	  <table width="100%" cellpadding="0" cellspacing="0">
	  <form name="frm" action="member.asp" method="post" onsubmit="if (this.Oldpassword.value.length<1){alert('<%=lang.Tip.MemBer.EditForm(3)%>');this.Oldpassword.focus();return false}">
	  <input name="UID" type="hidden" value="<%=blog_Mem("mem_ID")%>"/>
	  <tr><td align="right" width="85"><strong> <%=lang.Tip.MemBer.EditForm(4)%>:</strong></td><td align="left" style="padding:3px;"><%=blog_Mem("mem_Name")%></td></tr>
	  <tr><td align="right" width="85"><strong>　<%=lang.Tip.MemBer.EditForm(5)%>:</strong></td><td align="left" style="padding:3px;"><input name="Oldpassword" type="password" size="18" class="userpass" maxlength="16"/><font color="#FF0000">&nbsp;*</font> <%=lang.Tip.MemBer.EditForm(6)%></td></tr>
	  <tr><td align="right" width="85"><strong>　<%=lang.Action.PassWord%>:</strong></td><td align="left" style="padding:3px;"><input name="password" type="password" size="18" class="userpass" maxlength="16" id="cpassword" onkeyup="istrong()"/> <%=lang.Tip.MemBer.EditForm(7)%></td></tr>
      <tr><td align="right" width="85"><strong><%=lang.Tip.MemBer.EditForm(8)%>:</strong></td><td align="left"> &nbsp;<img src="images/0.gif" id="strong"></td></tr>
	  <tr><td align="right" width="85"><strong><%=lang.Tip.MemBer.EditForm(9)%>:</strong></td><td align="left" style="padding:3px;"><input name="Confirmpassword" type="password" size="18" class="userpass" maxlength="16"/> <%=lang.Tip.MemBer.EditForm(10)%></td></tr>
      <%
	  	If blog_PasswordProtection Then
	  %>
      <tr><td align="right" width="85"><strong><%=lang.Tip.MemBer.EditForm(11)%>:</strong></td><td align="left">&nbsp;<b><a href="javascript:;" onclick="ModiyPassProtect('<%=blog_Mem("mem_Name")%>', 350, '<%=blog_Mem("mem_Question")%>', <%=blog_Mem("mem_ID")%>)"><%=lang.Tip.MemBer.EditForm(12)%></a></b></td></tr>
      <%
	  	End If
	  %>
	  <tr><td align="right" width="85"><strong>　<%=lang.Tip.MemBer.EditForm(13)%>:</strong></td><td align="left" style="padding:3px;"><input name="Gender" type="radio" value="0" <%if blog_Mem("mem_Sex")=0 then response.write "checked"%>/> <%=lang.Tip.MemBer.EditForm(14)%> <input name="Gender" type="radio" value="1" <%if blog_Mem("mem_Sex")=1 then response.write "checked"%>/><%=lang.Tip.MemBer.EditForm(15)%> <input name="Gender" type="radio" value="2" <%if blog_Mem("mem_Sex")=2 then response.write "checked"%>/><%=lang.Tip.MemBer.EditForm(16)%></td></tr>
	  <tr><td align="right" width="85"><strong><%=lang.Tip.MemBer.EditForm(17)%>:</strong></td><td align="left" style="padding:3px;"><input name="email" type="text" size="38" class="userpass" maxlength="255" value="<%=blog_Mem("mem_Email")%>"/> <input id="hiddenEmail" name="hiddenEmail" type="checkbox" value="1" <%if blog_Mem("mem_HideEmail") then response.write "checked"%>/> <label for="hiddenEmail"><%=lang.Tip.MemBer.EditForm(18)%></label></td></tr>
	  <tr><td align="right" width="85"><strong><%=lang.Tip.MemBer.EditForm(19)%>:</strong></td><td align="left" style="padding:3px;"><input name="homepage" type="text" size="38" class="userpass" maxlength="255" value="<%=blog_Mem("mem_HomePage")%>"/></td></tr>
	  <tr><td align="right" width="85"><strong><%=lang.Tip.MemBer.EditForm(20)%>:</strong></td><td align="left" style="padding:3px;"><input name="QQ" type="text" size="10" class="userpass" value="<%=blog_Mem("mem_QQ")%>"/></td></tr>

          <tr>
            <td colspan="2" align="center" style="padding:3px;">
              <input name="action" type="hidden" value="save"/>
			  <input name="submit2" type="submit" class="userbutton" value="<%=lang.Tip.MemBer.EditForm(21)%>"/>
              <input name="button" type="reset" class="userbutton" value="<%=lang.Action.ReSet%>"/></td>
          </tr>
		  </form>
	<%
End If
blog_Mem.Close
Set blog_Mem = Nothing
%>
	  </table>
		</div>
	  </div>
	</div>
<br/><br/>
 <%
ElseIf Request.QueryString("action") = "view" Then
%>
<br/><br/>
   <div style="text-align:center;">
    <div id="MsgContent" style="width:420px">
      <div id="MsgHead"><%=lang.Tip.MemBer.EditForm(22)%></div>
      <div id="MsgBody">
	  <table width="100%" cellpadding="0" cellspacing="0">
	  <%
If CheckStr(Request.QueryString("memName")) = Empty Then

%>
	     <div class="ErrorIcon"></div>
         <div class="MessageText"  align="center"><%=lang.Err.info(999)%><%=lang.Tip.MemBer.EditForm(23)%><br/><a href="javascript:history.go(-1)"><%=lang.Tip.SysTem(2)%></a></div>
         <script>
           document.getElementById("MsgContent").style.width="300px"
           document.getElementById("MsgHead").innerText = "<%=lang.Tip.SysTem(1)%>";
         </script>
         <%
Else
    Set blog_Mem = conn.Execute("select * from blog_Member where mem_Name='"&CheckStr(Request.QueryString("memName"))&"'")
    If blog_Mem.EOF Or blog_Mem.bof Then

%>
	     <div class="ErrorIcon"></div>
         <div class="MessageText"  align="center"><%=lang.Tip.MemBer.EditForm(2)%><br/><a href="javascript:history.go(-1)"><%=lang.Tip.SysTem(2)%></a></div>
         <script>
           document.getElementById("MsgContent").style.width="300px"
           document.getElementById("MsgHead").innerText = "<%=lang.Tip.SysTem(1)%>";
         </script>
       <%else%>
	  <tr><td align="right" width="85"><strong>　<%=lang.Tip.MemBer.EditForm(4)%>:</strong></td><td align="left" style="padding:3px;"><%=blog_Mem("mem_Name")%></td></tr>
	  <tr><td align="right" width="85"><strong>　<%=lang.Tip.MemBer.EditForm(13)%>:</strong></td><td align="left" style="padding:3px;"><%
Select Case Int(blog_Mem("mem_Sex"))
    Case 1
        response.Write lang.Tip.MemBer.EditForm(24)
    Case 2
        response.Write lang.Tip.MemBer.EditForm(25)
    Case Else
        response.Write lang.Tip.MemBer.EditForm(14)
End Select

%></td></tr>
	  <tr><td align="right" width="85"><strong><%=lang.Tip.MemBer.EditForm(17)%>:</strong></td><td align="left" style="padding:3px;"><%if (blog_Mem("mem_HideEmail") and (not stat_Admin)) or len(blog_Mem("mem_Email"))<1 or isnull(blog_Mem("mem_Email")) then response.write lang.Tip.MemBer.EditForm(26) else response.write blog_Mem("mem_Email") end if%></td></tr>
	  <tr><td align="right" width="85"><strong><%=lang.Tip.MemBer.EditForm(19)%>:</strong></td><td align="left" style="padding:3px;"><a href="<%=blog_Mem("mem_HomePage")%>" target="_blank"><%=blog_Mem("mem_HomePage")%></a></td></tr>
	  <tr><td align="right" width="85"><strong>　<%=lang.Tip.MemBer.EditForm(20)%>:</strong></td><td align="left" style="padding:3px;"><%=blog_Mem("mem_QQ")%></td></tr>
	  <tr><td align="right" width="85"><strong><%=lang.Tip.MemBer.EditForm(27)%>:</strong></td><td align="left" style="padding:3px;"><%=lang.Tip.MemBer.EditForm(28)(blog_Mem("mem_PostLogs"), blog_Mem("mem_PostComms"), blog_Mem("mem_PostMessageNums"))%></td></tr>
          <tr>
            <td colspan="2" align="center" style="padding:3px;">
			  <input type="button" class="userbutton" value="<%=lang.Tip.SysTem(6)%>" onclick="history.go(-1)"/>
          </tr>
        <%
End If
blog_Mem.Close
Set blog_Mem = Nothing
End If

%>
	  </table>
		</div>
	  </div>
	</div>
<br/><br/>
<%
ElseIf Request.Form("action") = "save" Then
    Dim reg
    Referer_Url = Session(CookieName & "_Member_Referer_Url")
    If len(Referer_Url) < 8 then Referer_Url = Cstr(Request.ServerVariables("HTTP_REFERER"))
    If len(Referer_Url) < 8 then Referer_Url = "http://" & Request.ServerVariables("HTTP_HOST")
    reg = SaveMem
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
    Dim PageCount, BM
    Set blog_Mem = Server.CreateObject("ADODB.RecordSet")
    SQL = "SELECT * FROM blog_Member order by mem_RegTime desc"
    blog_Mem.Open SQL, Conn, 1, 1
    SQLQueryNums = SQLQueryNums + 1
    blog_Mem.PageSize = 20
    blog_Mem.AbsolutePage = CurPage
%><br/><br/>
   <div style="text-align:center;">
    <div id="MsgContent" style="width:480px">
      <div id="MsgHead"><%=lang.Tip.MemBer.EditForm(29)%></div>
      <div id="MsgBody" style="padding:0px">
<%If blog_Mem.EOF Or blog_Mem.bof Then
    response.Write lang.Tip.MemBer.EditForm(30)
Else
%>
	  <table width="100%" border="0" cellspacing="0"><tr><th width="80"><%=lang.Action.LoginForm(2)%></th><th><%=lang.Tip.MemBer.EditForm(17)%></th><th><%=lang.Tip.MemBer.EditForm(19)%></th><th><%=lang.Tip.MemBer.EditForm(20)%></th><th width="34"><%=lang.Tip.MemBer.EditForm(31)%></th><th width="34"><%=lang.Tip.MemBer.EditForm(32)%></th><th width="34"><%=lang.Tip.MemBer.EditForm(33)%></th><th><%=lang.Tip.MemBer.EditForm(34)%></th></tr>
 <%
Do Until blog_Mem.EOF Or PageCount = blog_Mem.PageSize
    If blog_Mem("mem_HideEmail") Or Len(blog_Mem("mem_Email"))<1 Or IsNull(blog_Mem("mem_Email")) Then
        BM = "<td><img border=""0"" src=""images/noemail.gif"" alt=""" & lang.Tip.MemBer.EditForm(35) & """/></td>"
    Else
        BM = "<td><img border=""0"" src=""images/email.gif"" alt="""&blog_Mem("mem_Email")&"""/></td>"
    End If
    If Len(blog_Mem("mem_HomePage"))<1 Or IsNull(blog_Mem("mem_HomePage")) Then
        BM = BM&"<td><img border=""0"" src=""images/nourl.gif"" alt=""" & lang.Tip.MemBer.EditForm(36) & """/></td>"
    Else
        BM = BM&"<td><a href="""&blog_Mem("mem_HomePage")&""" target=""_blank""><img border=""0"" src=""images/url.gif"" alt="""&blog_Mem("mem_HomePage")&"""/></a></td>"
    End If
    If Len(blog_Mem("mem_QQ"))<1 Or IsNull(blog_Mem("mem_QQ")) Then
        BM = BM&"<td><img border=""0"" src=""images/nooicq.gif"" alt=""" & lang.Tip.MemBer.EditForm(37) & """/></td>"
    Else
        BM = BM&"<td><a href=""http://wpa.qq.com/msgrd?V=1&Uin="&blog_Mem("mem_QQ")&"&Site="&SiteName&""" target=""_blank""><img border=""0"" src=""images/oicq.gif"" alt=""" & lang.Tip.MemBer.EditForm(38) & """/></a></td>"
    End If
    response.Write "<tr><td align=""left""><a href=""member.asp?action=view&memName="&Server.URLEncode(blog_Mem("mem_Name"))&""">"&blog_Mem("mem_Name")&"</a></td>"&BM&"<td>"&blog_Mem("mem_PostLogs")&"</td><td>"&blog_Mem("mem_PostComms")&"</td><td>"&blog_Mem("mem_PostMessageNums")&"</td><td>"&DateToStr(blog_Mem("mem_RegTime"), "Y-m-d H:I A")&"</td>"
    PageCount = PageCount + 1
    blog_Mem.movenext
Loop
response.Write "</table>"
response.Write "<div class=""pageContent"">"&MultiPage(blog_Mem.RecordCount, 20, CurPage, "?", "", "float:left","")&"</div>"
End If
blog_Mem.Close
Set blog_Mem = Nothing

%>
	   </div>
	</div>
  </div><br/><br/>
 
 <%End if%>
 </div>
  <!--#include file="footer.asp" -->
  <%
Function SaveMem
    Dim ReInfo
    Dim UID, username, Oldpassword, password, Confirmpassword, Gender, email, homepage, QQ, HideEmail, checkUser
    UID = CLng(Trim(CheckStr(request.Form("UID"))))
    ReInfo = Array(lang.Tip.SysTem(1), "", "MessageIcon")
    Oldpassword = Trim(CheckStr(request.Form("Oldpassword")))
    password = Trim(CheckStr(request.Form("password")))
    Confirmpassword = Trim(CheckStr(request.Form("Confirmpassword")))
    Gender = CheckStr(request.Form("Gender"))
    email = Trim(CheckStr(request.Form("email")))
    homepage = Trim(checkURL(CheckStr(request.Form("homepage"))))
    QQ = CheckStr(request.Form("QQ"))
    If request.Form("hiddenEmail") = 1 Then
        HideEmail = True
    Else
        HideEmail = False
    End If

    If IsInteger(Gender) = False Then
        ReInfo(0) = lang.Tip.SysTem(1)
        ReInfo(1) = "<b>" & lang.Err.info(999) & "</b><br/><a href=""javascript:history.go(-1);"">" & lang.Tip.SysTem(2) & "</a>"
        ReInfo(2) = "ErrorIcon"
        SaveMem = ReInfo
        Exit Function
    End If

    Set checkUser = conn.Execute("select top 1 * from blog_Member where mem_id="&UID&" and mem_Name='"&CheckStr(memName)&"'")
    If checkUser.EOF Then
        ReInfo(0) = lang.Tip.SysTem(1)
        ReInfo(1) = lang.Tip.MemBer.EditForm(39) & "<br/><a href=""javascript:history.go(-1);"">" & lang.Tip.SysTem(2) & "</a>"
        ReInfo(2) = "ErrorIcon"
        SaveMem = ReInfo
        Exit Function
    End If
    If Len(password)>0 Then
        If Len(password)<6 Or Len(password)>16 Then
            ReInfo(0) = lang.Tip.SysTem(1)
            ReInfo(1) = lang.Tip.MemBer.EditForm(40) & "<br/><a href=""javascript:history.go(-1);"">" & lang.Tip.SysTem(2) & "</a>"
            ReInfo(2) = "WarningIcon"
            SaveMem = ReInfo
            Exit Function
        End If
        If password<>Confirmpassword Then
            ReInfo(0) = lang.Tip.SysTem(1)
            ReInfo(1) = lang.Tip.MemBer.EditForm(41) & "<br/><a href=""javascript:history.go(-1);"">" & lang.Tip.SysTem(2) & "</a>"
            ReInfo(2) = "ErrorIcon"
            SaveMem = ReInfo
            Exit Function
        End If
    End If

    If Len(QQ)>0 And IsInteger(QQ) = False Then
        ReInfo(0) = lang.Tip.SysTem(1)
        ReInfo(1) = lang.Tip.MemBer.EditForm(42) & "<br/><a href=""javascript:history.go(-1);"">" & lang.Tip.SysTem(2) & "</a>"
        ReInfo(2) = "ErrorIcon"
        SaveMem = ReInfo
        Exit Function
    End If

    If Len(email)>0 And IsValidEmail(email) = False Then
        ReInfo(0) = lang.Tip.SysTem(1)
        ReInfo(1) = lang.Tip.MemBer.EditForm(43) & "<br/><a href=""javascript:history.go(-1);"">" & lang.Tip.SysTem(2) & "</a>"
        ReInfo(2) = "ErrorIcon"
        SaveMem = ReInfo
        Exit Function
    End If

    Set checkUser = conn.Execute("select top 1 * from blog_Member where mem_id="&UID&" and mem_Name='"&CheckStr(memName)&"'")
    If checkUser("mem_Password")<>SHA1(Oldpassword&checkUser("mem_salt")) Then
        ReInfo(0) = lang.Tip.SysTem(1)
        ReInfo(1) = lang.Tip.MemBer.EditForm(44) & "<br/><a href=""javascript:history.go(-1);"">" & lang.Tip.MemBer.EditForm(45) & "</a>"
        ReInfo(2) = "ErrorIcon"
        SaveMem = ReInfo
        Exit Function
    End If

    Conn.Execute("update blog_member set mem_Sex="&Gender&",mem_Email='"&email&"',mem_HideEmail="&HideEmail&",mem_HomePage='"&homepage&"',mem_QQ='"&QQ&"' where mem_id="&UID&" and mem_Name='"&CheckStr(memName)&"'")
    SQLQueryNums = SQLQueryNums + 1
    If Len(password)>0 Then
        Dim strSalt
        strSalt = randomStr(6)
        password = SHA1(password&strSalt)
        Conn.Execute("update blog_member set mem_Password='"&password&"',mem_salt='"&strSalt&"' where mem_id="&UID&" and mem_Name='"&CheckStr(memName)&"'")
        SQLQueryNums = SQLQueryNums + 1
        logout(True)
        ReInfo(0) = lang.Tip.MemBer.EditForm(46)
        ReInfo(1) = lang.Tip.MemBer.EditForm(47) & "<meta http-equiv=""refresh"" content=""3;url=login.asp""/>"
        ReInfo(2) = "MessageIcon"
        SaveMem = ReInfo
        Session(CookieName&"_LastDo") = "EditUser"
        Exit Function
    End If
    getInfo(2)
    ReInfo(0) = lang.Tip.MemBer.EditForm(46)
    ReInfo(1) = "<b>您的资料已经修改成功</b><br/><a href=""default.asp"">返回首页</a>&nbsp;|&nbsp;<a href="""&Referer_Url&""">单击返回修改前页面</a><br/>三秒后自动返回修改前页面<meta http-equiv=""refresh"" content=""3;url="&Referer_Url&"""/>"
    ReInfo(2) = "MessageIcon"
    SaveMem = ReInfo
    Session(CookieName&"_LastDo") = "EditUser"
End Function
%>
