<!--#include file="md5.asp" -->
<!--#include file="sha1.asp" -->
<%
'===============================================================
'  Check User For PJblog2
'    更新时间: 2006-5-29
'===============================================================
Dim UserID,memName,memStatus
memStatus="Guest"
function login(UserName,Password)
 Dim validate,ReInfo,HashKey
 UserName=CheckStr(UserName)
 Password=CheckStr(Password)
 
 validate=trim(request.form("validate"))
 ReInfo=Array("错误信息","","MessageIcon",false)
 IF trim(UserName)="" OR trim(Password)="" Then
	 ReInfo(0)="错误信息"
	 ReInfo(1)="<b>请将信息输入完整</b><br/><a href=""javascript:history.go(-1);"">请返回重新输入</a><br/>"
	 ReInfo(2)="WarningIcon"
	 login=ReInfo
	 logout(false)
	 exit function
  end if
  
  IF validate="" Then
	  ReInfo(0)="错误信息"
	  ReInfo(1)="<b>请输入登录验证码</b><br/><a href=""javascript:history.go(-1);"">请返回重新输入</a><br/>"
	  ReInfo(2)="WarningIcon"
	  login=ReInfo
	  logout(false)
      exit function
  end if
  
 if IsValidUserName(UserName)=false then
	 ReInfo(0)="错误信息"
	 ReInfo(1)="<b>非法用户名！<br/>请尝试使用其他用户名！</b><br/><a href=""javascript:history.go(-1);"">单击返回</a><br/>"
	 ReInfo(2)="ErrorIcon"
	 login=ReInfo
	 logout(false)
	 exit function
 end if
 
  IF cstr(lcase(Session("GetCode")))<>cstr(lcase(validate)) then
	  ReInfo(0)="错误信息"
	  ReInfo(1)="<b>验证码有误，请返回重新输入</b><br/><a href=""javascript:history.go(-1);"">请返回重新输入</a><br/>"
	  ReInfo(2)="ErrorIcon"
	  login=ReInfo
	  logout(false)
      exit function
  end if
         HashKey=SHA1(randomStr(6)&now())
	     Dim memLogin
		 Set memLogin=Server.CreateObject("ADODB.Recordset")
		 SQL="SELECT Top 1 mem_Name,mem_Password,mem_salt,mem_Status,mem_LastIP,mem_lastVisit,mem_hashKey FROM blog_member WHERE mem_Name='"&UserName&"' AND mem_salt<>''"
         memLogin.Open SQL,conn,1,3
         SQLQueryNums=SQLQueryNums+1
         IF memLogin.EOF And memLogin.BOF Then
           memLogin.Close
		 SQL="SELECT Top 1 mem_Name,mem_Password,mem_salt,mem_Status,mem_LastIP,mem_lastVisit,mem_hashKey FROM blog_member WHERE mem_Name='"&UserName&"' AND mem_Password='"&md5(Password)&"'"
           memLogin.Open SQL,conn,1,3
           SQLQueryNums=SQLQueryNums+1
		   IF memLogin.EOF AND memLogin.BOF Then
		    	ReInfo(0)="错误信息"
		     	ReInfo(1)="<b>用户名与密码错误</b><br/><a href=""javascript:history.go(-1);"">请返回重新输入</a><br/>"
		    	ReInfo(2)="ErrorIcon"
		    	logout(false)
	       Else
	           '进行MD5密码验证，转换旧帐户密码验证方式
	            dim strSalt
	            strSalt=randomStr(6)
	            memLogin("mem_salt")=strSalt
            	memLogin("mem_LastIP")=getIP()
            	memLogin("mem_lastVisit")=now()
            	memLogin("mem_hashKey")=HashKey
            	memLogin("mem_Password")=SHA1(Password&strSalt)
		    	Response.Cookies(CookieName)("memName")=memLogin("mem_Name")
		    	Response.Cookies(CookieName)("memHashKey")=HashKey
		    	if Request.Form("KeepLogin")="1" then Response.Cookies(CookieName).Expires=Date+365
		    	memLogin.Update
		    	ReInfo(0)="登录成功"
		    	ReInfo(1)="<b>"&memLogin("mem_Name")&"</b>，欢迎你的再次光临。<br/><a href=""default.asp"">点击返回主页</a>"
		    	ReInfo(2)="MessageIcon"
	            ReInfo(3)=true
		   End IF
		 else
		   if memLogin("mem_Password")<>SHA1(Password&memLogin("mem_salt")) then
		    	ReInfo(0)="错误信息"
		     	ReInfo(1)="<b>用户名与密码错误</b><br/><a href=""javascript:history.go(-1);"">请返回重新输入</a><br/>"
		    	ReInfo(2)="ErrorIcon"
		    	logout(false)
		   else
            	memLogin("mem_LastIP")=getIP()
            	memLogin("mem_lastVisit")=now()
            	memLogin("mem_hashKey")=HashKey
		    	Response.Cookies(CookieName)("memName")=memLogin("mem_Name")
		    	Response.Cookies(CookieName)("memHashKey")=HashKey
		    	if Request.Form("KeepLogin")="1" then Response.Cookies(CookieName).Expires=Date+365
		    	memLogin.Update
		    	ReInfo(0)="登录成功"
		    	ReInfo(1)="<b>"&memLogin("mem_Name")&"</b>，欢迎你的再次光临。<br/><a href=""default.asp"">点击返回主页</a><meta http-equiv=""refresh"" content=""3;url=default.asp""/>"
		    	ReInfo(2)="MessageIcon"
	            ReInfo(3)=true
		   end if
		 end if
		 memLogin.Close
		 Set memLogin=Nothing
  login=ReInfo
end function

function login2(UserName,Password)
 Dim validate,ReInfo,HashKey
 UserName=CheckStr(UserName)
 Password=CheckStr(Password)
 
 ReInfo=Array("错误信息","","MessageIcon",false)
 IF trim(UserName)="" OR trim(Password)="" Then
	 ReInfo(0)="错误信息"
	 ReInfo(1)="<b>请将信息输入完整</b><br/><a href=""javascript:history.go(-1);"">请返回重新输入</a><br/>"
	 ReInfo(2)="WarningIcon"
	 login2=ReInfo
	 logout(false)
	 UserRight(1)
	 exit function
  end if
 
 if IsValidUserName(UserName)=false then
	 ReInfo(0)="错误信息"
	 ReInfo(1)="<b>非法用户名！<br/>请尝试使用其他用户名！</b><br/><a href=""javascript:history.go(-1);"">单击返回</a><br/>"
	 ReInfo(2)="ErrorIcon"
	 login2=ReInfo
	 logout(false)
	 UserRight(1)
	 exit function
 end if
 
         HashKey=SHA1(randomStr(6)&now())
	     Dim memLogin
		 Set memLogin=Server.CreateObject("ADODB.Recordset")
		 SQL="SELECT Top 1 mem_Name,mem_Password,mem_salt,mem_Status,mem_LastIP,mem_lastVisit,mem_hashKey FROM blog_member WHERE mem_Name='"&UserName&"' AND mem_salt<>''"
         memLogin.Open SQL,conn,1,3
         SQLQueryNums=SQLQueryNums+1
         IF memLogin.EOF And memLogin.BOF Then
           memLogin.Close
		 SQL="SELECT Top 1 mem_Name,mem_Password,mem_salt,mem_Status,mem_LastIP,mem_lastVisit,mem_hashKey FROM blog_member WHERE mem_Name='"&UserName&"' AND mem_Password='"&md5(Password)&"'"
           memLogin.Open SQL,conn,1,3
           SQLQueryNums=SQLQueryNums+1
		   IF memLogin.EOF AND memLogin.BOF Then
		    	ReInfo(0)="错误信息"
		     	ReInfo(1)="<b>用户名与密码错误</b><br/><a href=""javascript:history.go(-1);"">请返回重新输入</a><br/>"
		    	ReInfo(2)="ErrorIcon"
		    	logout(false)
	       Else
	           '进行MD5密码验证，转换旧帐户密码验证方式
	            dim strSalt
	            strSalt=randomStr(6)
	            memLogin("mem_salt")=strSalt
            	memLogin("mem_LastIP")=getIP()
            	memLogin("mem_lastVisit")=now()
            	memLogin("mem_hashKey")=HashKey
            	memLogin("mem_Password")=SHA1(Password&strSalt)
		    	memLogin.Update
		    	memName=memLogin("mem_Name")
		    	memStatus=memLogin("mem_Status")
		    	ReInfo(0)="登录成功"
		    	ReInfo(1)="<b>"&memLogin("mem_Name")&"</b>，欢迎你的再次光临。<br/><a href=""default.asp"">点击返回主页</a>"
		    	ReInfo(2)="MessageIcon"
	            ReInfo(3)=true
		   End IF
		 else
		   if memLogin("mem_Password")<>SHA1(Password&memLogin("mem_salt")) then
		    	ReInfo(0)="错误信息"
		     	ReInfo(1)="<b>用户名与密码错误</b><br/><a href=""javascript:history.go(-1);"">请返回重新输入</a><br/>"
		    	ReInfo(2)="ErrorIcon"
		    	logout(false)
		   else
		    	memName=memLogin("mem_Name")
		    	memStatus=memLogin("mem_Status")
		    	ReInfo(0)="登录成功"
		    	ReInfo(1)="<b>"&memLogin("mem_Name")&"</b>，欢迎你的再次光临。<br/><a href=""default.asp"">点击返回主页</a><meta http-equiv=""refresh"" content=""3;url=default.asp""/>"
		    	ReInfo(2)="MessageIcon"
	            ReInfo(3)=true
		   end if
		 end if
		 UserRight(1)
		 memLogin.Close
		 Set memLogin=Nothing
  login2=ReInfo
end function

sub checkCookies()
Dim Guest_IP,Guest_Browser,Guest_Refer
 Guest_IP=getIP()
 Guest_Browser=getBrowser(Request.ServerVariables("HTTP_USER_AGENT"))
 
IF Session("GuestIP")<>Guest_IP Then
	Conn.ExeCute("UPDATE blog_Info SET blog_VisitNums=blog_VisitNums+1")
	SQLQueryNums=SQLQueryNums+1
	getInfo(2)
	Session("GuestIP")=Guest_IP
	if blog_CountNum>0 and Guest_Browser(1)<>"Unkown" then
	    dim tmpC
	    tmpC=conn.execute("select count(coun_ID) as cnt from [blog_Counter]")(0)
		SQLQueryNums=SQLQueryNums+1
		Guest_Refer=Trim(Request.ServerVariables("HTTP_REFERER"))
	    if tmpC>=blog_CountNum then
	        dim tmpLC
	        tmpLC=conn.execute("select top 1 coun_ID from [blog_Counter] order by coun_Time ASC")(0)
			Conn.ExeCute("update [blog_Counter] set coun_Time=#"&now()&"#,coun_IP='"&Guest_IP&"',coun_OS='"&Guest_Browser(1)&"',coun_Browser='"&Guest_Browser(0)&"',coun_Referer='"&HTMLEncode(CheckStr(Guest_Refer))&"' where coun_ID="&tmpLC)
			SQLQueryNums=SQLQueryNums+2
	    else
			Conn.ExeCute("INSERT INTO blog_Counter(coun_IP,coun_OS,coun_Browser,coun_Referer) VALUES ('"&Guest_IP&"','"&Guest_Browser(1)&"','"&Guest_Browser(0)&"','"&HTMLEncode(CheckStr(Guest_Refer))&"')")
			SQLQueryNums=SQLQueryNums+1
	    end if
	end if
End IF  

Dim tempName,tempHashKey
 tempName=CheckStr(Request.Cookies(CookieName)("memName"))
 tempHashKey=CheckStr(Request.Cookies(CookieName)("memHashKey"))
 if tempHashKey="" then 
  logout(false)
 else
  Dim CheckCookie
  Set CheckCookie=Server.CreateObject("ADODB.RecordSet")
  SQL="SELECT Top 1 mem_ID,mem_Name,mem_Password,mem_salt,mem_Status,mem_LastIP,mem_lastVisit,mem_hashKey FROM blog_member WHERE mem_Name='"&tempName&"' AND mem_hashKey='"&tempHashKey&"' AND mem_hashKey<>''"
  CheckCookie.Open SQL,Conn,1,1
  SQLQueryNums=SQLQueryNums+1
  If CheckCookie.EOF AND CheckCookie.BOF Then
    logout(false)
  Else
    UserID=CheckCookie("mem_ID")
    if CheckCookie("mem_LastIP")<>Guest_IP Or isNull(CheckCookie("mem_LastIP")) then
      logout(true)
     else
      memName=CheckStr(Request.Cookies(CookieName)("memName"))
      memStatus=CheckCookie("mem_Status")
    end if
  end if
    CheckCookie.Close
    Set CheckCookie=Nothing
  end if

end sub

sub logout(clearHashKey)
 On Error Resume Next
 if clearHashKey then conn.Execute("UPDATE blog_member set mem_hashKey='' where mem_ID="&UserID)
 If Err Then err.Clear
 Response.Cookies(CookieName)("memName")=""
 Response.Cookies(CookieName)("memHashKey")="" 
 memStatus="Guest"
end sub
%>