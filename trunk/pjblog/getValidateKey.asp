<!--#include file="const.asp" -->
<!--#include file="common/function.asp" -->
<!--#include file="common/sha1.asp" -->

<%
	Response.Expires = -9999
	Response.AddHeader "pragma", "no-cache"
	Response.AddHeader "cache-ctrol", "no-cache"

	If Not ChkPost() Then
		Session("GetCode") = empty
		response.end
	End If
		
	'处理trackback的关键key
	if request("type") = "trackback" then
		if len(request("vcode")) = 0 then 
			response.write "setTBKey('codeError');"
			Session("GetCode") = empty
			response.end
		end if
		
		if Int(request("vcode")) <> Int(Session("GetCode")) then
			response.write "setTBKey('codeError');"
			Session("GetCode") = empty
			response.end
		end if
		
		dim tbID
		tbID = Request.QueryString("tbID")
		if IsInteger(tbID) then 
			dim tbKey,mi
			mi = int(Minute(now()) / 10)
			tbKey = sha1(tbID & getServerKey & mi)
			response.write "setTBKey('"&tbKey&"');"
		else
			response.write "setTBKey('');"
		end if 
		Session("GetCode") = empty
	end if

	'处理评论的关键key
	if request("type") = "comment" then
		
	end if
%>