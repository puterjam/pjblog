<!--#include file="const.asp" -->
<!--#include file="common/function.asp" -->
<!--#include file="common/sha1.asp" -->

<%
Response.Expires = -9999
Response.AddHeader "pragma", "no-cache"
Response.AddHeader "cache-ctrol", "no-cache"

If Not ChkPost() Then
    Session("GetCode") = Empty
    response.End
End If

'处理trackback的关键key
If request("type") = "trackback" Then
    If Len(request("vcode")) = 0 Then
        response.Write "setTBKey('codeError','');"
        Session("GetCode") = Empty
        response.End
    End If

    If Int(request("vcode")) <> Int(Session("GetCode")) Then
        response.Write "setTBKey('codeError','');"
        Session("GetCode") = Empty
        response.End
    End If

    Dim tbID
    tbID = Request.QueryString("tbID")
    If IsInteger(tbID) Then
        Dim tbKey, mi, baseUrl
        mi = Int(Minute(Now()) / 10)
        tbKey = sha1(tbID & getServerKey & mi)
		baseUrl = "http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("URL")
		baseUrl = Left(baseUrl, InStrRev(baseUrl,"/"))
        response.Write "setTBKey('"&tbKey&"','"&baseUrl&"');"
    Else
        response.Write "setTBKey('','');"
    End If
    Session("GetCode") = Empty
End If

'处理评论的关键key
If request("type") = "comment" Then

End If
%>
