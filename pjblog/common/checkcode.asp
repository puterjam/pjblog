<%
Response.CharSet="utf-8"
Response.ContentType="text/html"
Dim CheckValue
CheckValue = Trim(Request("value"))

If (memName=empty or blog_validate=true) and (cstr(lcase(Session("GetCode")))<>cstr(lcase(CheckValue)) or IsEmpty(Session("GetCode"))) Then
	Response.Write ErrCode
Else
	Response.Write YesCode
End If

Function ErrCode()
	ErrCode = "<img src=""images/code_err.gif"" border=""0""/>"
End Function

Function YesCode()
	YesCode = "<img src=""images/code_ok.gif"" border=""0""/>"
End Function
%>