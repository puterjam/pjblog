<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file = "../pjblog.model/cls_upload.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>

<body>
<%
if request.QueryString("c") = "up" then
dim h, r, u
set h = new UpLoadClass
h.Charset="UTF-8"
h.FileType = "jpg"
h.MaxSize = 1024000
h.SavePath = "../upload/"
h.open()
set h = nothing
end if
%>
	<form action="?c=up" method="post" enctype="multipart/form-data">
	<input type="file" name="f-1" value="" />
    <input type="file" name="f-2" value="" />
     <input type="file" name="f-3" value="" />
     <input type="text" value="" name="33" />
    <input type="submit" value="submit" />
    </form>
</body>
</html>
