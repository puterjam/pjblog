<!--#include file="BlogCommon.asp" -->
<!--#include file="common/function.asp" -->
<!--#include file="common/library.asp" -->
<!--#include file="common/cache.asp" -->
<!--#include file="common/checkUser.asp" -->
<!--#include file="control/f_control.asp" -->
<%
Response.expires = -1
Response.AddHeader"pragma","no-cache"
Response.AddHeader"cache-control","no-store"
Session.CodePage = 65001
Dim Id, Rs, Antidown, Path
Id = Request("id")
Antidown = Split(blog_UpLoadSet,"|")(0)
If id = "" or not IsInteger(id) then
	showmsg lang.Tip.SysTem(1), lang.Err.info(8) & "<br/><a href=""default.asp"">" & lang.Tip.SysTem(2) & "</a>", "ErrorIcon", ""
Else
    Set Rs = conn.execute("select FilesPath,FilesCounts from blog_Files where id ="&id)
	If Rs.eof and Rs.bof then
	  Set Rs = Nothing
	  showmsg lang.Tip.SysTem(1), lang.Err.info(8) & "<br/><a href=""default.asp"">" & lang.Tip.SysTem(2) & "</a>", "ErrorIcon", ""
	Else
      Path = Rs(0)
	  conn.Execute("update blog_Files set FilesCounts='"&Rs(1)+1&"' where id ="&id)
	  Set Rs = Nothing
	  If left(Path,7)="http://" or left(Path,7)="rtsp://" then Antidown = 0
	  If Request("code") = right(md5(right(Ucase(path),15)),10) then Antidown = 0
      If Antidown = 1 then
        If Session(CookieName & "Antimdown") = "pjblog_Antimdown" Or ChkPost Or Request("code") = right(md5(right(Ucase(path),15)),10) Or left(Path,7)="http://" or left(Path,7)="rtsp://" Then
		  Response.ContentType=GetContentType(getFileInfo(path)(9))
		  Response.AddHeader "Content-Disposition", "null;filename="&Year(now)&Month(now)&Day(now)&Hour(now)&Minute(now)&Second(now)&""&Mid(path,InStrRev(path,"."))&""
		  Response.Binarywrite ReadBinaryFile(server.mappath(path))
        Else
	      showmsg lang.Tip.SysTem(1), lang.Err.info(9) & "<br/><a href=""default.asp"">" & lang.Tip.SysTem(2) & "</a>", "ErrorIcon", ""
        End If
	  Else
	    Response.Redirect path
	End If
  End If
End If

Function GetContentType(FileType)
Select Case lcase(FileType)
     Case "asf"
           GetContentType = "video/x-ms-asf"
     Case "avi"
           GetContentType = "video/x-msvideo"
	 Case "mid"
           GetContentType = "audio/midi"
	 Case "midi"
           GetContentType = "audio/x-midi"
	 Case "bmp"
           GetContentType = "image/bmp"
	 Case "png"
           GetContentType = "image/png"
     Case "gif"
           GetContentType = "image/gif"
     Case "jpg"
           GetContentType = "image/jpeg"
     Case "jpeg"
           GetContentType = "image/jpeg"
     Case "ram"
           GetContentType = "video/vnd.rn-realvideo"
     Case "rm"
           GetContentType = "video/vnd.rn-realvideo"
     Case "rmvb"
           GetContentType = "video/vnd.rn-realvideo"
     Case "rpm"
           GetContentType = "audio/x-pn-realaudio-plugin"
     Case "ra"
           GetContentType = "video/vnd.rn-realvideo"
     Case "wav"
           GetContentType = "audio/wav"
     Case "wmv"
           GetContentType = "video/x-ms-wmv"
     Case "wma"
           GetContentType = "audio/x-ms-wma"
     Case "mp3"
           GetContentType = "audio/mpeg3"
     Case "mpg"
           GetContentType = "video/mpeg"
     Case "mpeg"
           GetContentType = "video/mpeg"
     Case "flv"
           GetContentType = "video/x-flv"
     Case "swf"
           GetContentType = "application/x-shockwave-flash"
     Case Else
           GetContentType = "application/ms-download"
End Select
End Function

Function ReadBinaryFile(FileName)
	Const adTypeBinary = 1
	Dim BinaryStream
	Set BinaryStream = CreateObject("ADODB.Stream")
	BinaryStream.Type = adTypeBinary
	BinaryStream.Open
	BinaryStream.LoadFromFile FileName
	ReadBinaryFile = BinaryStream.Read
End Function
%>