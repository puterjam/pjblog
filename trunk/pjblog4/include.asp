<!--#include file = "pjblog.data/cls_conn.asp" -->
<!--#include file = "pjblog.common/cls_code.asp" -->
<!--#include file = "pjblog.data/cls_Cache.asp" -->
<!--#include file = "pjblog.logic/log_init.asp" -->
<!--#include file = "pjblog.data/cls_SoeData.asp" -->
<!--#include file = "pjblog.model/cls_ubbcode.asp" -->
<script language="jscript" src="pjblog.model/base64.js" runat="server"></script>
<%
'*******************************************
'  Tag Gravatar
' 设置Gravatar头像信息
' <img alt="Gravatar Icon" src="http://www.gravatar.com/avatar/email md5?d=identicon&s=80&r=g"/>
'*******************************************
Class Gravatar
	Public Gravatar_d, Gravatar_s, Gravatar_r, Gravatar_EmailMd5
	Private Sub Class_Initialize()
        Gravatar_d = "http%3A%2F%2Fwww.gravatar.com%2Favatar%2Fad516503a11cd5ca435acc9bb6523536%3Fs%3D35" ' 默认图片，如d=http%3A%2F%2Fexample.com%2Fimages%2Fexample.jpg(其中“%3A”代“:”，“%2F”代“/”)，也可以用三个特殊参数：identicons、monsterids、wavatars
        Gravatar_s = "40" ' 图片大小，单位是px，默认是80，可以取1~512之间的整数
        Gravatar_r = "g" ' 限制等级，默认为g，(G 普通级、PG 辅导级、R 和 X 为限制级)
        Gravatar_EmailMd5 = "" ' 邮箱的MD5值
    End Sub

    Private Sub Class_Terminate()

    End Sub
    
    Public Function outPut()
    	outPut = Lcase("http://www.gravatar.com/avatar/" & Gravatar_EmailMd5 & "?s=80&r=" & Gravatar_r & "&d=" & Gravatar_d)
    End Function
End Class
%>