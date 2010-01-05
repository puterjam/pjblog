<!--#include file = "../include.asp" -->

<%
Dim CateArr, Arr, Arri
CateArr = Cache.CategoryCache(1)
If UBound(CateArr, 1) = 0 Then
	Arr = ""
Else
	Arr = ""
	For Arri = 0 To UBound(CateArr, 2)
		If Not CateArr(4, Arri) Then
			Arr = Arr & "<div class=""item"">"
            Arr = Arr & 	"<div class=""icon""><img src=""../"&CateArr(6, Arri)&"""></div>"
            Arr = Arr & 	"<div class=""txt"">" & CateArr(1, Arri) & "</div>"
            Arr = Arr & 	"<div class=""id"">" & CateArr(0, Arri) & "</div>"
            Arr = Arr & "</div>"
		End If
	Next
End If
%>
<div class="xml" style="display:none"><%=Arr%></div>
<div class="blog-wrap">
	<div class="blog-line">
    	<div class="blog-top">
        	<div class="box-title">发表新日志</div>
            <div class="box-close" onClick="blog.close()"></div>
            <div class="box-tip"></div>
            <div class="clear"></div>
        </div>
        <div class="blog-body">
        	<div class="blog-title">
            	<input type="text" value="" class="input-text" style="width:520px;">
            </div>
            <div class="blog-action">
            	<ul id="cate-select">
                	<li class="ig"><img src="../images/wenhao.png"></li>
                    <li class="tnt">请选择分类</li>
                    <li class="igput"><img src="../images/select.gif"/></li>
                </ul>
            </div>
            <div class="blog-content">
            	<textarea style="width:100%; height:200px;"></textarea>
            </div>
            <div class="blog-submit">
            	<input type="submit" value="保存新日志">
            </div>
        </div>
    </div>
</div>
<%
Set Sys = Nothing
%>
