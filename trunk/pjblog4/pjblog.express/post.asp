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
			Arr = Arr & "<div class=""cateItem"">"
            Arr = Arr & 	"<input type=""radio"" value=""" & CateArr(0, Arri) & """ name=""log_CateID""> " & CateArr(1, Arri)
            Arr = Arr & "</div>"
		End If
	Next
End If

' ----------------- 别名后缀列表 --------------------
		  Dim blog_html_option, blog_html_option_i, blog_html_option_Str
		  blog_html_option_Str = ""
		  blog_html_option = Split(blog_html, ",")
		  For blog_html_option_i = 0 To Ubound(blog_html_option)
			  blog_html_option_Str = blog_html_option_Str & "<option value=""" & blog_html_option(blog_html_option_i) & """>" & blog_html_option(blog_html_option_i) & "</option>"
		  Next
%>
<div class="title" title="点击此处可拖动" id="blogPost" style="cursor:move">
	<div class="left">发表新日志</div>
    <div class="right" onclick="$('#cateAdd').eBoxClose();">关闭</div>
    <div class="clear"></div>
</div>
<div class="mainBody" style="height:280px!important;">
	<ul id="blogSlide">
    	<li>日志内容</li>
    	<li>标题分类</li>
        <li>日志预览</li>
        <li>其他设置</li>
        <li>附件设置</li>
        <div class="clear"></div>
    </ul>
    <ul id="blogSlideContent">
    
    	<li>
        	<div class="conArticle">
                <div class="aContent">
                	<textarea class="text" style="width:100%; height:160px;" name="Message"></textarea><br />
                    <input type="checkbox" value="1"> 禁止自动转换链接
                </div>
            </div>
        </li>
        
    	<li>
        	<div class="conArticle">
            	<div class="aTitle lineheight30">
                	日志标题<br />
                	<input type="text" class="text" value="" style="width:100%; height:30px; line-height:28px; font-size:18px;">
                </div>
                <div class="aCate">
                	 选择分类
                     <div class="cateList"><%=Arr%><div class="clear"></div></div>
                </div>
            </div>
        </li>
	
        <li>
        	<div class="conArticle">
            	<div class="lineheight30">
            	<label for="b1">
                	使用自定义日志预览内容 <input type="checkbox" value="1">
                </label>
                </div>
                <div>
                	<textarea class="text" style="width:100%; height:200px;"></textarea>
                </div>
            </div>
        </li>
        <li>
        	<div>
            
            	<div class="lineheight30" style=" border-bottom:1px solid #efefef; margin-bottom:10px;">博客基本设置</div>
                <div style="padding:0 20px;">
                    <div>一些设置 <label for="label1"><input type="checkbox" name="log_IsTop" value="1" id="label1"> 日志置顶</label>
                        <label for="label2"><input type="checkbox" name="log_DisComment" value="1" id="label2"> 禁止评论</label>
                        <label for="label3"><input type="checkbox" name="log_comorder" value="1" id="label3"> 评论倒序</label>
                        <label for="label4"><input type="checkbox" name="log_comorder" value="1" id="label4"> 自动发送ping</label></div>
                    
                    <div class="lineheight30">
                        日志别名 <input type="text" value="" class="text" style="width:200px;" /> <strong>.</strong> <select><%=blog_html_option_Str%></select>
                    </div>
                    <div class="lineheight30">
                        日志标签  <input type="text" value="" class="text" style="width:200px;" /> <span style="margin-left:10px;">插入已使用的标签</span>
                    </div>
                </div>
                
                <div class="lineheight30" style=" border-bottom:1px solid #efefef; margin-top:10px;margin-bottom:10px;">博客SEO设置</div>
                <div class="lineheight30" style="padding:0 20px;">
                	<div>日志关键字 <input type="text" value="" class="text" style="width:300px;" /></div>
                    <div class="lineheight30">
                		<div>本日志描述 <input type="text" value="" class="text" style="width:300px;" /></div>
                	</div>
                </div>
            </div>
        </li>
        <li>5</li>
    </ul>
    
</div>
<div class="blogSlideBottom"><input type="submit" value="保存" class="button"></div>
<%
Set Sys = Nothing
%>
