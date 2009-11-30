<!--#include file = "../include.asp" --><!--#include file = "../FCKeditor/fckeditor.asp" --><!--#include file = "../pjblog.model/cls_ubbconfig.asp" --><!--#include file = "../pjblog.model/cls_fso.asp" --><!--#include file = "../pjblog.model/cls_Stream.asp" --><!--#include file = "../pjblog.data/cls_webconfig.asp" --><!--#include file = "../pjblog.model/cls_template.asp" --><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="UTF-8">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
	<meta http-equiv="Content-Language" content="UTF-8" />
	<meta name="robots" content="all" />
	<meta name="author" content="<%=blog_email%>, <%=blog_master%>" />
	<meta name="Copyright" content="PJBlog3 CopyRight 2008" />
	<meta name="keywords" content="<%=blog_KeyWords%>" />
	<meta name="description" content="<%=blog_Description%>" />
	<meta name="generator" content="PJBlog3" />
	<link rel="service.post" type="application/atom+xml" title="<%=SiteName%> - Atom" href="<%=SiteURL%>xmlrpc.asp" />
	<link rel="EditURI" type="application/rsd+xml" title="RSD" href="<%=SiteURL%>rsd.asp" />
	<title><%=SiteName%> - <%=blog_Title%></title>
	<link rel="alternate" type="application/rss+xml" href="<%=SiteURL%>feed.asp" title="订阅 I Miss You 所有文章(rss2)" />
	<link rel="alternate" type="application/atom+xml" href="<%=SiteURL%>atom.asp" title="订阅 I Miss You 所有文章(atom)" />
   	<link rel="stylesheet" rev="stylesheet" href="../pjblog.common/blog.css" type="text/css" media="all" />
	<link rel="stylesheet" rev="stylesheet" href="../pjblog.common/editor.css" type="text/css" media="all" />
	<link rel="icon" href="favicon.ico" type="image/x-icon" />
	<link rel="shortcut icon" href="../favicon.ico" type="image/x-icon" />
    <script language="javascript" type="text/javascript" src="../pjblog.common/language.js"></script>
	<script language="javascript" type="text/javascript" src="../pjblog.common/jquery.js"></script>
	<script language="javascript" type="text/javascript" src="../pjblog.common/jqueryform.js"></script>
	<script language="javascript" type="text/javascript" src="../pjblog.common/common.js"></script>
	<script language="javascript" type="text/javascript" src="../pjblog.common/editor.js"></script>
	<script language="javascript" type="text/javascript" src="../pjblog.common/ubbutils.js"></script>
	<script language="javascript" type="text/javascript" src="../pjblog.model/base64.js"></script>  
	<script language="javascript" type="text/javascript">
		$(function() {
			var editor = new UbbEditor('message'); // 传入textarea的Id
		
			// 也可以传入第二个参数，决定默认的编辑模式
			// var editor = new UbbEditor('editor', UbbEditor.EditMode.Visual);
	
			// UbbUtils.smilePath = '../build/smiles/'; // 在需要的时候设置表情路径以保证表情工作正常
		});
	</script>
</head>
<%
	If Asp.ChkPost() Then
    	If stat_AddAll <> True And stat_Add <> True Then
			Response.Write("您没有权限发表日志")
		Else
		' ----------------- 分类列表 --------------------
		  Dim CateArr, Arr, Arri
		  CateArr = Cache.CategoryCache(1)
		  If UBound(CateArr, 1) = 0 Then
			  Arr = ""
		  Else
			  Arr = ""
			  For Arri = 0 To UBound(CateArr, 2)
				  If Not CateArr(4, Arri) Then Arr = Arr & "<option value=""" & CateArr(0, Arri) & """>" & CateArr(1, Arri) & "</option>"
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
<body>
<div class="wrap">
	<div class="head"></div>
    <div class="content">
    	<div id="tool">
        	<div class="info">
            	<div>欢迎 evio 发表日志</div>
                <div style=" margin-top:10px">超级管理员</div>
            </div>
        	<div class="bar">
            	<div class="title">辅助工具</div>
            	<div class="ccontent">
                	<div class="item"><a href="javascript:;">Ajax草稿保存</a></div>
                </div>
            </div>
        </div>
        <div id="main" style="margin-left:200px;">
        	<div style="font-style:oblique; font-size:16px"><strong>发表新日志</strong> - <span id="PostWhere">未分类</span></div>
            <div class="right">
            	<div class="category side">
                	<div class="title">请选择分类</div>
                    <div class="body">
                    	<div>已有分类 
                            <select name="log_CateID" id="log_CateID">
                                <option value="0">请选择分类</option>
                                <%=Arr%>
                            </select>
                        </div>
                        <div>添加新分类</div>
                    </div>
                </div>
                
                <div class="cname side">
                	<div class="title">日志别名</div>
                    <div class="body">
                    	<div><input type="text" value="" class="text" name="log_cname" id="log_cname"> <strong>.</strong> <select name="log_ctype" id="log_ctype"><%=blog_html_option_Str%></select></div>
                        <div>别名是URL友好的另一种写法</div>
                    </div>
                </div>
                
                <div class="tag side">
                	<div class="title">日志标签</div>
                    <div class="body">
                    	<div><input type="text" value="" class="text" style="width:100%"></div>
                        <div>插入已使用的标签</div>
                    </div>
                </div>
                
                <div class="tag side">
                	<div class="title">基本设置</div>
                    <div class="body">
                    	<div>
                        	<div class="jbar"><label for="label1"><input type="checkbox" name="log_IsTop" value="1" id="label1"> 日志置顶</label></div>
                            <div class="jbar"><label for="label2"><input type="checkbox" name="log_DisComment" value="1" id="label2"> 禁止评论</label></div>
                            <div class="jbar"><label for="label3"><input type="checkbox" name="log_comorder" value="1" id="label3"> 评论倒序</label></div>
                        </div>
                        <div style=" margin-top:15px">其他设置</div>
                    </div>
                </div>
                
            </div>
            <div class="left">
            	<div class="article">
                	<div class="title">
                    	<input type="text" value="" class="text" name="log_Title" id="log_Title" />
                    	<div>注意 : 大师傅哈吉斯飞艾丝凡好噶时间</div>
                    </div>
                    <div style="width:100%; margin-top:20px; border:0px solid #ccc; line-height:30px;">
                    	<div style=" padding-left:10px; background:url(../images/control/gray-grad.png) repeat-x; border:1px solid #dfdfdf;">内容编辑</div>
                    	<div><textarea id="message" name="message" style=" width:100%; height:150px" class="text"></textarea></div>
                        <div style="margin-top:10px"><input type="checkbox" onclick="if (this.checked){$('#log_Intro').show();}else{$('#log_Intro').hide();}" name="log_IntroC" id="shC" value="1"> 编辑日志摘要</div>
                        <div><textarea id="log_Intro" name="log_Intro" style="border-left: 1px solid #666666; border-top: 1px solid #666666; border-bottom:1px solid #CCCCCC; border-right: 1px solid #CCCCCC; background:#F9F9F9; width:100%; height:80px; display:none"></textarea></div>   
                    </div>
                    
                    <div style=" padding-left:10px; background:url(../images/control/gray-grad.png) repeat-x; border:1px solid #dfdfdf; line-height:30px; height:30px; margin-top:20px">其他设置</div>
                    
                    <div style="border:1px solid #dfdfdf; margin-top:10px; padding:6px 10px 6px 10px;" class="other">
                    	<div class="item">
                        	<div class="title">1. SEO设置</div>
                            <div class="body">
                            	<div><input type="checkbox" value="1" name="log_Meta" id="log_Meta"> 自定义Meta &nbsp;&nbsp;<input type="checkbox" value="" name="log_ping" id="log_ping" /> 自动发送ping</div>
                            	<div>日志关键字 <input name="log_KeyWords" type="text" id="log_KeyWords" title="<%=lang.Set.Asp(60)%>" size="80" maxlength="254" class="text" /></div>
                                <div style="margin-top:10px">日志描述 <input name="log_KeyWords" type="text" id="log_KeyWords" title="<%=lang.Set.Asp(60)%>" size="80" maxlength="254" class="text" /></div>
                            </div>
                        </div>
                    </div>
                    
                </div>
            </div>
        </div>
        <div class="clear"></div>
    </div>
</div>
</body>
<%
		End If
	Else
		Response.Write("不允许外部提交")	
	End If
%>
</html><%Sys.Close%>