<!--#include file="conCommon.asp" -->
<!--#include file="common/function.asp" -->
<!--#include file="common/library.asp" -->
<!--#include file="common/cache.asp" -->
<!--#include file="common/checkUser.asp" -->
<!--#include file="FCKeditor/fckeditor.asp" -->
<!--#include file="common/XML.asp" -->
<!--#include file="common/ModSet.asp" -->
<!--#include file="class/cls_logAction.asp" -->
<!--#include file="class/cls_article.asp" -->
<!--#include file="class/cls_control.asp" -->
<%
'***************PJblog2 后台管理页面*******************
' PJblog2 Copyright 2005
' Update:2006-6-10
'**************************************************
If Not ChkPost() Then
    session(CookieName&"_System") = ""
    session(CookieName&"_disLink") = ""
    session(CookieName&"_disCount") = ""

%>
   <script>try{top.location="default.asp"}catch(e){location="default.asp"}</script>
  <%
End If
If session(CookieName&"_System") = True And memName<>Empty And stat_Admin = True Then
    Dim i
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="UTF-8">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<meta http-equiv="Content-Language" content="UTF-8" />
<meta name="author" content="puter2001@21cn.com,PuterJam" />
<meta name="Copyright" content="PL-Blog 2 CopyRight 2005" />
<meta name="keywords" content="PuterJam,Blog,ASP,designing,with,web,standards,xhtml,css,graphic,design,layout,usability,ccessibility,w3c,w3,w3cn" />
<meta name="description" content="PuterJam's BLOG" />
<link rel="stylesheet" rev="stylesheet" href="common/control.css" type="text/css" media="all" />
<script type="text/javascript" src="common/control.js"></script>
<title>后台管理-内容</title>
<style type="text/css">
<!--
.style1 {color: #999}
-->
</style>
</head>
<body class="ContentBody">
<div class="MainDiv">
<%
If Request.QueryString("Fmenu") = "General" Then '站点基本设置
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
  <tr>
      <th class="CTitle"><%=categoryTitle%></th>
  </tr>
  <tr>
    <td class="CPanel">
    <div class="SubMenu"><a href="?Fmenu=General">设置基本信息</a> | <a href="?Fmenu=General&Smenu=visitors">查看访客记录</a> | <a href="?Fmenu=General&Smenu=Misc">初始化数据</a> | <a href="?Fmenu=General&Smenu=clear">清理服务器缓存</a></div>
<%
If Request.QueryString("Smenu") = "visitors" Then
%>
	   <table border="0" cellpadding="2" cellspacing="1" class="TablePanel">
<%
If CheckStr(Request.QueryString("Page"))<>Empty Then
    Curpage = CheckStr(Request.QueryString("Page"))
    If IsInteger(Curpage) = False Or Curpage<0 Then Curpage = 1
Else
    Curpage = 1
End If
Dim bCounter, PageC
Set bCounter = Server.CreateObject("ADODB.RecordSet")
SQL = "SELECT * FROM blog_Counter order by coun_Time desc"
bCounter.Open SQL, Conn, 1, 1
PageC = 0
If Not bCounter.EOF Then
    bCounter.PageSize = 30
    bCounter.AbsolutePage = CurPage
    Dim bCounter_nums
    bCounter_nums = bCounter.RecordCount
    response.Write "<tr><td colspan=""6"" style=""border-bottom:1px solid #999;""><div class=""pageContent"">"&MultiPage(bCounter_nums, 30, CurPage, "?Fmenu=General&Smenu=visitors&", "", "float:left","")&"</div><div class=""Content-body"" style=""line-height:200%""></td></tr>"
%>
        <tr align="center">
          <td width="100" nowrap="nowrap" class="TDHead">访客IP</td>
          <td width="120" nowrap="nowrap" class="TDHead">访客操作系统</td>
          <td nowrap="nowrap" class="TDHead">访客浏览器</td>
          <td width="300" nowrap="nowrap" class="TDHead">访客介入地址</td>
          <td class="TDHead">访客访问时间</td>
	   </tr>
	   <%
Do Until bCounter.EOF Or PageC = bCounter.PageSize

%>
        <tr align="center">
          <td nowrap><%=bCounter("coun_IP")%></td>
          <td nowrap><%=bCounter("coun_OS")%></td>
          <td nowrap><%=bCounter("coun_Browser")%></td>
          <td nowrap align="left"><a href="<%=bCounter("coun_Referer")%>" target="_blank" title="<%=bCounter("coun_Referer")%>"><%=CutStr(bCounter("coun_Referer"),40)%></a></td>
          <td nowrap><%=bCounter("coun_Time")%></td>
	   </tr>
   <%
bCounter.MoveNext
PageC = PageC + 1
Loop
bCounter.Close
Set bCounter = Nothing
response.Write ("</table>")
End If
ElseIf Request.QueryString("Smenu") = "clear" Then
        Response.Write "<div style='padding:4px 0px 4px 10px;border: 1px dotted #999;margin:2px;background:#ffffee'>"
        Application.Lock
        FreeApplicationMemory
        Application.UnLock
        Response.Write "<br/><span><b style='color:#040'>缓存清理完毕...	</b></span>"
        Response.Write "</div>"
ElseIf Request.QueryString("Smenu") = "Misc" Then
%>
<form action="ConContent.asp" method="post" style="margin:0px">
	<input type="hidden" name="action" value="General"/>
	<input type="hidden" name="whatdo" value="Misc"/>
    <div align="left" style="padding:5px;"><%getMsg%>
    
   <b>1.基础数据初始化</b><br/>
    <input type="checkbox" name="ReBulid" value="1" id="T1"/> <label for="T1">重建数据缓存</label><br/>
    <input type="checkbox" name="ReTatol" value="1" id="T2"/> <label for="T2">重新统计网站数据</label><br/>
    <input type="checkbox" name="CleanVisitor" value="1" id="T5"/> <label for="T5">清除访客记录</label><br/>
    
    <br/>
   <b>2.日志缓存和静态日志重建</b><span style="color:#f00">（这个过程可能会花很长时间,由你的日志数量来决定）</span><br/>
    <input type="radio" name="ReBulidArticle" value="1" id="B1"/> <label for="B1">更新所有日志到文件，并且包含日志列表缓存 <span style="color:#666">（静态化所有日志内容数据，速度较慢）</span></label> <br/>
    <input type="radio" name="ReBulidArticle" value="2" id="B2"/> <label for="B2">只更新日志列表缓存<span style="color:#666">（在半静态和全静态之间切换的时候需要重新生成）</span></label><br/>
    <input type="radio" name="ReBulidArticle" value="0" id="B3" checked	/> <label for="B3">什么都不做</label><br/><br/>
    
    <b>3.日志列表索引</b><br/>
    <input type="checkbox" name="ReBulidIndex" value="1" id="T4"/> <label for="T4">重新建立日志列表翻页索引<span style="color:#666">（可以修复日志列表翻页错误的问题）</span></label><br/>
   </div>
   <div class="SubButton">
      <input type="submit" name="Submit" value="保存配置" class="button"/> 
     </div>
	 </form>
	 <%else%>
	 	<form action="ConContent.asp" method="post">
	<input type="hidden" name="action" value="General"/>
	<input type="hidden" name="whatdo" value="General"/>
	   <%getMsg%>
<fieldset>
    <legend> 站点基本信息</legend>
    <div align="left">      
      <table border="0" cellpadding="2" cellspacing="1">
        <tr>
          <td width="180"><div align="right"> BLOG 名称 </div></td>
          <td align="left"><input name="SiteName" type="text" size="30" class="text" value="<%=SiteName%>"/></td>
        </tr>
		<tr>
          <td width="180"><div align="right"> BLOG 副标题 </div></td>
          <td align="left"><input name="blog_Title" type="text" size="50" class="text" value="<%=blog_Title%>"/></td>
        </tr>
		<tr>
          <td width="180"><div align="right"> 站长昵称 </div></td>
          <td align="left"><input name="blog_master" type="text" size="10" class="text" value="<%=blog_master%>" maxlength="10"/></td>
        </tr>
		<tr>
          <td width="180"><div align="right"> 站长邮件地址 </div></td>
          <td align="left"><input name="blog_email" type="text" size="50" class="text" value="<%=blog_email%>"/></td>
        </tr>
        <tr>
          <td width="180"><div align="right"> BLOG 地址
                  <div class="shuom">关系到<strong>RSS</strong>地址的可读性</div>
          </div></td>
          <td align="left"><input name="SiteURL" type="text" size="50" class="text" value="<%=SiteURL%>"/></td>
        </tr>
		<tr>
          <td width="180"><div align="right"> 网站备案信息 </div></td>
          <td align="left"><input name="blog_about" type="text" size="50" class="text" value="<%=blogabout%>"/></td>
        </tr>        
        <tr>
          <td><div align="right">Blog对外开放</div></td>
          <td align="left"><input name="SiteOpen" type="checkbox" value="1" <%if Application(CookieName & "_SiteEnable")=1 then response.write ("checked=""checked""")%>/></td>
        </tr>
      </table>
    </div>
	</fieldset>
	<fieldset>
    <legend> 显示设置</legend>
    <div align="left">
      <table border="0" cellpadding="2" cellspacing="1">
	  <tr><td width="180" align="right">每页显示日志</td><td width="300"><input name="blogPerPage" type="text" size="5" class="text" value="<%=blogPerPage%>"/> 篇</td></tr>
	  <tr><td width="180" align="right">默认显示模式</td><td width="300"><select name="blog_DisMod"><option value="0">普通</option><option  value="1" <%if blog_DisMod then response.write ("selected=""selected""")%>>列表</option></select></td></tr>
	  <tr><td width="180" align="right">是否在首页显示图片友情链接</td><td width="300"><input name="blog_ImgLink" type="checkbox" value="1" <%if blog_ImgLink then response.write ("checked=""checked""")%>  /> </td></tr>
      </table>
    </div>
	</fieldset>
	<fieldset>
    <legend> 日志保存设置</legend>
    <div align="left">
      <table border="0" cellpadding="2" cellspacing="1">
 	  <tr><td width="180" align="right" valign="top" style="padding-top:8px">日志输出模式</td><td>
 	   
 	   <label for="p1" >
	 	  <input id="p1" name="blog_postFile" type="radio" value="0" <%if blog_postFile = 0 then response.write ("checked=""checked""")%>/> 全动态模式 (不推荐)
	 	  <br/> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:#999">文章数据从数据库里直接获取</span> <br/>
 	  </label>
 	  
 	  <label for="p2" >
	 	  <input id="p2" name="blog_postFile" type="radio" value="1" <%if blog_postFile = 1 then response.write ("checked=""checked""")%>/> 半静态模式 (适合喜欢个性化的用户) 
	 	  <br/> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:#999">把文章缓存成文件，并且插件功能不受影响. </span> <br/>
 	  </label>
 	  
 	  <label for="p3" >
	 	  <input id="p3" name="blog_postFile" type="radio" value="2" <%if blog_postFile = 2 then response.write ("checked=""checked""")%>/> 全静态模式<span style="color:#f00;font-size:8px">new</span> (适合在乎seo和速度的用户) 
	 	  <br/> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:#999">把文章保存成全静态文件. 需要注意的是，这种模式下文章页面无法使用插件. </span> <br/>
 	  </label>
	  <div style="border-top:1px solid #ccc;padding-top:5px;margin-top:5px">
 	  <img src="images/notify.gif"/> <b>温馨提示:</b> 进行半静态和全静态切换后，需要到 <a href="ConContent.asp?Fmenu=General&Smenu=Misc">初始化数据</a> 里更新<b>日志列表缓存</b>。
 	  </div>
 	  <br/>
 	  <%if not CheckObjInstalled("ADODB.Stream") then response.write "<b style='color:#f00'>需要 ADODB.Stream 组件支持</b>"%>
 	  </td></tr>
 	  <tr><td width="180" align="right">日志预览分割类型</td><td><select name="blog_SplitType"><option value="0">按照字符数量分割</option><option  value="1" <%if blog_SplitType then response.write ("selected=""selected""")%>>按照行数分割</option></select> <span style="color:#999">只对重新编辑日志或新建日志有效</span></td></tr>
	  <tr><td width="180" align="right">日志预览最大字符数</td><td><input name="blog_introChar" type="text" size="5" class="text" value="<%=blog_introChar%>"/> 个 <span style="color:#999">只对UBB编辑器有效</span></td></tr>
	  <tr><td width="180" align="right">日志预览切割行数</td><td><input name="blog_introLine" type="text" size="5" class="text" value="<%=blog_introLine%>"/> 行</td></tr>
      </table>
    </div>
	</fieldset>	
	<fieldset>
    <legend> 评论设置</legend>
		<table border="0" cellpadding="2" cellspacing="1">
	    <tr><td width="180" align="right">每页显示评论数</td><td width="300"><input name="blogcommpage" type="text" size="5" class="text" value="<%=blogcommpage%>"/> 篇</td></tr>
		<tr><td width="180" align="right">发表评论时间间隔</td><td width="300"><input name="blog_commTimerout" type="text" size="5" class="text" value="<%=blog_commTimerout%>"/> 秒</td></tr>
		<tr><td width="180" align="right">发表评论字数限制</td><td width="300"><input name="blog_commLength" type="text" size="5" class="text" value="<%=blog_commLength%>"/> 字</td></tr>
		<tr><td width="180" align="right">发表评论必须输入验证码</td><td width="300"><input name="blog_validate" type="checkbox" value="1" <%if blog_validate then response.write ("checked=""checked""")%>  /> <span style="color:#999">可以让会员不用输入验证码，只有全动态模式有效</span> </td></tr>
		<tr><td width="180" align="right">禁用评论UBB代码</td><td width="300"><input name="blog_commUBB" type="checkbox" value="1" <%if blog_commUBB then response.write ("checked=""checked""")%>  /></td></tr>
		<tr><td width="180" align="right">禁用评论贴图</td><td width="300"><input name="blog_commIMG" type="checkbox" value="1" <%if blog_commIMG then response.write ("checked=""checked""")%> /></td></tr>
		</table>
	</fieldset>
	<fieldset>
    <legend> Wap设置</legend>
    <div align="left">
      <table border="0" cellpadding="2" cellspacing="1">
	  <tr><td width="180" align="right">允许使用wap方式浏览Blog</td><td width="300"><input name="blog_wap" type="checkbox" value="1" <%if blog_wap then response.write ("checked=""checked""")%>  /></td></tr>
	  <tr><td width="180" align="right">允许wap使用简单HTML</td><td width="300"><input name="blog_wapHTML" type="checkbox" value="1" <%if blog_wapHTML then response.write ("checked=""checked""")%>  /></td></tr>
	  <tr><td width="180" align="right">允许wap显示图片</td><td width="300"><input name="blog_wapImg" type="checkbox" value="1" <%if blog_wapImg then response.write ("checked=""checked""")%>  /></td></tr>
	  <tr><td width="180" align="right">允许wap保留文章超链接</td><td width="300"><input name="blog_wapURL" type="checkbox" value="1" <%if blog_wapURL then response.write ("checked=""checked""")%>  /></td></tr>
	  <tr><td width="180" align="right">允许通过wap登录</td><td width="300"><input name="blog_wapLogin" type="checkbox" value="1" <%if blog_wapLogin then response.write ("checked=""checked""")%>  /></td></tr>
	  <tr><td width="180" align="right">允许通过wap发评论</td><td width="300"><input name="blog_wapComment" type="checkbox" value="1" <%if blog_wapComment then response.write ("checked=""checked""")%>  /></td></tr>
	  <tr><td width="180" align="right">Wap日志显示数量</td><td width="300"><input name="blog_wapNum" type="text" size="5" class="text" value="<%=blog_wapNum%>"/> 篇</td></tr>
      </table>
    </div>
	</fieldset>
	<fieldset>
    <legend> 用户注册与过滤</legend>
    <div align="left">
      <table border="0" cellpadding="2" cellspacing="1">
  		<tr><td width="180" align="right">不允许注册新用户</td><td width="300"><input name="blog_Disregister" type="checkbox" value="1" <%if blog_Disregister then response.write ("checked=""checked""")%>  /> </td></tr>
  	    <tr><td width="180" align="right">访客记录最大值</td><td width="300"><input name="blog_CountNum" type="text" size="5" class="text" value="<%=blog_CountNum%>"/>  <span style="color:#999">设置为0则不进行任何访客记录</span></td></tr>
     <tr>
         <td width="180" valign="top"><div align="right">注册名字过滤
              <div class="shuom">用"|"分割需要过滤的名字</div>              
              <div class="shuom"></div>
          </div></td>
          <td align="left"><textarea name="Register_UserNames" cols="50" rows="5"><%=Register_UserNames%></textarea></td>
        </tr>
        <tr>
          <td width="180" valign="top"><div align="right">IP过滤
              <div class="shuom">以下Ip地址将无法访问Blog</div>              
              <div class="shuom">使用"|"分割IP地址,IP地址可以包含通配符号"*"用来禁止某个IP段的无法访问Blog</div>
          </div></td>
          <td align="left"><textarea name="FilterIPs" cols="50" rows="5"><%=FilterIPs%></textarea></td>
        </tr>
     </table>
    </div>
	</fieldset>	<div class="SubButton">
      <input type="submit" name="Submit" value="保存配置" class="button"/> 
     </div>
	 </form>
	 <%end if%>
	 </td>
  </tr></table>
<%
ElseIf Request.QueryString("Fmenu") = "Categories" Then '日志分类管理
    Dim Arr_Category, blog_Cate, blog_Cate_Item, Icons_Lists, Icons_List, CateInOpstions
    Dim CategoryListDB
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
  <tr>
    <th class="CTitle"><%=categoryTitle%></th>
  </tr>
  <tr>
    <td class="CPanel">
    <div class="SubMenu"><a href="?Fmenu=Categories">设置日志分类</a> | <a href="?Fmenu=Categories&Smenu=move">批量移动日志</a> | <a href="?Fmenu=Categories&Smenu=del">批量删除日志</a> | <a href="?Fmenu=Categories&Smenu=tag">Tag管理</a></div>
<%
If Request.QueryString("Smenu") = "tag" Then
%>
   <form action="ConContent.asp" method="post" style="margin:0px">
   <input type="hidden" name="action" value="Categories"/>
   <input type="hidden" name="whatdo" value="Tag"/>
      <div align="left" style="padding:5px;"><%getMsg%>

	   <table border="0" cellpadding="2" cellspacing="1" class="TablePanel">
        <tr align="center">
		  <td width="16" nowrap="nowrap" class="TDHead">&nbsp;</td>
          <td width="100" nowrap="nowrap" class="TDHead">Tag名称</td>
          <td  nowrap="nowrap" class="TDHead">日志数量</td>
	   </tr>
	   <%
Dim BTag
Set BTag = conn.Execute("select * from blog_tag order by tag_id desc")
Do Until BTag.EOF

%>
	   <tr align="center">
		  <td><input name="selectTagID" type="checkbox" value="<%=BTag("tag_id")%>"/></td>
          <td><input name="TagID" type="hidden" value="<%=BTag("tag_id")%>"/><input name="tagName" type="text" size="14" class="text" value="<%=BTag("tag_name")%>"/></td>
          <td><input name="tagCount" type="text" size="2" class="text" value="<%=BTag("tag_count")%>" readonly style="background:#ffe"/> 篇</td>
	   </tr>
	   <%
BTag.movenext
Loop

%>
	    <tr align="center" bgcolor="#D5DAE0">
        <td colspan="3" class="TDHead" align="left" style="border-top:1px solid #999"><a name="AddLink"></a><img src="images/add.gif" style="margin:0px 2px -3px 2px"/>添加Tag</td>
       </tr>	
        <tr align="center">
		  <td></td>
          <td><input name="TagID" type="hidden" value="-1"/><input name="tagName" type="text" size="14" class="text"/></td>
          <td></td>
	   </tr>
	  </table>
  <div class="SubButton" style="text-align:left;">
    	 <select name="doModule">
			 <option value="SaveAll">保存所有Tag</option>
			 <option value="DelSelect">删除所选Tag</option>
		 </select>
	  <input type="submit" name="Submit" value="提交" class="button" style="margin-bottom:0px"/> 
     </div>	</form>
    <div style="color:#f00">提示:在发表和编辑日志的同时,也可以直接输入多个tag.系统会自动添加不存在的tag</div>
     
    </div>
</td></tr></table>
</div>
<%
ElseIf Request.QueryString("Smenu") = "move" Then
    Set CategoryListDB = conn.Execute("select * from blog_Category order by cate_local asc, cate_Order desc")
    Do While Not CategoryListDB.EOF
        If Not CategoryListDB("cate_OutLink") Then
            CateInOpstions = CateInOpstions&"<option value="""&CategoryListDB("cate_ID")&""">&nbsp;&nbsp;"&CategoryListDB("cate_Name")&" ["&CategoryListDB("cate_count")&"]</option>"
        End If
        CategoryListDB.movenext
    Loop
    Set CategoryListDB = Nothing
%>
  <form action="ConContent.asp" method="post" style="margin:0px" onSubmit="return CheckMove()">
   <input type="hidden" name="action" value="Categories"/>
   <input type="hidden" name="whatdo" value="move"/>
    <div align="left" style="padding:5px;"><%getMsg%>
    
<select name="source">
<option value="null" style="color:#333">源日志分类</option>
<option value="null" style="color:#333">-----------------</option>

<%=CateInOpstions%>
</select> 移动到 
<select name="target">
<option value="null" style="color:#333">目标日志分类</option>
<option value="null" style="color:#333">-----------------</option>
<%=CateInOpstions%>
</select> <input type="submit" name="Submit" value="移动日志" class="button" style="margin-bottom:-1px;"/>
    </div>
     </form></td>
  </tr>
	 <%
ElseIf Request.QueryString("Smenu") = "del" Then
    Dim TempOsel, FilterO
    Set CategoryListDB = conn.Execute("select * from blog_Category order by cate_local asc, cate_Order desc")
    FilterO = checkstr(request.QueryString("filter"))
    If Len(FilterO)<1 Then FilterO = -1
    Do While Not CategoryListDB.EOF
        If Not CategoryListDB("cate_OutLink") Then
            If Int(FilterO) = CategoryListDB("cate_ID") Then TempOsel = "selected"
            CateInOpstions = CateInOpstions&"<option value="""&CategoryListDB("cate_ID")&""" "&TempOsel&">&nbsp;&nbsp;&nbsp; - "&CategoryListDB("cate_Name")&" ["&CategoryListDB("cate_count")&"]</option>"
            TempOsel = ""
        End If
        CategoryListDB.movenext
    Loop
    Set CategoryListDB = Nothing

%>
  <form action="ConContent.asp" method="post" style="margin:0px">
   <input type="hidden" name="action" value="Categories"/>
   <input type="hidden" name="whatdo" value="batdel"/>
    <div align="left" style="padding:5px;"><%getMsg%>
		&nbsp;过滤器: <select name="filter" onChange="location='?Fmenu=Categories&Smenu=del&filter='+this.value">
			<option value="-1">显示所有日志</option>
				<%=CateInOpstions%>
			<option value="-3" <%if int(FilterO)=-3 then response.write "selected"%>>&nbsp;&nbsp;显示所有隐藏日志</option>
			<option value="-2" <%if int(FilterO)=-2 then response.write "selected"%>>&nbsp;&nbsp;显示所有草稿</option>
		</select>
		
		<table style="font-size:12px;margin:8px 0px 8px 0px">
		<%
Dim DelContent
If Int(FilterO) = -1 Then
    SQL = "select * from blog_content order by log_posttime desc"
ElseIf Int(FilterO) = -2 Then
    SQL = "select * from blog_content where log_IsDraft=true order by log_posttime desc"
ElseIf Int(FilterO) = -3 Then
    SQL = "select * from blog_content where log_IsShow=false order by log_posttime desc"
Else
    SQL = "select * from blog_content where log_CateID="&Int(FilterO)&" and log_IsDraft=false order by log_posttime desc"
End If
Set DelContent = conn.Execute(SQL)
If DelContent.EOF And DelContent.bof Then

%>
				 <tr><td>没有找到符合条件的查询</td></tr>
				 <%Else
    Dim TempImg
    Do Until DelContent.EOF
        If DelContent("log_IsShow") = False Then TempImg = "<img src=""images/icon_lock.gif"" alt="""" border=""0"" style=""margin:0px 0px -3px 2px""/>"
        If Int(FilterO) = -2 Then

%>
						 <tr><td><input name="CID" type="checkbox" value="<%=DelContent("log_ID")%>"/></td><td><a href="blogedit.asp?id=<%=DelContent("log_ID")%>" target="_blank"><%=DelContent("log_ID")%>. <%=DelContent("log_Title")%> <%=TempImg%></a></td></tr>
						 <%
Else

%>
						 <tr><td><input name="CID" type="checkbox" value="<%=DelContent("log_ID")%>"/></td><td><a href="default.asp?id=<%=DelContent("log_ID")%>" target="_blank"><%=DelContent("log_ID")%>. <%=DelContent("log_Title")%> <%=TempImg%></a></td></tr>
						 <%
End If
TempImg = ""
DelContent.movenext
Loop
End If
DelContent.Close
Set DelContent = Nothing

%>
		</table>
&nbsp;<input type="checkbox" name="checkbox" style="margin-bottom:-1px;" onClick="checkAll(this)"/> 全选
<input type="submit" name="Submit" value="删除选中的日志" class="button" style="margin-bottom:-1px;"/>
    </div>
     </form>
     <br/>
	 <%Else
    Set CategoryListDB = conn.Execute("select * from blog_Category order by cate_local asc, cate_Order asc")
    If CheckObjInstalled("Scripting.FileSystemObject") Then Icons_Lists = Split(getPathList("images\icons")(1), "*")

%>
	 <script>
	 	var il = [];
        <%
For Each Icons_List in Icons_Lists
    response.Write ("il.push('"&Icons_List&"');" & Chr(13))
Next

%>
       
       function fillList(o){
	     	var v = o.defaultValue;
	     	for (var i=0;i<il.length;i++){
	     		var n = new Option(il[i],"images/icons/" + il[i]);
	     		o.options.add(n);
	     	}
	     	if (!v) o.selectedIndex = 0; else o.value  = v;
       }
       
       function fillAllList(){
       		var ci = document.getElementsByName("Cate_icons");
       		for (var i=0;i<ci.length;i++)
       			fillList(ci[i]);
       		fillList(document.getElementsByName("New_Cate_icons")[0]);
       }
	 </script>
   <form action="ConContent.asp" method="post" style="margin:0px">
   <input type="hidden" name="action" value="Categories"/>
   <input type="hidden" name="whatdo" value="Cate"/>
   <input type="hidden" name="DelCate" value=""/>
    <div align="left" style="padding:5px;"><%getMsg%>
      <table border="0" cellpadding="2" cellspacing="1" class="TablePanel">
        <tr align="center">
          <td class="TDHead" nowrap>分类图标</td>
          <td class="TDHead" nowrap>标题</td>
          <td class="TDHead" nowrap>提示说明</td>
          <td width="180" class="TDHead" nowrap>外部链接</td>
          <td width="29" class="TDHead" nowrap>排序</td>
          <td class="TDHead" nowrap>位置</td>
          <td class="TDHead" nowrap>保密</td>
          <td class="TDHead" nowrap>日志数量</td>
          <td align="center" class="TDHead">&nbsp;</td>
        </tr>
        <%
Do While Not CategoryListDB.EOF

%>
        <tr id="Catetr_<%=CategoryListDB("cate_ID")%>" style="background:<%
If Int(CategoryListDB("cate_local")) = 1 Then
    response.Write ("#a9c9e9")
ElseIf Int(CategoryListDB("cate_local")) = 2 Then
    response.Write ("#bcf39e")
Else

End If

%>">
          <td align="center" nowrap>
          <img name="CateImg_<%=CategoryListDB("cate_ID")%>" src="<%=CategoryListDB("cate_icon")%>" width="16" height="16" />
         <%if CheckObjInstalled("Scripting.FileSystemObject") then%>
          <select name="Cate_icons" defaultValue="<%=CategoryListDB("cate_icon")%>" onChange="document.images['CateImg_<%=CategoryListDB("cate_ID")%>'].src=this.value;" style="width:120px;"></select>
          <%else%>
          <input name="Cate_icons" type="text" class="text" value="<%=CategoryListDB("cate_icon")%>" size="18" onChange="document.images['CateImg_<%=CategoryListDB("cate_ID")%>'].src=this.value"/>
          <%end if%>
          </td>
          <td><input name="Cate_Name" type="text" class="text" value="<%=CategoryListDB("cate_Name")%>" size="14"/></td>
          <td align="left"><input name="Cate_Intro" type="text" class="text" value="<%=CategoryListDB("cate_Intro")%>" size="30"/></td>
          <td align="left"><input name="cate_URL" type="text" value="<%=CategoryListDB("cate_URL")%>" size="30" class="text" <%if CategoryListDB("cate_count")>0 Then response.write "readonly=""readonly"" style=""background:#e5e5e5"""%>/></td>
          <td align="left"><input name="cate_Order" type="text" value="<%=CategoryListDB("cate_Order")%>" size="2" class="text"/></td>
          <td align="center">
           <select name="Cate_local" onChange="getElementById('Catetr_<%=CategoryListDB("cate_ID")%>').style.backgroundColor=(this.value==1)?'#a9c9e9':(this.value==2)?'#bcf39e':''">
	            <option value="0">同时</option>
	            <option value="1" <%if int(CategoryListDB("cate_local"))=1 then response.write "selected=""selected"""%>>顶部</option>
	            <option value="2" <%if int(CategoryListDB("cate_local"))=2 then response.write "selected=""selected"""%>>侧边</option>
           </select>
          </td>
          <td>
           <select name="cate_Secret" <%If CategoryListDB("cate_OutLink") Then response.write "disabled=""disabled"""%>>
            <option value="0" style="background:#0f0">公开</option>
            <option value="1" style="background:#f99" <%if int(CategoryListDB("cate_Secret")) then response.write "selected=""selected"""%>>保密</option>
           </select>
           <%If CategoryListDB("cate_OutLink") Then response.write "<input type=""hidden"" name=""cate_Secret"" value=""0""/>"%>
           </td>
          <td align="left" nowrap><input type="text" class="text" name="cate_count" value="<%=CategoryListDB("cate_count")%>" size="2" readonly="readonly" style="background:#ffe"/> 篇</td>
          <td align="center">
		   <%if not CategoryListDB("cate_Lock") then response.write ("<a href=""javascript:CheckDel("&CategoryListDB("cate_ID")&")"" title=""删除该分类""><img border=""0"" src=""images/icon_del.gif"" width=""16"" height=""16"" /></a>")%>
           <input type="hidden" name="Cate_ID" value="<%=CategoryListDB("cate_ID")%>"/>
          </td>
        </tr>
        <%
CategoryListDB.movenext
Loop
Set CategoryListDB = Nothing

%>
        <tr align="center" bgcolor="#D5DAE0">
         <td colspan="9" class="TDHead" align="left" style="border-top:1px solid #999"><img src="images/add.gif" style="margin:0px 2px -3px 2px"/>添加日志分类</td>
        </tr>
        <tr>
          <td align="center" nowrap><img name="CateImg" src="images/icons/1.gif" width="16" height="16" />
          <%if CheckObjInstalled("Scripting.FileSystemObject") then%>
	          <select name="New_Cate_icons" onChange="document.images['CateImg'].src=this.value" style="width:120px;"></select>
          <%else%>
          <input name="New_Cate_icons" type="text" class="text" value="images/icons/1.gif" size="18" onChange="document.images['CateImg'].src=this.value"/>
          <%end if%>
          </td>
          <td><input name="New_Cate_Name" type="text" size="14" class="text"/></td>
          <td align="left"><input name="New_Cate_Intro" type="text" class="text" size="30"/></td>
          <td align="left"><input name="New_cate_URL" type="text" size="30" class="text"/></td>
          <td align="left"><input name="New_cate_Order" type="text" size="2" class="text"/></td>
          <td align="center">
           <select name="New_Cate_local">
            <option value="0">同时</option>
            <option value="1">顶部</option>
            <option value="2">侧边</option>
           </select>
          </td>
          <td>
          <select name="New_cate_Secret">
            <option value="0" style="background:#0f0">公开</option>
            <option value="1" style="background:#f99">保密</option>
           </select>
          </td>
          <td align="center">&nbsp;</td>
          <td align="center">&nbsp;</td>
        </tr>
      </table>
      <script>fillAllList()</script>
    </div>
    <div style="color:#f00">如果分类中存在日志，则无法使用外部连接.删除日志分类的时假如分类中存在日志,那么日志也会被删除</div>
	<div class="SubButton">
      <input type="submit" name="Submit" value="保存日志分类" class="button"/> 
     </div>   
     </form></td>
  </tr>
<%end if%>
</td></tr>
</table>
<%
'============================================
'      评论留言管理
'============================================
ElseIf Request.QueryString("Fmenu") = "Comment" Then
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
  <tr>
    <th class="CTitle"><%=categoryTitle%></th>
   </tr>
  <tr>
    <td class="CPanel">
	<div class="SubMenu" style="line-height:160%">
		<b>评论管理:</b> <a href="?Fmenu=Comment">评论管理</a> | <a href="?Fmenu=Comment&Smenu=msg">留言管理</a> | <a href="?Fmenu=Comment&Smenu=trackback">引用管理</a>
		<br/>
		<b>评论过滤:</b> <a href="?Fmenu=Comment&Smenu=spam" title="面向初级用户,提供简单的过滤功能">初级过滤功能</a> | <a href="?Fmenu=Comment&Smenu=reg" title="面向高级级用户,提供功能强大的过滤功能">高级过滤功能</a>
	</div>
   
   <%
If Request.QueryString("Smenu") = "spam" Then
    Dim spamXml, spamList
    Set spamXml = New PXML
    spamXml.XmlPath = "spam.xml"
    spamXml.Open

%>
     	 <div align="left" style="padding-top:5px;"><%getMsg%>
     	 <form action="ConContent.asp" method="post" style="margin:0px" onSubmit="return addSpanKey()" name="filter">
			    <input type="hidden" name="action" value="Comment"/>
			    <input type="hidden" name="doModule" value="updateKey"/>
			    <input type="hidden" name="keyList" id="keyList" value=""/>
     	 <table border="0" cellpadding="2" cellspacing="1" class="TablePanel">
     	 <tr><td class="TDHead">过滤关键字</td></tr>
     	 <tr><td><div style="width:394px;overflow:hidden">
     	 <%
spamList = "<select name=""spamList"" id=""spamList"" size=""20"" multiple=""multiple"" style=""width:400px;margin:-3px 0 -3px -3px"">"
For i = 0 To spamXml.GetXmlNodeLength("key") -1
    spamList = spamList & "<option value=""" & spamXml.SelectXmlNodeItemText("key", i) & """>" & spamXml.SelectXmlNodeItemText("key", i) & "</option>"
Next
response.Write spamList
Set spamXml = Nothing

%>
       </div></td></tr>
       <tr><td style="padding-bottom:5px;background:#FAE1AF"><img src="images/add.gif" alt="" style="margin:0 5px -3px"/>添加关键字：<input id="keyWord" type="text" size="27" class="text" onKeyPress="this.style.backgroundColor='#fff';"/>
       <input type="Submit" name="Submit" value="添加" class="button" style="margin-bottom:-2px"/><input type="button" name="button" value="移除" class="button" style="margin-bottom:-2px" onClick="clearKey()"/>
       </td></tr>
       </table>
       		<div class="SubButton" style="text-align:left;padding:5px;margin:0px">
		       	<button accesskey="s" class="button" style="margin-bottom:0px;margin-left:-5px;" onClick="submitList()" >保存关键字列表(<u>S</u>)</button>
		       
	        </div>          
            <div style="color:#f00"><b>友情提示:</b><br/> - 添加或清除关键字后必须 <b>保存关键字列表</b>，垃圾关键字列表才生效。<br/> - 使用逗号或者空格输入字符串可以一次添加多个关键字<br/> - enter键直接插入关键字 ,用ctrl或shift键多选清除关键字</div>
      </div>
      </form>
      <%
ElseIf Request.QueryString("Smenu") = "reg" Then
    Dim reSpamXml, reSpamList
    Set reSpamXml = New PXML
    reSpamXml.XmlPath = "reg.xml"
    reSpamXml.Open


%>
		 <script src="common/reg.js" Language="javascript"></script>
     	 <div align="left" style="padding-top:5px;"><%getMsg%>
     	 	<form name="reFilter" action="ConContent.asp" method="post" style="margin:0px">
			   <input type="hidden" name="action" value="Comment"/>
			   <input type="hidden" name="doModule" value="updateRegKey"/>
			   <input type="hidden" name="keyList" id="keyList" value=""/>
		       <div align="left" style="padding-top:5px;">
		     	 <table border="0" cellpadding="2" cellspacing="1" class="TablePanel">
		     	 <tr>
			     	 <td class="TDHead" width="150" align="center">规则描述</td>
			     	 <td class="TDHead" width="300" align="center">正则表达式</td>
			     	 <td class="TDHead" align="center">允许匹配次数</td>
		     	 </tr>
		     	 <tr>
		     	 	<td colspan=3 id="reList"></td>
		     	 </tr>
		        <tr align="center" bgcolor="#D5DAE0">
		      	   <td colspan="3" class="TDHead" align="left" style="border-top:1px solid #999"><img src="images/add.gif" style="margin:0px 2px -3px 2px"/>添加新规则</td>
		        </tr>
		     	<tr>
		     	 	<td colspan=3>
		     	 		<div>
		     	 			<input id="reTitle" type="text" class="text" style="width:150px"/>
		 		     	 	<input id="reRules" type="text" class="text" style="width:300px"/>
		     	 			<input id="reTimes" type="text" class="text" style="width:30px" maxlength="3"/>
		     	 			<input name="des" type="button" class="button" value="添加" onClick="addRules()" style="font-size:12px;margin:0px;margin-top:-2px;margin-bottom:-1px;height:20px;"/>
    	 				</div>
		     	 	</td>
		     	 </tr>
		        <tr align="center" bgcolor="#D5DAE0">
		      	   <td colspan="3" class="TDHead" align="left" style="border-top:1px solid #999"><img src="images/html.gif" style="margin:0px 2px -3px 2px"/>测试文本区域</td>
		        </tr>
		     	<tr>
		     	 	<td colspan=3>
		     	 		<textarea id="testArea" name="testArea" style="width:100%;height:100px"><%=reSpamXml.SelectXmlNodeItemText("text",0)%></textarea>
		     	 	</td>
		     	 </tr>
		     	 <tr></tr>
			</table>
       		<div class="SubButton" style="text-align:left;padding:5px;margin:0px">
		       	<button accesskey="t" class="button" style="margin-bottom:0px;margin-left:-5px;" onClick="testRules()">测试(<u>T</u>)</button>
		        <button accesskey="s" class="button" style="margin-bottom:0px" onClick="submitReList()">保存(<u>S</u>)</button>
	        </div>          
            <div style="color:#f00"><b>友情提示:</b>
	            <br/> - <b>允许匹配次数:</b> 允许规则匹配的最大次数
	            <br/> - 当匹配次数设置为0时,说明评论中不允许有次规则的文字
            </div>
			</form>
			<script>
			<%
For i = 0 To reSpamXml.GetXmlNodeLength("key") -1
    response.Write "addRule('" & Replace(reSpamXml.GetAttributes("key", "des", i), "'", "\'") & "','" & Replace(reSpamXml.GetAttributes("key", "re", i), "'", "\'") & "','" & reSpamXml.GetAttributes("key", "times", i) & "');" & Chr(13)
Next

%>
			</script>
		</div>
		<%
Set reSpamXml = Nothing
Else

%>
		<form action="ConContent.asp" method="post" style="margin:0px">
		   <input type="hidden" name="action" value="Comment"/>
		   <input type="hidden" name="doModule" value="DelSelect"/>
		      <div align="left" style="padding-top:5px;"><%getMsg%>
		      <%
Dim blog_Comment, comm_Num, commArr, commArrLen, Pcount, aUrl, pSize, saveButton
Pcount = 0
saveButton = "<input type=""button"" value=""保存全部"" class=""button"" onclick=""updateComment()"" style=""font-weight:bold;margin:0px;margin-bottom:5px;float:right;margin-right:6px""/>"
Set blog_Comment = Server.CreateObject("Adodb.RecordSet")
If Request.QueryString("Smenu") = "trackback" Then
    SQL = "SELECT tb_ID,tb_Intro,tb_Site,tb_PostTime,tb_Title,blog_ID,tb_URL,C.log_Title FROM blog_Content C,blog_Trackback T WHERE T.blog_ID=C.log_ID ORDER BY tb_PostTime desc"
    aUrl = "?Fmenu=Comment&Smenu=trackback&"
    pSize = 100
    saveButton = ""
    response.Write "<input type=""hidden"" name=""whatdo"" value=""trackback""/>"
ElseIf Request.QueryString("Smenu") = "msg" Then
    SQL = "SELECT book_ID,book_Content,book_Messager,book_PostTime,book_IP,book_reply FROM blog_book ORDER BY book_PostTime desc"
    aUrl = "?Fmenu=Comment&Smenu=msg&"
    pSize = 12
    response.Write "<input type=""hidden"" name=""whatdo"" value=""msg""/>"
Else '评论
    SQL = "SELECT comm_ID,comm_Content,comm_Author,comm_PostTime,comm_PostIP,blog_ID,T.log_Title from blog_Comment C,blog_Content T WHERE C.blog_ID=T.log_ID ORDER BY C.comm_PostTime desc"
    aUrl = "?Fmenu=Comment&"
    pSize = 15
    response.Write "<input type=""hidden"" name=""whatdo"" value=""comment""/>"
End If

%>
			       <div style="height:24px;">
				       <input type="button" value="删除所选内容" onClick="DelComment()" class="button" style="margin:0px;margin-bottom:5px;float:right;"/> 
				       <input type="button" value="全选" onClick="checkAll()" class="button" style="margin:0px;margin-bottom:5px;float:right;margin-right:6px"/>
				       <%=saveButton%>
				       <div id="page1" class="pageContent"></div>
			       </div>
			       <div class="msgDiv">
					<%
blog_Comment.Open SQL, Conn, 1, 1
If blog_Comment.EOF And blog_Comment.BOF Then
    response.Write "</div>"
Else
    blog_Comment.PageSize = pSize
    blog_Comment.AbsolutePage = CurPage
    comm_Num = blog_Comment.RecordCount

    commArr = blog_Comment.GetRows(comm_Num)
    blog_Comment.Close
    Set blog_Comment = Nothing
    commArrLen = UBound(commArr, 2)
    'commArr(3,Pcount)
    Do Until Pcount = commArrLen + 1 Or Pcount = pSize
        If Request.QueryString("Smenu") = "trackback" Then

%>
							        <div class="item"><input type="hidden" name="CommentID" value="<%=commArr(0,Pcount)%>"/>
								        <div class="title"><span class="blogTitle"><a href="default.asp?id=<%=commArr(5,Pcount)%>" target="_blank" title="<%=commArr(7,Pcount)%>"><%=CutStr(commArr(7,Pcount),25)%></a></span><input type="checkbox" name="selectCommentID" value="<%=commArr(0,Pcount)%>|<%=commArr(5,Pcount)%>" onClick="highLight(this)"/><img src="images/icon_trackback.gif" alt=""/><b><a href="<%=commArr(6,Pcount)%>" target="_blank"><%=commArr(2,Pcount)%></a></b> <span class="date">[<%=DateToStr(commArr(3,Pcount),"Y-m-d H:I:S")%>]</span></div>
								        <div class="contentTB">
								         <b>标题:</b> <%=checkURL(HTMLDecode(commArr(4,Pcount)))%><br/>
								         <b>链接:</b> <a href="<%=commArr(6,Pcount)%>" target="_blank"><%=commArr(6,Pcount)%></a><br/>
								         <b>摘要:</b> <%=checkURL(HTMLDecode(commArr(1,Pcount)))%><br/>
								        </div>
								    </div>
							      <%
ElseIf Request.QueryString("Smenu") = "msg" Then

%>
							        <div class="item"><input type="hidden" name="CommentID" value="<%=commArr(0,Pcount)%>"/>
								        <div class="title"><input type="checkbox" name="selectCommentID" value="<%=commArr(0,Pcount)%>" onClick="highLight(this)"/><img src="images/reply.gif" alt=""/><b><%=HtmlEncode(commArr(2,Pcount))%></b> <span class="date">[<%=DateToStr(commArr(3,Pcount),"Y-m-d H:I:S")%> | <%=commArr(4,Pcount)%>]</span></div>
								        <div class="content"><textarea name="message_<%=commArr(0,Pcount)%>" onFocus="focusMe(this)" onBlur="blurMe(this)" onMouseOver="overMe(this)" onMouseOut="outMe(this)"><%=UnCheckStr(commArr(1,Pcount))%></textarea></div>
								        <div class="reply"><h5>回复内容:<%if len(trim(commArr(5,Pcount)))<1 or IsNull(commArr(5,Pcount)) then response.write "<span class=""tip"">(无回复留言)</span>"%></h5><textarea name="reply_<%=commArr(0,Pcount)%>" onFocus="focusMe(this)" onBlur="blurMe(this)" onMouseOver="overMe(this)" onMouseOut="outMe(this)" onChange="checkMe(this);document.getElementById('edited_<%=commArr(0,Pcount)%>').value=1"><%=UnCheckStr(commArr(5,Pcount))%></textarea><input id="edited_<%=commArr(0,Pcount)%>" name="edited_<%=commArr(0,Pcount)%>" type="hidden" value="0"></div>
								    </div>
							      <%
Else '评论

%>
							        <div class="item"><input type="hidden" name="CommentID" value="<%=commArr(0,Pcount)%>"/>
								        <div class="title"><span class="blogTitle"><a href="default.asp?id=<%=commArr(5,Pcount)%>" target="_blank" title="<%=commArr(6,Pcount)%>"><%=CutStr(commArr(6,Pcount),25)%></a></span><input type="checkbox" name="selectCommentID" value="<%=commArr(0,Pcount)%>|<%=commArr(5,Pcount)%>" onClick="highLight(this)"/><img src="images/icon_quote.gif" alt=""/><b><%=HtmlEncode(commArr(2,Pcount))%></b> <span class="date">[<%=DateToStr(commArr(3,Pcount),"Y-m-d H:I:S")%> | <%=commArr(4,Pcount)%>]</span></div>
								        <div class="content"><textarea name="message_<%=commArr(0,Pcount)%>" onFocus="focusMe(this)" onBlur="blurMe(this)" onMouseOver="overMe(this)" onMouseOut="outMe(this)"><%=UnCheckStr(commArr(1,Pcount))%></textarea></div>
								    </div>
							      <%
End If
Pcount = Pcount + 1
Loop

%>
						</div>
						<div style="height:24px;">
						       <input type="button" value="删除所选内容" onClick="DelComment()" class="button" style="margin:0px;margin-bottom:5px;float:right;"/> 
						       <input type="button" value="全选" onClick="checkAll()" class="button" style="margin:0px;margin-bottom:5px;float:right;margin-right:6px"/>
						       <%=saveButton%>
						       <div id="page2" class="pageContent"><%=MultiPage(comm_Num,pSize,CurPage,aUrl,"","","")%></div>
						       <script>document.getElementById("page1").innerHTML=document.getElementById("page2").innerHTML</script>
					    </div>
			  </div>
		 </form>
		   <%
End If
Set blog_Comment = Nothing
End If

%>
  </td></tr>
  </table>
<%
'============================================
'      界面设置
'============================================
ElseIf Request.QueryString("Fmenu") = "Skins" Then
    Dim bmID, bMInfo
    Dim blog_module
    Dim PluginsFolders, PluginsFolder, Bmodules, Bmodule, tempB, SubItemLen, tempI
    Dim PluginsXML, DBXML, TypeArray
    TypeArray = Array("sidebar", "content", "function")
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
  <tr>
    <th class="CTitle"><%=categoryTitle%></th>
  </tr>
  <tr>
    <td class="CPanel">
	<div class="SubMenu">
	<b>风格设置: </b> <a href="?Fmenu=Skins">设置外观</a> <br/> 
	<b>插件相关: </b> <a href="?Fmenu=Skins&Smenu=module">设置模块</a> | <a href="?Fmenu=Skins&Smenu=Plugins">已装插件管理</a> | <a href="?Fmenu=Skins&Smenu=PluginsInstall">安装模块插件</a></div>

<%
If Request.QueryString("Smenu") = "module" Then

%>
<script>
	var lastFlag = 7;
	function filterModuleList(flag){
		var items = document.getElementsByTagName("tr");
		for (k in items) {
			if (items[k].name!="moduleItem") {
				continue;
			}
			items[k].style.display = "none";
			
			if (items[k].mType&flag){
				items[k].style.display = "";
			}
		}
		
		document.getElementById("a" + lastFlag).style.fontWeight  = "normal";
		document.getElementById("a" + flag).style.fontWeight  = "bold";
		lastFlag = flag;
	}
</script>
   <div align="left" style="padding:5px;"><%getMsg%>
   <form action="ConContent.asp" method="post" style="margin:0px">
       <input type="hidden" name="action" value="Skins"/>
       <input type="hidden" name="whatdo" value="UpdateModule"/>
       <input type="hidden" name="DoID" value=""/>
       
       <div style="width:630px;">
	     	<div style="float:right;margin-top:15px;">
	     	<b>显示过滤: </b> <a id="a7" href="#" onclick="filterModuleList(7)" style="font-weight:bold">所有</a> | <a id="a1" href="#" onclick="filterModuleList(1)">隐藏</a> | <a id="a2" href="#" onclick="filterModuleList(2)">置顶</a> | <a id="a4" href="#" onclick="filterModuleList(4)">系统</a></div>
			<div class="SubButton" style="text-align:left;padding:5px;margin:0px">
				<select id="d1" name="doModule">
					 <option>对模块进行屏蔽和置顶设置</option>
					 <option>------------------------</option>
					 <option value="dohidden">&nbsp;&nbsp;- 设置隐藏</option>
					 <option value="cancelhidden">&nbsp;&nbsp;- 取消隐藏</option>
					 <option>------------------------</option>
					 <option value="doIndex">&nbsp;&nbsp;- 设置首页独享</option>
					 <option value="cancelIndex">&nbsp;&nbsp;- 取消首页独享</option>
				 </select>
				<input type="submit" name="Submit" value="保存" class="button" style="margin-bottom:0px"/> 
	     	</div>         
     </div>  
     	
      <table border="0" cellpadding="2" cellspacing="1" class="TablePanel">
        <tr align="center">
          <td width="18" class="TDHead"><nobr>&nbsp;</nobr></td>
          <td width="18" class="TDHead">&nbsp;</td>
          <td width="18" class="TDHead">&nbsp;</td>
          <td width="18" class="TDHead"><nobr>&nbsp;</nobr></td>
          <td width="58" class="TDHead"><nobr>类型</nobr></td>
		  <td width="82" class="TDHead"><nobr>模块标识</nobr></td>
		  <td width="160" class="TDHead"><nobr>模块名称</nobr></td>
          <td width="42" class="TDHead"><nobr>排序</nobr></td>
          <td  class="TDHead"><nobr>模块操作</nobr></td>
          </tr>
<%
Dim blogModule,flag
Set blogModule = conn.Execute("select * from blog_module where type<>'function' order by type desc,SortID asc")
Do Until blogModule.EOF
flag = 0
if blogModule("IsHidden") then flag = flag + 1
if blogModule("IndexOnly") then flag = flag + 2
if blogModule("IsSystem") then flag = flag + 4


%>     <tr name="moduleItem" mType="<%=flag%>" id="tr_<%=blogModule("id")%>" align="center" style="background:<%if blogModule("type")="content" then response.write ("#a9c9e9")%>;">
          <td><input type="checkbox" name="selectID" value="<%=blogModule("id")%>"/></td>
          <td><%if blogModule("IsHidden") then response.write "<img src=""images/hidden.gif"" alt=""隐藏模块""/>" %></td>
	      <td><%if blogModule("IndexOnly") then response.write "<img src=""images/top.gif"" alt=""只在首页出现""/>" %></td>
          <td><img name="Mimg_<%=blogModule("id")%>" src="images/<%=blogModule("type")%>.gif" width="16" height="16"/></td>
          <td align="left" style="padding-left:5px;"><input type="hidden" name="mID" value="<%=blogModule("id")%>"/>
	           <%if blogModule("IsSystem") then
	         		 response.write "<input name=""mType"" type=""hidden"" value="""&blogModule("type")&"""> &nbsp;<span style='color:#999'>系统</span> "
	            else
	            %>
		          <select name="mType" onChange="document.getElementById('tr_<%=blogModule("id")%>').style.backgroundColor=(this.value=='content')?'#a9c9e9':'';document.images['Mimg_<%=blogModule("id")%>'].src='images/'+this.value+'.gif'">
			           <option value="sidebar">侧边</option>
			           <option value="content" <%if blogModule("type")="content" then response.write ("selected=""selected""")%>>内容</option>
		          </select>
	        <%end if%>
          </td>
          <td><input name="mName" type="text" size="12" class="text" title="<%=blogModule("name")%>" value="<%=blogModule("name")%>" readonly="readonly" style="background:#ffe;"/></td>
          <td><input name="mTitle" type="text" size="24" class="text" value="<%=blogModule("title")%>" <%if blogModule("name")="ContentList" then response.write ("readonly=""readonly"" style=""background:#e5e5e5;""")%>/></td>
          <td><input name="mOrder" type="text" size="3" class="text" value="<%=blogModule("SortID")%>" <%if blogModule("name")="ContentList" then response.write ("readonly=""readonly"" style=""background:#e5e5e5;""")%>/></td>
          <td align="left"><nobr>
          <%if blogModule("name")<>"ContentList" then %>
            <a href="?Fmenu=Skins&Smenu=editModule&miD=<%=blogModule("id")%>" title="可视化编辑模块"><img border="0" src="images/html.gif" style="margin:0px 2px -3px 0px"/>可视化编辑</a> <a href="?Fmenu=Skins&Smenu=editModuleNormal&miD=<%=blogModule("id")%>" title="编辑HTML源代码"><img border="0" src="images/code.gif" style="margin:0px 2px -3px 0px"/>编辑HTML</a> <%if not blogModule("IsSystem") then%><a href="javascript:DelModule(<%=blogModule("id")%>,'<%=blogModule("IsSystem")%>')" title="删除该模块"><img border="0" src="images/icon_del.gif" style="margin:0px 2px -3px 0px"/>删除</a><%end if%>
          <%end if%>
          </nobr></td>
          </tr>
      <%
blogModule.movenext
Loop

%>
        <tr align="center" bgcolor="#D5DAE0">
         <td colspan="9" class="TDHead" align="left" style="border-top:1px solid #999"><img src="images/add.gif" style="margin:0px 2px -3px 2px"/>添加新模块</td>
        </tr>    
        <tr align="center">
          <td></td>
          <td></td>
	      <td></td>
         <td><img src="images/sidebar.gif" width="16" height="16"/></td>
          <td><input type="hidden" name="mID" value="-1"/><select name="mType">
           <option value="sidebar">侧边</option>
           <option value="content">内容</option>
          </select></td>
          <td><input name="mName" type="text" size="12" class="text" /></td>
          <td><input name="mTitle" type="text" size="24" class="text" /></td>
          <td><input name="mOrder" type="text" size="3" class="text" /></td>
          <td></td>
          </tr>
          </table>
		<div class="SubButton" style="text-align:left;padding:5px;margin:0px">
			<input type="submit" name="Submit" value="保存" class="button" style="margin-bottom:0px"/> 
     </div>          
               <div style="color:#f00">模块标识是唯一标记.一旦确定就无法修改.系统自带的模块不允许删除,内容模块只在首页有效.<br/><b>ContentList</b> 是系统自带的日志列表模块，不允许做任何修改</div>
    </div>
  </form>
 <%
'========================================================
' 可视化编辑模块HTML代码
'========================================================
ElseIf Request.QueryString("Smenu") = "editModule" Then

%>

     <div align="center" style="padding:5px;"><%getMsg%>
   <form action="ConContent.asp" method="post" style="margin:0px" onsubmit="return checkSystem()">
     <%
bmID = Request.QueryString("miD")
If IsInteger(bmID) = True Then
    Set bMInfo = conn.Execute("select * from blog_module where id="&bmID)
    If bMInfo.EOF Or bMInfo.bof Then
        session(CookieName&"_ShowMsg") = True
        session(CookieName&"_MsgText") = "没找到符合条件的模块!"
        RedirectUrl("ConContent.asp?Fmenu=Skins&Smenu=module")
    Else

%>
      <input type="hidden" name="action" value="Skins"/>
      <input type="hidden" name="whatdo" value="editModule"/>
      <input type="hidden" name="DoID" value="<%=bmID%>"/>
      <input type="hidden" name="DoName" value="<%=bMInfo("name")%>"/>
      <input type="hidden" name="editType" value="fck"/>
      <div style="font-weight:bold;font-size:14px;margin:3px;margin-left:0px;text-align:left"><img src="images/<%=bMInfo("type")%>.gif" width="16" height="16" style="margin:0px 4px -3px 0px"/>模块名称: <%=bMInfo("Title")%></div>
      <%
Dim sBasePath
sBasePath = "fckeditor/"
Dim oFCKeditor
Set oFCKeditor = New FCKeditor
oFCKeditor.BasePath = sBasePath
oFCKeditor.Config("AutoDetectLanguage") = True
oFCKeditor.Config("DefaultLanguage") = "zh-cn"
oFCKeditor.Config("FormatSource") = True
oFCKeditor.Config("FormatOutput") = True
oFCKeditor.Config("EnableXHTML") = True
oFCKeditor.Config("EnableSourceXHTML") = True
oFCKeditor.Value = UnCheckStr(bMInfo("HtmlCode"))
oFCKeditor.Create "HtmlCode"

End If
Else
    session(CookieName&"_ShowMsg") = True
    session(CookieName&"_MsgText") = "你提交了非法字符!"
    RedirectUrl("ConContent.asp?Fmenu=Skins&Smenu=module")
End If

%>
	<div class="SubButton">
      <input type="submit" name="Submit" value="保存模块代码" class="button" /> 
      <input type="button" name="Submit" value="返回模块列表" class="button" onClick="location='ConContent.asp?Fmenu=Skins&Smenu=module'"/> 
     </div>	
   </div>
   <script>
   function checkSystem(){
    <%if bMInfo("IsSystem") then%>
     if (confirm("这个是系统内置的模块，修改不当，会使系统不正常。\n确定继续吗？")){
      return true 
     }
     else
     {return false}
    <%end if%>
    return true
   }
   </script>
 <%
'========================================
' 编辑模块HTML代码
'========================================
ElseIf Request.QueryString("Smenu") = "editModuleNormal" Then
%>
     <div align="left" style="padding:5px;"><%getMsg%>
   <form action="ConContent.asp" method="post" style="margin:0px" onsubmit="return checkSystem()">
     <%
bmID = Request.QueryString("Mid")
If IsInteger(bmID) = True Then
    Set bMInfo = conn.Execute("select * from blog_module where id="&bmID)
    If bMInfo.EOF Or bMInfo.bof Then
        session(CookieName&"_ShowMsg") = True
        session(CookieName&"_MsgText") = "没找到符合条件的模块!"
        RedirectUrl("ConContent.asp?Fmenu=Skins&Smenu=module")
    Else

%>
      <input type="hidden" name="action" value="Skins"/>
      <input type="hidden" name="whatdo" value="editModule"/>
      <input type="hidden" name="DoID" value="<%=bmID%>"/>
      <input type="hidden" name="DoName" value="<%=bMInfo("name")%>"/>
      <input type="hidden" name="editType" value="normal"/>
      <div style="font-weight:bold;font-size:14px;margin:3px;margin-left:0px;text-align:left"><img src="images/<%=bMInfo("type")%>.gif" width="16" height="16" style="margin:0px 4px -3px 0px"/>模块名称: <%=bMInfo("Title")%></div>
      <textarea style="width:700px;height:200px" name="HtmlCode"><%=UnCheckStr(bMInfo("HtmlCode"))%></textarea>
      <%
End If
Else
    session(CookieName&"_ShowMsg") = True
    session(CookieName&"_MsgText") = "你提交了非法字符!"
    RedirectUrl("ConContent.asp?Fmenu=Skins&Smenu=module")
End If

%>
	<div class="SubButton" style="text-align:left">
      <input type="submit" name="Submit" value="保存HTML代码" class="button" <%If Not CheckObjInstalled("ADODB.Stream") Then Response.Write("disabled")%>/>
      <input type="button" name="Submit" value="返回模块列表" class="button" onClick="location='ConContent.asp?Fmenu=Skins&Smenu=module'"/> 
     </div>	
   </div>
   <script>
   function checkSystem(){
    <%if bMInfo("IsSystem") then%>
     if (confirm("这个是系统内置的模块，修改不当，会使系统不正常。\n确定继续吗？")){
      return true 
     }
     else
     {return false}
    <%end if%>
    return true
   }
   </script>
<%
ElseIf Request.QueryString("Smenu") = "PluginsInstall" Then
    Bmodule = getModName
    PluginsFolders = Split(getPathList("Plugins")(0), "*")
%>
    <div align="left" style="padding:5px;"><%getMsg%>
      <table border="0" cellpadding="2" cellspacing="1" class="TablePanel">
        <tr align="center">
          <td width="18" class="TDHead">&nbsp;</td>
		  <td width="200" align="left" class="TDHead">名称</td>
		  <td width="82" class="TDHead">作者</td>
          <td width="100" class="TDHead">发布日期</td>
          <td width="140" class="TDHead">&nbsp;</td>
          </tr>
	 <%
If CheckObjInstalled(getXMLDOM()) Then
    If CheckObjInstalled("Scripting.FileSystemObject") Then
        Set PluginsXML = New PXML
        For Each PluginsFolder in PluginsFolders
            PluginsXML.XmlPath = "Plugins/"&PluginsFolder&"/install.xml"
            PluginsXML.Open
            If PluginsXML.getError = 0 Then
                If Len(PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginName"))>0 And IsvalidPlugins(PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginName")) And (Not IsvalidValue(Bmodule, PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginName"))) And IsvalidValue(TypeArray, PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginType")) Then

%>
             <tr>
               <td align="center"><img src="images/<%=PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginType")%>.gif" width="16" height="16"/></td>
     		   <td align="left"><%=PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginTitle")%></td>
     		   <td align="center"><%=PluginsXML.SelectXmlNodeText("PluginInstall/main/Author")%></td>
               <td align="center"><%=PluginsXML.SelectXmlNodeText("PluginInstall/main/pubDate")%></td>
               <td align="center"><a href="?Fmenu=Skins&Smenu=InstallPlugins&Plugins=<%=PluginsFolder%>">安装此插件</a> | <a href="javascript:alert('<%=PluginsXML.SelectXmlNodeText("PluginInstall/main/About")%>')">关于</a></td>
             </tr>
     	<%
End If
End If
PluginsXML.CloseXml()
Next
Else
    response.Write ("<tr><td colspan=""6"" align=""center""><div style=""background:#ffffe8;border:1px solid #95801c;padding:3px;text-align:left;"">你的系统不支持 <b>Scripting.FileSystemObject</b> 只能手动输入插件的文件夹名称</div>")
    response.Write ("<div style=""text-align:left;padding:3px""><b>插件路径:</b> Plugins / <input id=""SPath"" type=""text"" size=""16"" class=""text"" value=""""/> <input type=""button"" value=""安装插件"" class=""button"" style=""margin-bottom:-2px"" onclick=""if (document.getElementById('SPath').value.length>0) {location='ConContent.asp?Fmenu=Skins&Smenu=InstallPlugins&Plugins='+document.getElementById('SPath').value}else{alert('请输入插件路径!')}""/></div>")
    response.Write ("</td></tr>")
End If
Else
    response.Write ("<tr><td colspan=""6"" align=""center""><div style=""background:#ffffe8;border:1px solid #95801c;padding:3px;text-align:left;"">你的系统不支持 <b>"&getXMLDOM()&"</b>，无法使用插件管理功能，请与服务商联系！</div></td></tr>")
End If

%>
      </table>
  <div style="color:#f00;text-align:left">
  此处列出系统找到的合法的PJBlog3插件，安装插件前需要把插件连同其目录一起上传到Plugins文件夹内
  <br/><b>注意:这里只列出没有安装的插件。</b>
  <br/><%If Not CheckObjInstalled("ADODB.Stream") Then %><b>你的服务器不支持</b> <b><a href="http://www.google.com/search?hl=zh-CN&q=ADODB.Stream&btnG=Google+%E6%90%9C%E7%B4%A2&lr=" target="_blank"><b>ADODB.Stream</b></a> 组件,那将意味着大部分插件的无法正常工作</b><%End If%>
  </div>    
</div>
<%


'============================================================
'    安装插件
'============================================================
ElseIf Request.QueryString("Smenu") = "InstallPlugins" Then
    InstallPlugins
    '============================================================
    '    显示已经安装插件
    '============================================================
ElseIf Request.QueryString("Smenu") = "Plugins" Then
    Dim Blog_Plugins
    Set Blog_Plugins = conn.Execute("select * from blog_module where IsInstall=true order by id desc")

%>
    <div align="left" style="padding:5px;"><%getMsg%>
      <table border="0" cellpadding="2" cellspacing="1" class="TablePanel">
        <tr align="center">
          <td width="18" class="TDHead">&nbsp;</td>
		  <td width="150" align="left" class="TDHead">名称</td>
          <td width="150" class="TDHead">插件所在目录</td>
          <td width="150" class="TDHead">安装时间</td>
          <td width="140" class="TDHead">&nbsp;</td>
        </tr>
       <%do until Blog_Plugins.eof%>
        <tr align="center">
          <td ><img src="images/<%=Blog_Plugins("type")%>.gif" width="16" height="16"/></td>
		  <td align="left"><%=Blog_Plugins("title")%></td>
          <td align="left">Plugins/<%=Blog_Plugins("InstallFolder")%>/</td>
          <td ><%=Blog_Plugins("InstallDate")%></td>
          <td>
          <%if len(Blog_Plugins("SettingXML"))>0 then %>
		          <a href="?Fmenu=Skins&Smenu=PluginsOptions&Plugins=<%=Blog_Plugins("name")%>">基本设置</a>
	          <%else%>
		          <span style="color:#999">基本设置</span>
          <%end if%>
           | 
          <%if len(Blog_Plugins("ConfigPath"))>0 then %>
		          <a href="Plugins/<%=Blog_Plugins("InstallFolder")%>/<%=Blog_Plugins("ConfigPath")%>">高级设置</a>
	          <%else%>
		          <span style="color:#999">高级设置</span>
          <%end if%>
           | <a href="javascript:DelPlugins('<%=Blog_Plugins("InstallFolder")%>','<%=Blog_Plugins("type")%>')">反安装此插件</a></td>
          </tr>
           <%
Blog_Plugins.movenext
Loop
Set Blog_Plugins = Nothing

%>
     </table>
       <div style="color:#f00;text-align:left">
       <%If Not CheckObjInstalled("ADODB.Stream") Then %>
	       <b>你的服务器不支持</b> <b><a href="http://www.google.com/search?hl=zh-CN&q=ADODB.Stream&btnG=Google+%E6%90%9C%E7%B4%A2&lr=" target="_blank"><b>ADODB.Stream</b></a> 组件,那将意味着大部分插件的无法正常工作</b>
       <%else%>
	       <input type="button" name="button" value="修复插件" class="button" onClick="FixPlugins()"/>
       <%End If%>
<br/>
 假如插件反安装不成功，请到 <b>数据库与附件-数据库管理</b> 压缩修复数据再反安装
</div>
  <%



'============================================================
'    反安装插件
'============================================================
ElseIf Request.QueryString("Smenu") = "UnInstallPlugins" Then
    Dim UnPlugName, getCateID, DropTable, KeepTable, ModSetTemp1, getMod
    PluginsFolder = CheckStr(Request.QueryString("Plugins"))
    KeepTable = CBool(Request.QueryString("Keep"))
    Set PluginsXML = New PXML
    PluginsXML.XmlPath = "Plugins/"&PluginsFolder&"/install.xml"
    PluginsXML.Open
    If PluginsXML.getError = 0 Then
        UnPlugName = PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginName")
        Set ModSetTemp1 = New ModSet
        ModSetTemp1.Open UnPlugName
        ModSetTemp1.RemoveApplication
        DropTable = PluginsXML.SelectXmlNodeText("PluginInstall/main/DropTable")
        Set getMod = conn.Execute("select CateID from blog_module where name='"&UnPlugName&"'")
        If getMod.EOF Then
            session(CookieName&"_ShowMsg") = True
            session(CookieName&"_MsgText") = "<font color=""#ff0000"">"&UnPlugName&"</font> 无法反安装,数据库没有找到相应的信息!"
            RedirectUrl("ConContent.asp?Fmenu=Skins&Smenu=Plugins")
        Else
            getCateID = getMod(0)
            If Len(getCateID)>0 Then conn.Execute("delete * from blog_Category where cate_ID="&getCateID)
            delPlugins UnPlugName
            If Len(DropTable)>0 And KeepTable = False Then conn.Execute("DROP TABLE "&DropTable)
            SubItemLen = Int(PluginsXML.GetXmlNodeLength("PluginInstall/SubItem/item"))

            For tempI = 0 To SubItemLen -1
                If Not PluginsXML.SelectXmlNodeText("PluginInstall/SubItem/item/PluginType") = "function" Then
                    delPlugins UnPlugName&"SubItem"&(tempI + 1)
                End If
            Next
            If Len(PluginsXML.SelectXmlNodeText("PluginInstall/main/SettingFile"))>0 Then
                If KeepTable = False Then InstallPlugingSetting "", UnPlugName, "delete"
            End If
        End If
    End If

    PluginsXML.CloseXml()
    log_module(2)
    CategoryList(2)
    FixPlugins(0)
    session(CookieName&"_ShowMsg") = True
    session(CookieName&"_MsgText") = "<font color=""#ff0000"">"&UnPlugName&"</font> 插件反安装完成!"
    RedirectUrl("ConContent.asp?Fmenu=Skins&Smenu=Plugins")

    '============================================================
    '    修复插件
    '============================================================
ElseIf Request.QueryString("Smenu") = "FixPlugins" Then
    FixPlugins(1)
    '============================================================
    '    插件配置
    '============================================================
ElseIf Request.QueryString("Smenu") = "PluginsOptions" Then
    Dim PluginsSetting, LoadSetXML, KeyLen, Si, LoadModSet, SelectTemp
    Set PluginsSetting = conn.Execute("select top 1 * from blog_module where name='" + checkstr(Request.QueryString("Plugins")) + "'")
    Set LoadSetXML = New PXML
    Set LoadModSet = New ModSet
    LoadModSet.Open(PluginsSetting("name"))
    LoadSetXML.XmlPath = "Plugins/"&PluginsSetting("InstallFolder")&"/"&PluginsSetting("SettingXML")
    LoadSetXML.Open
    If LoadSetXML.getError = 0 Then
        KeyLen = Int(LoadSetXML.GetXmlNodeLength("PluginOptions/Key"))
        getMsg
        Response.Write ("<div align=""center""><form action=""ConContent.asp"" method=""post"" style=""margin:0px"">")
        Response.Write ("<input type=""hidden"" name=""action"" value=""Skins""/>")
        Response.Write ("<input type=""hidden"" name=""whatdo"" value=""SavePluginsSetting""/>")
        Response.Write ("<input type=""hidden"" name=""PluginsName"" value="""&PluginsSetting("name")&"""/>")
        response.Write "<table border=""0"" cellpadding=""2"" cellspacing=""1"" class=""TablePanel"" style=""margin:6px"">"
        response.Write("<tr><td colspan=""2"" align=""left"" style=""background:#e5e5e5;padding:6px""><div style=""font-weight:bold;font-size:14px;"">"&PluginsSetting("title")&" 的基本设置</div></td></tr>")
        For tempI = 0 To KeyLen -1
            response.Write "<tr><td align=""right"" width=""200"" valign=""top"" style=""padding-top:6px"">"&LoadSetXML.GetAttributes("PluginOptions/Key", "description", TempI)&"</td><td width=""300"">"
            If LCase(LoadSetXML.GetAttributes("PluginOptions/Key", "type", TempI)) = "select" Then
                response.Write "<select name="""&LoadSetXML.GetAttributes("PluginOptions/Key", "name", TempI)&""">"
                For Si = 0 To LoadSetXML.SelectXmlNode("PluginOptions/Key", TempI).getElementsByTagName("option").Length -1
                    If LoadModSet.getKeyValue(LoadSetXML.GetAttributes("PluginOptions/Key", "name", tempI)) = LoadSetXML.SelectXmlNode("PluginOptions/Key", TempI).getElementsByTagName("option").Item(Si).Attributes(0).Value Then SelectTemp = "selected"
                    If LoadSetXML.SelectXmlNode("PluginOptions/Key", TempI).getElementsByTagName("option").Item(Si).Attributes.Length>0 Then
                        Response.Write "<option "&SelectTemp&" value="""&LoadSetXML.SelectXmlNode("PluginOptions/Key", TempI).getElementsByTagName("option").Item(Si).Attributes(0).Value&""">"&LoadSetXML.SelectXmlNode("PluginOptions/Key", TempI).getElementsByTagName("option").Item(Si).text&"</option>"
                    Else
                        Response.Write "<option "&SelectTemp&">"&LoadSetXML.SelectXmlNode("PluginOptions/Key", TempI).getElementsByTagName("option").Item(Si).text&"</option>"
                    End If
                    SelectTemp = ""
                Next
                response.Write "</select></td></tr>"
            ElseIf LCase(LoadSetXML.GetAttributes("PluginOptions/Key", "type", TempI)) = "textarea" Then
                response.Write "<textarea name="""&LoadSetXML.GetAttributes("PluginOptions/Key", "name", tempI)&""" rows="""&LoadSetXML.GetAttributes("PluginOptions/Key", "rows", TempI)&""" cols="""&LoadSetXML.GetAttributes("PluginOptions/Key", "cols", TempI)&""">"&LoadModSet.getKeyValue(LoadSetXML.GetAttributes("PluginOptions/Key", "name", tempI))&"</textarea></td></tr>"
            Else
                response.Write "<input name="""&LoadSetXML.GetAttributes("PluginOptions/Key", "name", TempI)&""" type="""&LoadSetXML.GetAttributes("PluginOptions/Key", "type", TempI)&""" size="""&LoadSetXML.GetAttributes("PluginOptions/Key", "size", TempI)&""" value="""&LoadModSet.getKeyValue(LoadSetXML.GetAttributes("PluginOptions/Key", "name", tempI))&"""/></td></tr>"
            End If
        Next
        response.Write "<tr><td colspan=""2"" align=""center""><input type=""submit"" name=""Submit"" value=""保存设置"" class=""button""/><input type=""button"" value=""放弃返回"" class=""button"" onclick=""history.go(-1)""/></td></tr>"
        response.Write "</table>"
        response.Write "</form></div>"
    Else
        response.Write "无法找到配置文件"
    End If
    Set LoadSetXML = Nothing
    Set PluginsSetting = Nothing
    '============================================================
    '    设置外观
    '============================================================
Else
    Dim SkinFolders, SkinFolder
    SkinFolders = Split(getPathList("skins")(0), "*")
%>
    <div align="left" style="padding:5px;"><%getMsg%>
   <form action="ConContent.asp" method="post" style="margin:0px">
       <input type="hidden" name="action" value="Skins"/>
       <input type="hidden" name="whatdo" value="setDefaultSkin"/>
       <input type="hidden" name="SkinName" value=""/>
       <input type="hidden" name="SkinPath" value=""/>
  </form>
      <table border="0" cellpadding="2" cellspacing="1" class="TablePanel" width="700">
        <tr>
	          <td width="700" class="TDHead" colspan="2">界面列表</td>
        </tr>
	 <%
If CheckObjInstalled(getXMLDOM()) And CheckObjInstalled("Scripting.FileSystemObject") Then
    Dim SkinXML, k, SkinPreview
    k = 2
    Set SkinXML = New PXML
    For Each SkinFolder in SkinFolders
        SkinXML.XmlPath = "skins/"&SkinFolder&"/skin.xml"
        SkinXML.Open
        If SkinXML.getError = 0 Then
            If k / 2 = Int(k / 2) Then response.Write "<tr>"
            SkinPreview = "images/Control/skin.jpg"
            If FileExist("skins/"&SkinFolder&"/Preview.jpg") Then SkinPreview = "skins/"&SkinFolder&"/Preview.jpg"

%>
 		<td width="50%" style='border-bottom:1px dotted #ccc'>
 			<div class="<%if Lcase(blog_DefaultSkin)<>Lcase(SkinFolder) then response.write ("un")%>selectskin">
 			<img src="<%=SkinPreview%>" alt="" border="0" class="skinimg"/>
 			  <div class="skinDes">
 			  <div style="height:38px;overflow:hidden"><b style="color:#004000"><%=SkinXML.SelectXmlNodeText("SkinName")%></b></div>
 			  <span style="height:16px;overflow:hidden;cursor:default" title="设计者:<%=SkinXML.SelectXmlNodeText("SkinDesigner")%>"><B>设计者:</B> <%=SkinXML.SelectXmlNodeText("SkinDesigner")%></span><br/>
 			  <B>发布时间:</B> <%=SkinXML.SelectXmlNodeText("pubDate")%><br/></div>
		 <%
If LCase(blog_DefaultSkin) = LCase(SkinFolder) Then
    response.Write ("<div class=""cskin""><img src=""images/Control/select.gif"" alt="""" border=""0"" /></div>当前界面")
Else
    response.Write ("<a href=""javascript:setSkin('"&SkinFolder&"','"&SkinXML.SelectXmlNodeText("SkinName")&"')"">设置为当前主题</a>")
End If

%>
		  </div>
 		</td>
	<%
End If
SkinXML.CloseXml
If k / 2<>Int(k / 2) Then response.Write "</tr>"
k = k + 1
Next
If k / 2<>Int(k / 2) Then response.Write "</tr>"

Set SkinXML = Nothing
Else
    response.Write ("<tr><td colspan=""2"" align=""center""><div style=""background:#ffffe8;border:1px solid #95801c;padding:3px;text-align:left;"">你的系统不支持 <b>"&getXMLDOM()&"</b> 或 <b>Scripting.FileSystemObject</b> 只能手动输入Skin的文件夹名称</div>")
    response.Write ("<div style=""text-align:left;padding:3px""><b>界面路径:</b> Skins / <input id=""SPath"" type=""text"" size=""16"" class=""text"" value=""" + blog_DefaultSkin + """/> <input type=""button"" value=""保存界面"" class=""button"" style=""margin-bottom:-2px"" onclick=""if (document.getElementById('SPath').value.length>0) {setSkin(document.getElementById('SPath').value,document.getElementById('SPath').value)}else{alert('请输入界面路径!')}""/></div>")
    response.Write ("</td></tr>")
End If

%>
      </table>
</div>
<%end if%>
 </td>
  </tr></table>
<%

ElseIf Request.QueryString("Fmenu") = "SQLFile" Then '数据库与文件
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
  <tr>
    <th class="CTitle"><%=categoryTitle%></th>
  </tr>
  <tr>
    <td class="CPanel">
	<div class="SubMenu"><a href="?Fmenu=SQLFile">数据库管理</a> | <a href="?Fmenu=SQLFile&Smenu=Attachments">附件管理</a></div>
    <div align="left" style="padding:5px;"><%getMsg%>
     <%
If Request.QueryString("Smenu") = "Attachments" Then

%>
   <form action="ConContent.asp" method="post" style="margin:0px" onSubmit="if (confirm('是否删除选择的文件或文件夹')) {return true} else {return false}">
   <input type="hidden" name="action" value="Attachments"/>
   <input type="hidden" name="whatdo" value="DelFiles"/>
     <%
Dim AttPath, ArrFolder, Arrfile, ArrFolders, Arrfiles, arrUpFolders, arrUpFolder, TempF
TempF = ""
AttPath = Request.QueryString("AttPath")
If Len(AttPath)<1 Then
    AttPath = "attachments"
ElseIf bc(server.mapPath(AttPath), server.mapPath("attachments")) Then
    AttPath = "attachments"
End If
ArrFolders = Split(getPathList(AttPath)(0), "*")
Arrfiles = Split(getPathList(AttPath)(1), "*")
response.Write "<div style=""font-weight:bold;font-size:14px;margin:3px;margin-left:0px"">"&AttPath&"</div><div style=""margin:3px;margin-left:0px;"">"
If AttPath<>"attachments" Then
    arrUpFolders = Split(AttPath, "/")
    For i = 0 To UBound(arrUpFolders) -1
        arrUpFolder = arrUpFolder&TempF&arrUpFolders(i)
        TempF = "/"
    Next
End If
If Len(arrUpFolder)>0 Then
    response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=""?Fmenu=SQLFile&Smenu=Attachments&AttPath="&arrUpFolder&"""><img border=""0"" src=""images/file/folder_up.gif"" style=""margin:4px 3px -3px 0px""/>返回上级目录</a><br>"
End If
For Each ArrFolder in ArrFolders
    response.Write "<input name=""folders"" type=""checkbox"" value="""&AttPath&"/"&ArrFolder&"""/>&nbsp;<a href=""?Fmenu=SQLFile&Smenu=Attachments&AttPath="&AttPath&"/"&ArrFolder&"""><img border=""0"" src=""images/file/folder.gif"" style=""margin:4px 3px -3px 0px""/>"&ArrFolder&"</a><br>"
Next
For Each Arrfile in Arrfiles
    response.Write "<input name=""Files"" type=""checkbox"" value="""&AttPath&"/"&Arrfile&"""/>&nbsp;<a href="""&AttPath&"/"&Arrfile&""" target=""_blank"">"&getFileIcons(getFileInfo(AttPath&"/"&Arrfile)(1))&Arrfile&"</a>&nbsp;&nbsp;"&getFileInfo(AttPath&"/"&Arrfile)(0)&" | "&getFileInfo(AttPath&"/"&Arrfile)(2)&" | "&getFileInfo(AttPath&"/"&Arrfile)(3)&"<br>"
Next
response.Write "</div>"

%>
    <div style="color:#f00">如果文件夹内存在文件，那么该文件夹将无法删除!</div>
	  	<div class="SubButton">
      <input type="button" value="全选" class="button" onClick="checkAll()"/>  <input type="submit" name="Submit" value="删除所选的文件或文件夹" class="button"/> 
     </div>
     </form>
	  <%else%>
	  <b>数据库路径:</b> <%=Server.MapPath(AccessFile)%><br/>
	  <b>数据库大小:</b> <span id="accessSize"><%=getFileInfo(AccessFile)(0)%></span><br/>
	  <b>数据库操作:</b> <a href="?Fmenu=SQLFile&do=Compact">压缩修复</a> | <a href="?Fmenu=SQLFile&do=Backup">备份</a><br/>
	  <%
Dim AccessFSO, AccessEngine, AccessSource
'-------------压缩数据库-----------------
If Request.QueryString("do") = "Compact" Then
    Set AccessFSO = Server.CreateObject("Scripting.FileSystemObject")
    If AccessFSO.FileExists(Server.Mappath(AccessFile)) Then
        Response.Write "<div style='padding:4px 0px 4px 10px;border: 1px dotted #999;margin:2px;background:#ffffee'>"
        Response.Write "压缩数据库开始，网站暂停一切用户的前台操作...<br/>"
        Response.Write "关闭数据库操作...<br/>"
        Call CloseDB
        Application.Lock
        FreeApplicationMemory
        Application(CookieName & "_SiteEnable") = 0
        Application(CookieName & "_SiteDisbleWhy") = "网站暂停中，请稍候几分钟后再来..."
        Application.UnLock
        Set AccessEngine = CreateObject("JRO.JetEngine")
        AccessEngine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath(AccessFile), "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath(AccessFile & ".temp")
        AccessFSO.CopyFile Server.Mappath(AccessFile & ".temp"), Server.Mappath(AccessFile)
        AccessFSO.DeleteFile(Server.Mappath(AccessFile & ".temp"))
        Set AccessFSO = Nothing
        Set AccessEngine = Nothing
        Application.Lock
        Application(CookieName & "_SiteEnable") = 1
        Application(CookieName & "_SiteDisbleWhy") = ""
        Application.UnLock
        Response.Write "压缩数据库完成...<br/>"
        Application.Lock
        Application(CookieName & "_SiteEnable") = 1
        Application(CookieName & "_SiteDisbleWhy") = ""
        Application.UnLock
        Response.Write "网站恢复正常访问..."
        Response.Write "</div>"
        Response.Write "<script>document.getElementById('accessSize').innerText='"&getFileInfo(AccessFile)(0)&"'</script>"
    End If
End If
'-------------备份数据库数据库-----------------
If Request.QueryString("do") = "Backup" Then
    Set AccessFSO = Server.CreateObject("Scripting.FileSystemObject")
    If AccessFSO.FileExists(Server.Mappath(AccessFile)) Then
        Response.Write "<div style='padding:4px 0px 4px 10px;border: 1px dotted #999;margin:2px;background:#ffffee'>"
        Response.Write "备份数据库开始，网站暂停一切用户的前台操作...<br/>"
        Response.Write "关闭数据库操作...<br/>"
        Call CloseDB
        Application.Lock
        FreeApplicationMemory
        Application(CookieName & "_SiteEnable") = 0
        Application(CookieName & "_SiteDisbleWhy") = "网站暂停中，请稍候几分钟后再来..."
        Application.UnLock
        CopyFiles Server.Mappath(AccessFile), Server.Mappath("backup/Backup_" & DateToStr(Now(), "YmdHIS") & "_" & randomStr(8) &".mbk")
        Application.Lock
        Application(CookieName & "_SiteEnable") = 1
        Application(CookieName & "_SiteDisbleWhy") = ""
        Application.UnLock
        Response.Write "压缩数据库完成...<br/>"
        Application.Lock
        Application(CookieName & "_SiteEnable") = 1
        Application(CookieName & "_SiteDisbleWhy") = ""
        Application.UnLock
        Response.Write "网站恢复正常访问..."
        Response.Write "</div>"
        Response.Write "<script>document.getElementById('accessSize').innerText='"&getFileInfo(AccessFile)(0)&"'</script>"
    End If
    Set AccessFSO = Nothing
End If
'---------------还原数据库------------
If Request.QueryString("do") = "Restore" Then
    AccessSource = Request.QueryString("source")
    Set AccessFSO = Server.CreateObject("Scripting.FileSystemObject")
    If AccessFSO.FileExists(Server.Mappath(AccessSource)) Then
        Response.Write "<div style='padding:4px 0px 4px 10px;border: 1px dotted #999;margin:2px;background:#ffffee'>"
        Response.Write "还原数据库开始，网站暂停一切用户的前台操作...<br/>"
        Response.Write "关闭数据库操作...<br/>"
        Call CloseDB
        Application.Lock
        FreeApplicationMemory
        Application(CookieName & "_SiteEnable") = 0
        Application(CookieName & "_SiteDisbleWhy") = "网站暂停中，请稍候几分钟后再来..."
        Application.UnLock
        CopyFiles Server.Mappath(AccessFile), Server.Mappath(AccessFile & ".TEMP")
        If DeleteFiles(Server.Mappath(AccessFile)) Then response.Write ("原数据库删除成功<br/>")
        response.Write CopyFiles(Server.Mappath(AccessSource), Server.Mappath(AccessFile))
        If DeleteFiles(Server.MapPath(AccessSource)) Then response.Write ("数据库备份删除成功<br/>")
        If DeleteFiles(Server.Mappath(AccessFile & ".TEMP")) Then response.Write ("Temp备份删除成功<br/>")
        Application.Lock
        Application(CookieName & "_SiteEnable") = 1
        Application(CookieName & "_SiteDisbleWhy") = ""
        Application.UnLock
        Response.Write "数据库还原完成...<br/>"
        Application.Lock
        Application(CookieName & "_SiteEnable") = 1
        Application(CookieName & "_SiteDisbleWhy") = ""
        Application.UnLock
        Response.Write "网站恢复正常访问..."
        Response.Write "</div>"
        Response.Write "<script>document.getElementById('accessSize').innerText='"&getFileInfo(AccessFile)(0)&"'</script>"
    End If
    Set AccessFSO = Nothing
End If
'---------------删除备份数据库------------
If Request.QueryString("do") = "DelFile" Then
    AccessSource = Request.QueryString("source")
    Set AccessFSO = Server.CreateObject("Scripting.FileSystemObject")
    If AccessFSO.FileExists(Server.Mappath(AccessSource)) Then
        Response.Write "<div style='padding:4px 0px 4px 10px;border: 1px dotted #999;margin:2px;background:#ffffee'>"
        If DeleteFiles(Server.MapPath(AccessSource)) Then response.Write ("数据库备份删除成功<br/>")
        Response.Write "</div>"
    End If
    Set AccessFSO = Nothing
End If

%>
	  	  <br/><b>数据库备份列表:</b> <br/>
	  <%
Dim BackUpFiles, BackUpFile
BackUpFiles = Split(getPathList("backup")(1), "*")
For Each BackUpFile in BackUpFiles
    response.Write "<a href=""backup/"&BackUpFile&""" target=""_blank"">"&getFileIcons(getFileInfo("backup/"&BackUpFile)(1))&BackUpFile&"</a>"
    response.Write "&nbsp;&nbsp;<a href=""?Fmenu=SQLFile&do=DelFile&source=backup/"&BackUpFile&""" title=""删除备份文件"">删除</a> | <a href=""?Fmenu=SQLFile&do=Restore&source=backup/"&BackUpFile&""" title=""删除备份文件"">还原数据库</a>"
    response.Write " | "&getFileInfo("backup/"&BackUpFile)(0)&" | "&getFileInfo("backup/"&BackUpFile)(2)&"<br/>"
Next

%>
	 <%end if%>
	 </div>
 </td>
  </tr></table>
<%
ElseIf Request.QueryString("Fmenu") = "Members" Then '帐户与权限
    Dim blog_Status, blog_Statu, StatusItem, blog_Status_Len
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
  <tr>
    <th class="CTitle"><%=categoryTitle%></th>
  </tr>
  <tr>
    <td class="CPanel">
	<div class="SubMenu"><a href="?Fmenu=Members">权限管理</a> | <a href="?Fmenu=Members&Smenu=Users">帐户管理</a></div>
<%
If Request.QueryString("Smenu") = "Users" Then

%>
<%getMsg%>
 <form action="ConContent.asp" method="post" style="margin:0px">
   <input type="hidden" name="action" value="Members"/>
   <input type="hidden" name="whatdo" value="SaveUserRight"/>
   <input type="hidden" name="DelID" value=""/>
<table border="0" cellpadding="2" cellspacing="1" class="TablePanel" style="margin:5px;">
<%
blog_Status = Application(CookieName&"_blog_rights")
Dim FindUser, FindUserFilter
FindUser = Request.QueryString("User")
If Len(FindUser)<1 Then
    FindUserFilter = ""
Else
    FindUserFilter = " AND M.mem_Name='" & FindUser & "'"
End If
If CheckStr(Request.QueryString("Page"))<>Empty Then
    Curpage = CheckStr(Request.QueryString("Page"))
    If IsInteger(Curpage) = False Or Curpage<0 Then Curpage = 1
Else
    Curpage = 1
End If
Dim bMember, PageCM
Set bMember = Server.CreateObject("ADODB.RecordSet")
SQL = "SELECT M.*,S.stat_name,S.stat_title FROM blog_Member as M,blog_status as S where M.mem_Status=S.stat_name"&FindUserFilter&" order by mem_RegTime desc"
bMember.Open SQL, Conn, 1, 1
PageCM = 0
response.Write ("<tr><td colspan=""8"" style=""border-bottom:1px solid #999;background:#fae1af;height:36px"">&nbsp;用户名&nbsp;<input id=""FindUser"" type=""text"" class=""text"" size=""16""/><input type=""button"" value=""查找用户"" class=""button"" style=""margin-bottom:-2px"" onclick=""location='ConContent.asp?Fmenu=Members&Smenu=Users&User='+escape(document.getElementById('FindUser').value)""/></td></tr>")
If Not bMember.EOF Then
    bMember.PageSize = 30
    bMember.AbsolutePage = CurPage
    Dim bMember_nums
    bMember_nums = bMember.RecordCount
    response.Write "<tr><td colspan=""8"" style=""border-bottom:1px solid #999;""><div class=""pageContent"">"&MultiPage(bMember_nums, 30, CurPage, "?Fmenu=Members&Smenu=Users&", "", "float:left","")&"</div><div class=""Content-body"" style=""line-height:200%""></td></tr>"

%>
        <tr align="center">
          <td nowrap="nowrap" class="TDHead">编号</td>
          <td width="100" nowrap="nowrap" class="TDHead">会员名称</td>
          <td width="100" class="TDHead">会员身份</td>
          <td class="TDHead">注册时间</td>
          <td class="TDHead">上次访问时间</td>
          <td class="TDHead">最后登录IP地址</td>
          <td class="TDHead">设置权限</td>
          <td class="TDHead">&nbsp;</td>
	   </tr>
	   <%
blog_Status_Len = UBound(blog_Status, 2)
Do Until bMember.EOF Or PageCM = bMember.PageSize

%>
        <tr align="center">
          <td nowrap><%=bMember("mem_ID")%>
          <%
If bMember("mem_Name") <> memName Then
    response.Write "<input type=""hidden"" name=""mem_ID"" value="""&bMember("mem_ID")&"""/>"
End If

%>
		</td>
          <td nowrap align="left"><a href="member.asp?action=view&memName=<%=Server.URLEncode(bMember("mem_Name"))%>" target="_blank"><%=bMember("mem_Name")%></a></td>
          <td nowrap>&nbsp;<span id="RightStr_<%=bMember("mem_ID")%>"><%=bMember("stat_title")%></span>&nbsp;</td>
          <td nowrap>&nbsp;<%=DateToStr(bMember("mem_RegTime"),"Y-m-d")%>&nbsp;</td>
          <td nowrap>&nbsp;<%=DateToStr(bMember("mem_lastVisit"),"Y-m-d H:I A")%>&nbsp;</td>
          <td nowrap>&nbsp;<%=bMember("mem_lastIP")%>&nbsp;</td>
          <td>
	          <select name="mem_Status" onChange="ChValue(this.value,'RightStr_<%=bMember("mem_ID")%>')" <%if bMember("mem_Name") = memName then response.write "disabled"%>><%
For i = 0 To blog_Status_Len
    If bMember("stat_name") = blog_Status(0, i) Then
        response.Write "<option value="""&blog_Status(0, i)&""" selected=""selected"">"&blog_Status(0, i)&"</option>"
    Else
        response.Write "<option value="""&blog_Status(0, i)&""">"&blog_Status(0, i)&"</option>"
    End If
Next

%></select>
          </td>
          <td>
           <%if bMember("mem_Name") <> memName then%>
          	<a href="javascript:delUser(<%=bMember("mem_ID")%>)"><img border="0" src="images/icon_del.gif" width="16" height="16" /></a>
          <%end if%>
          </td>
	   </tr>
   <%
bMember.MoveNext
PageCM = PageCM + 1
Loop
bMember.Close
Set bMember = Nothing
Else
    response.Write ("<tr><td colspan=""8"" align=""center"" >你所查询的用户不存在！</td></tr>")
End If

%></table>
 	<div class="SubButton">
      <input type="submit" name="Submit" value="保存用户权限" class="button"/> 
     </div>
     </form>
 <script>
  function ChValue(str,obj){
   <%
For i = 0 To blog_Status_Len
    response.Write "if (str=='"&blog_Status(0, i)&"') {document.getElementById(obj).innerText='"&blog_Status(1, i)&"'};"
Next

%>
   }
 </script>
 <%
ElseIf Request.QueryString("Smenu") = "EditRight" Then
    Dim RightDB
    sql = "select * from blog_status where stat_name='" & checkstr(Request.QueryString("id")) & "'"
    Set RightDB = conn.Execute(sql)
    If RightDB.EOF Then
        response.Write "没找到该权限组.请重新更新blog缓存信息"
    Else

%>
	    <form action="ConContent.asp" method="post" style="margin:0px">
	     <input type="hidden" name="action" value="Members"/>
	     <input type="hidden" name="whatdo" value="EditGroup"/>
	     <input type="hidden" name="status_name" value="<%=checkstr(Request.QueryString("id"))%>"/>
	     <div align="left" style="padding:5px;"><%getMsg%>
	     <table border="0" cellpadding="3" cellspacing="1" class="TablePanel" style="margin:6px">
	     <tr><td colspan="2" align="left" style="background:#e5e5e5;padding:6px">
	     <div style="font-weight:bold;font-size:14px;"><%=RightDB("stat_name")%> 权限设置</div></td></tr>
		     <tr><td align="right" width="100">权限名称</td><td width="300"><input name="status_title" type="text" size="20" class="text" value="<%=RightDB("stat_title")%>"/></td></tr>
		     <tr><td align="right">添加日志</td>
		         <td><select name="AddArticle">
	                  <option value="11" style="background:#C5FDB7">允许</option>
	                  <option value="00" style="background:#FABABA" <%if not CBool(mid(RightDB("stat_code"),2,1)) then response.write ("selected=""selected""") %>>不允许</option>
	                 </select>
	            </td></tr>
	          
	         <tr><td align="right">编辑日志</td>
	             <td><select name="EditArticle">
					  <option value="11" style="background:#C5FDB7">所有</option>
					  <option value="01" style="background:#B7D8FD" <%if not CBool(mid(RightDB("stat_code"),3,1)) and CBool(mid(RightDB("stat_code"),4,1)) then response.write ("selected=""selected""")%>>自己</option>
					  <option value="00" style="background:#FABABA" <%if not CBool(mid(RightDB("stat_code"),3,1)) and not CBool(mid(RightDB("stat_code"),4,1)) then response.write ("selected=""selected""")%>>不允许</option>
                     </select>
                 </td></tr>

	         <tr><td align="right">删除日志</td>
	             <td><select name="DelArticle">
			            <option value="11" style="background:#C5FDB7">所有</option>
			            <option value="01" style="background:#B7D8FD" <%if not CBool(mid(RightDB("stat_code"),5,1)) and CBool(mid(RightDB("stat_code"),6,1)) then response.write ("selected=""selected""")%>>自己</option>
			            <option value="00" style="background:#FABABA" <%if not CBool(mid(RightDB("stat_code"),5,1)) and not CBool(mid(RightDB("stat_code"),6,1)) then response.write ("selected=""selected""")%>>不允许</option>
			        </select>
			    </td></tr>
	         <tr><td align="right">发表评论</td>
	            <td ><select name="AddComment">
			            <option value="11" style="background:#C5FDB7">允许</option>
			            <option value="00" style="background:#FABABA" <%if not CBool(mid(RightDB("stat_code"),7,1)) then response.write ("selected=""selected""")%>>不允许</option>
                    </select>
                </td></tr>
	         <tr><td align="right">删除评论</td>
		          <td ><select name="DelComment">
		            <option value="1" style="background:#C5FDB7">允许</option>
		            <option value="0" style="background:#FABABA" <%if not CBool(mid(RightDB("stat_code"),9,1)) then response.write ("selected=""selected""")%>>不允许</option>
		          </select>
		        </td></tr>
	          <tr><td align="right">允许查看隐藏分类</td>
	            <td ><select name="ShowHiddenCate">
			            <option value="1" style="background:#C5FDB7">允许</option>
			            <option value="0" style="background:#FABABA" <%if not CBool(mid(RightDB("stat_code"),12,1)) then response.write ("selected=""selected""")%>>不允许</option>
                    </select>
                </td></tr>
               <tr><td align="right">管理员</td>
		         <td ><select name="IsAdmin">
		            <option value="1" style="background:#C5FDB7">是</option>
		            <option value="0" style="background:#FABABA" <%if not CBool(mid(RightDB("stat_code"),11,1)) then response.write ("selected=""selected""")%>>否</option>
		          </select>
		        </td></tr> 
	         <tr><td align="right">上传附件</td>
		          <td ><select name="CanUpload">
		            <option value="1" style="background:#C5FDB7">允许</option>
		            <option value="0" style="background:#FABABA" <%if not CBool(mid(RightDB("stat_code"),10,1)) then response.write ("selected=""selected""")%>>不允许</option>
		          </select>
		      </td></tr>
	         <tr><td align="right">附件大小</td><td ><input name="UploadSize" type="text" size="20" class="text" title="<%=RightDB("stat_attSize")%>字节" value="<%=RightDB("stat_attSize")%>" style="font-size:11px" onChange="this.title=this.value+' 字节'"/></td></tr>
             <tr><td align="right">附件类型</td><td ><input name="UploadType" type="text" size="50" class="text" title="<%=RightDB("stat_attType")%>" value="<%=RightDB("stat_attType")%>" style="font-size:11px" onChange="this.title=this.value"/></td></tr>
             <tr><td colspan="2" align="center"><input type="submit" name="Submit" value="保存设置" class="button"/><input type="button" value="放弃返回" class="button" onClick="history.go(-1)"/></td></tr>
	         
	     </table>
	     </div>
	    </form>
<%
End If
Else
%>
   <form action="ConContent.asp" method="post" style="margin:0px">
   <input type="hidden" name="action" value="Members"/>
   <input type="hidden" name="whatdo" value="Group"/>
   <input type="hidden" name="DelGroup" value=""/>
    <div align="left" style="padding:5px;"><%getMsg%>
	      <table border="0" cellpadding="2" cellspacing="1" class="TablePanel">
        <tr align="center">
          <td nowrap="nowrap" class="TDHead">权限标识</td>
          <td nowrap="nowrap" class="TDHead">权限标题</td>
          <td width="16" nowrap="nowrap" class="TDHead">&nbsp;</td>
	   </tr>
<%
blog_Status = Application(CookieName&"_blog_rights")
blog_Status_Len = UBound(blog_Status, 2)

For i = 0 To blog_Status_Len
%>
        <tr align="center">
          <td ><input name="status_name" type="text" size="15" class="text" value="<%=blog_Status(0,i)%>"  readonly="readonly" style="background:#ffe;font-size:11px"/></td>
          <td ><input name="status_title" type="text" size="20" class="text" value="<%=blog_Status(1,i)%>"/></td>
          <td align="left">
          <a href="?Fmenu=Members&Smenu=EditRight&id=<%=blog_Status(0,i)%>" title="编辑该权限组"><img border="0" src="images/icon_edit.gif" width="16" height="16" style="margin:0px 2px -3px 0px"/>编辑权限</a>
		  <%if lcase(blog_Status(0,i))<>"supadmin" and lcase(blog_Status(0,i))<>"member" and lcase(blog_Status(0,i))<>"guest" then%>
		  <a href="javascript:CheckDelGroup('<%=blog_Status(0,i)%>')" title="删除该权限组"><img border="0" src="images/icon_del.gif" width="16" height="16" style="margin:0px 2px -3px 0px"/>删除权限</a>
		  <%end if%>
		  </td>
	   </tr>
<%next%>
        <tr align="center" bgcolor="#D5DAE0">
         <td colspan="12" class="TDHead" align="left" style="border-top:1px solid #999"><img src="images/add.gif" style="margin:0px 2px -3px 2px"/>添加新权限分组</td>
        </tr>
        <tr align="center">
          <td ><input name="status_name" type="text" size="15" class="text" style="font-size:11px"/></td>
          <td ><input name="status_title" type="text" size="20" class="text"/></td>
          <td >&nbsp;</td>
	   </tr>
	 </table>
  <div style="color:#f00">权限标识是唯一标记.一旦确定就无法修改.系统自带的权限组不允许删除.</div>
	<div class="SubButton">
      <input type="submit" name="Submit" value="保存权限组" class="button"/> 
     </div>	  
	 </div></form>
 </td>
  </tr></table>
<%
End If
ElseIf Request.QueryString("Fmenu") = "Link" Then '友情链接管理
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
  <tr>
	<th class="CTitle"><%=categoryTitle%></th>
  </tr>
  <tr>
    <td class="CPanel"><form name="filter" action="ConContent.asp" method="post" style="margin:0px">
	<%
Dim disLink, disCount
If session(CookieName&"_disLink") = "" Then session(CookieName&"_disLink") = "All"
If session(CookieName&"_disCount") = "" Then session(CookieName&"_disCount") = "10"
disLink = session(CookieName&"_disLink")
disCount = session(CookieName&"_disCount")
Dim FilterWhere
If CheckStr(Request.QueryString("Page"))<>Empty Then
    Curpage = CheckStr(Request.QueryString("Page"))
    If IsInteger(Curpage) = False Or Curpage<0 Then Curpage = 1
Else
    Curpage = 1
End If

%>
	<div class="SubMenu">过滤器: 
   <input type="hidden" name="action" value="Links"/>
   <input type="hidden" name="whatdo" value="Filter"/>
  	<select name="disLink" onChange="document.forms['filter'].submit()">
	  <option value="All">显示所有链接</option>
	  <option value="Allow" <%if disLink="Allow" then response.write ("selected=""selected""")%>>已通过验证链接</option>
	  <option value="NoAllow" <%if disLink="NoAllow" then response.write ("selected=""selected""")%>>未通过验证链接</option>
	  <option value="Top" <%if disLink="Top" then response.write ("selected=""selected""")%>>置顶链接</option>
	  <option value="NoTop" <%if disLink="NoTop" then response.write ("selected=""selected""")%>>未置顶链接</option>
	</select> 每页显示 	<select name="disCount" onChange="document.forms['filter'].submit()">
	  <option <%if int(disCount)=100 then response.write ("selected=""selected""")%>>100</option>
	  <option <%if int(disCount)=50 then response.write ("selected=""selected""")%>>50</option>
	  <option <%if int(disCount)=25 then response.write ("selected=""selected""")%>>25 </option>
	  <option <%if int(disCount)=10 then response.write ("selected=""selected""")%>>10</option>
	</select>
	条 <a href="#AddLink">添加新友情链接</a></div> </form>
   <form name="Link" action="ConContent.asp" method="post" style="margin:0px">
    <div align="left" style="padding:5px;"><%getMsg%>
       <input type="hidden" name="action" value="Links"/>
       <input type="hidden" name="whatdo" value="SaveLink"/>
       <input type="hidden" name="ALinkID" value=""/>
       <input type="hidden" name="Page" value="<%=Curpage%>"/>       
	   <table border="0" cellpadding="2" cellspacing="1" class="TablePanel">
<%
Select Case disLink
    Case "All"
        FilterWhere = ""
    Case "Allow"
        FilterWhere = " where link_IsShow=true"
    Case "NoAllow"
        FilterWhere = " where link_IsShow=false"
    Case "Top"
        FilterWhere = " where link_IsMain=true"
    Case "NoTop"
        FilterWhere = " where link_IsMain=false"
    Case Else
        FilterWhere = ""
End Select
Dim bLink
Set bLink = Server.CreateObject("ADODB.RecordSet")
SQL = "SELECT * FROM blog_Links"&FilterWhere&" order by link_IsShow desc,link_Order desc"
bLink.Open SQL, Conn, 1, 1
If Not bLink.EOF Then
    bLink.PageSize = disCount
    bLink.AbsolutePage = CurPage
    Dim bLink_nums
    bLink_nums = bLink.RecordCount
    Dim MultiPages, PageCount
    response.Write "<tr><td colspan=""6"" style=""border-bottom:1px solid #999;""><div class=""pageContent"">"&MultiPage(bLink_nums, disCount, CurPage, "?Fmenu=Link&Smenu=&", "", "float:left","")&"</div><div class=""Content-body"" style=""line-height:200%""></td></tr>"
End If
%>
        <tr align="center">
          <td width="16" nowrap="nowrap" class="TDHead">&nbsp;</td>
          <td width="120" nowrap="nowrap" class="TDHead">网站名称</td>
          <td width="180" nowrap="nowrap" class="TDHead">网站地址</td>
          <td width="250" nowrap="nowrap" class="TDHead">Logo图片地址</td>
          <td class="TDHead">排序</td>
          <td class="TDHead">操作</td>
	   </tr>
	   <%
If Not bLink.EOF Then
    Do Until bLink.EOF Or PageCount = bLink.PageSize
        If Not bLink("link_IsShow") Then

%>
        <tr align="center" bgcolor="#FCF4BC">
          <td><img src="images/slink.gif" alt="没有通过验证链接"/></td>
          <td><input name="LinkID" type="hidden" value="<%=bLink("link_ID")%>"/><input name="LinkName" type="text" size="18" class="text" value="<%=bLink("link_Name")%>"/></td>
          <td><input name="LinkURL" type="text" size="30" class="text" value="<%=bLink("link_URL")%>"/></td>
          <td><input name="LinkLogo" type="text" size="40" class="text" value="<%=bLink("link_Image")%>"/></td>
          <td><input name="LinkOrder" type="text" class="text" size="2" value="<%=bLink("link_Order")%>"/></td>
          <td><a href="#" onClick="ShowLink(<%=bLink("link_ID")%>)" title="通过该链接的验证"><img border="0" src="images/alink.gif" width="16" height="16" style="margin:0px 2px -3px 0px"/>通过</a> <a href="<%=bLink("link_URL")%>" target="_blank" title="查看该链接"><img border="0" src="images/icon_trackback.gif" width="16" height="16" style="margin:0px 2px -3px 0px"/>查看</a> <a href="#" onClick="Dellink(<%=bLink("link_ID")%>)" title="删除该链接"><img border="0" src="images/icon_del.gif" width="16" height="16" style="margin:0px 2px -3px 0px"/>删除</a> </td>
        </tr>
		<%else%>
        <tr align="center">
          <td><%if bLink("link_IsMain") then response.write ("<img src=""images/urlInTop.gif"" alt=""置顶链接""/>") else response.write ("&nbsp;")%></td>
          <td><input name="LinkID" type="hidden" value="<%=bLink("link_ID")%>"/><input name="LinkName" type="text" size="18" class="text" value="<%=bLink("link_Name")%>"/></td>
          <td><input name="LinkURL" type="text" size="30" class="text" value="<%=bLink("link_URL")%>"/></td>
          <td><input name="LinkLogo" type="text" size="40" class="text" value="<%=bLink("link_Image")%>"/></td>
          <td><input name="LinkOrder" type="text" class="text" size="2" value="<%=bLink("link_Order")%>"/></td>
          <td><%if bLink("link_IsMain") then response.write ("<a href=""#"" onclick=""Toplink("&bLink("link_ID")&")"" title=""取消该链接在首页置顶""><img border=""0"" src=""images/ct.gif"" style=""margin:0px 2px -3px 0px""/>取消</a> ") else response.write ("<a href=""#"" onclick=""Toplink("&bLink("link_ID")&")"" title=""把该链接在首页置顶""><img border=""0"" src=""images/it.gif"" style=""margin:0px 2px -3px 0px""/>置顶</a> ")%>
		  <a href="<%=bLink("link_URL")%>" target="_blank" title="查看该链接"><img border="0" src="images/icon_trackback.gif" width="16" height="16" style="margin:0px 2px -3px 0px"/>查看</a> <a href="#" onClick="Dellink(<%=bLink("link_ID")%>)" title="删除该链接"><img border="0" src="images/icon_del.gif" width="16" height="16" style="margin:0px 2px -3px 0px"/>删除</a> </td>
	   </tr>
<%End If
bLink.MoveNext
PageCount = PageCount + 1
Loop
bLink.Close
Set bLink = Nothing
End If
%>
        <tr align="center" bgcolor="#D5DAE0">
         <td colspan="6" class="TDHead" align="left" style="border-top:1px solid #999"><a name="AddLink"></a><img src="images/add.gif" style="margin:0px 2px -3px 2px"/>添加新友情链接</td>
        </tr>	
        <tr align="center">
          <td>&nbsp;</td>
          <td><input name="new_LinkID" type="hidden" value="-1"/><input name="new_LinkName" type="text" size="18" class="text"/></td>
          <td><input name="new_LinkURL" type="text" size="30" class="text" /></td>
          <td><input name="new_LinkLogo" type="text" size="40" class="text" /></td>
          <td><input name="new_LinkOrder" type="text" class="text" size="2" /></td>
          <td>&nbsp;</td>
	   </tr>
	 	</table>
  </div>
  <div class="SubButton">
      <input type="submit" name="Submit" value="保存友情链接" class="button"/> 
     </div>	  
 </td>
  </tr></table></div></form>
<%
ElseIf Request.QueryString("Fmenu") = "smilies" Then '表情与关键字
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
  <tr>
	<th class="CTitle"><%=categoryTitle%></th>
  </tr>
  <tr>
    <td class="CPanel">
	<div class="SubMenu"><a href="?Fmenu=smilies">表情管理</a> | <a href="?Fmenu=smilies&Smenu=KeyWord">关键字管理</a></div>
    <div align="left" style="padding:5px;"><%getMsg%>
     <%if Request.QueryString("Smenu")="KeyWord" then%>
   <form action="ConContent.asp" method="post" style="margin:0px">
   <input type="hidden" name="action" value="smilies"/>
   <input type="hidden" name="whatdo" value="KeyWord"/>
   <input type="hidden" name="DelID" value=""/>
  	   <table border="0" cellpadding="2" cellspacing="1" class="TablePanel">
        <tr align="center">
		  <td width="16" nowrap="nowrap" class="TDHead">&nbsp;</td>
          <td width="120" nowrap="nowrap" class="TDHead">关键字</td>
          <td width="220" nowrap="nowrap" class="TDHead">关键字链接地址</td>
	   </tr>
	   <%
Dim bKeyWord
Set bKeyWord = conn.Execute("select * from blog_Keywords order by key_ID desc")
Do Until bKeyWord.EOF

%>
        <tr align="center">
          <td><input name="SelectKeyWordID" type="checkbox" value="<%=bKeyWord("key_ID")%>"/></td>
          <td><input name="KeyWordID" type="hidden" value="<%=bKeyWord("key_ID")%>"/><input name="KeyWord" type="text" size="18" class="text" value="<%=bKeyWord("key_Text")%>"/></td>
          <td><input name="KeyWordURL" type="text" size="34" class="text" value="<%=bKeyWord("key_URL")%>"/></td>
	   </tr>
	   <%
bKeyWord.movenext
Loop

%>
       <tr align="center" bgcolor="#D5DAE0">
        <td colspan="3" class="TDHead" align="left" style="border-top:1px solid #999"><a name="AddLink"></a><img src="images/add.gif" style="margin:0px 2px -3px 2px"/>添加新关键字</td>
       </tr>	
        <tr align="center">
		  <td></td>
          <td><input name="KeyWordID" type="hidden" value="-1"/><input name="KeyWord" type="text" size="18" class="text"/></td>
          <td><input name="KeyWordURL" type="text" size="34" class="text"/></td>
	   </tr>
	 </table>
  <div class="SubButton" style="text-align:left;">
      	 <select name="doModule">
			 <option value="SaveAll">保存所有关键字</option>
			 <option value="DelSelect">删除所选关键字</option>
		 </select>
      <input type="submit" name="Submit" value="保存关键字" class="button" style="margin-bottom:0px"/> 
     </div>	  </form>
     <%else%>
   <form action="ConContent.asp" method="post" style="margin:0px">
   <input type="hidden" name="action" value="smilies"/>
   <input type="hidden" name="whatdo" value="smilies"/>
   <input type="hidden" name="DelID" value=""/>
	   <table border="0" cellpadding="2" cellspacing="1" class="TablePanel">
        <tr align="center">
		  <td width="16" nowrap="nowrap" class="TDHead">&nbsp;</td>
          <td nowrap="nowrap" class="TDHead">图片</td>
          <td width="100" nowrap="nowrap" class="TDHead">表情图片代码</td>
          <td width="180" nowrap="nowrap" class="TDHead">表情图片地址</td>
	   </tr>
	   <%
Dim bSmile
Set bSmile = conn.Execute("select * from blog_Smilies order by sm_ID desc")
Do Until bSmile.EOF

%>
	   <tr align="center">
		  <td><input name="selectSmiliesID" type="checkbox" value="<%=bSmile("sm_ID")%>"/></td>
          <td><img src="images/smilies/<%=bSmile("sm_Image")%>" alt="<%=bSmile("sm_Image")%>"/></td>
          <td><input name="smilesID" type="hidden" value="<%=bSmile("sm_ID")%>"/><input name="smiles" type="text" size="14" class="text" value="<%=bSmile("sm_Text")%>"/></td>
          <td><input name="smilesURL" type="text" size="27" class="text" value="<%=bSmile("sm_Image")%>"/></td>
	   </tr>
	   <%
bSmile.movenext
Loop

%>
       <tr align="center" bgcolor="#D5DAE0">
        <td colspan="4" class="TDHead" align="left" style="border-top:1px solid #999"><a name="AddLink"></a><img src="images/add.gif" style="margin:0px 2px -3px 2px"/>添加新表情</td>
       </tr>	
        <tr align="center">
		  <td></td>
          <td></td>
          <td><input name="smilesID" type="hidden" value="-1"/><input name="smiles" type="text" size="14" class="text"/></td>
          <td><input name="smilesURL" type="text" size="27" class="text"/></td>
	   </tr>
	  </table>
  <div class="SubButton" style="text-align:left;">
    	 <select name="doModule">
			 <option value="SaveAll">保存所有表情</option>
			 <option value="DelSelect">删除所选表情</option>
		 </select>
	  <input type="submit" name="Submit" value="保存表情" class="button" style="margin-bottom:0px"/> 
     </div>	</form>
		<%end if%>
    </div>
</td></tr></table>
<%
ElseIf Request.QueryString("Fmenu") = "Status" Then '服务器配置
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
  <tr>
	<th class="CTitle"><%=categoryTitle%></th>
  </tr>
  <tr>
    <td class="CPanel">
    <div align="left" style="padding:5px;line-height:150%">
     <b>软件版本:</b> PJBlog3 v<%=blog_version%> - <%=DateToStr(blog_UpdateDate,"mdy")%><br/>
     <b>服务器时间:</b> <%=DateToStr(Now(),"Y-m-d H:I A")%><br/>
     <b>服务器物理路径:</b> <%=Request.ServerVariables("APPL_PHYSICAL_PATH")%><br/>
     <b>服务器空间占用:</b> <%=GetTotalSize(Request.ServerVariables("APPL_PHYSICAL_PATH"),"Folder")%><br/>
     <b>服务器CPU数量:</b> <%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%><br/>
     <b>服务器IIS版本:</b> <%=Request.ServerVariables("SERVER_SOFTWARE")%><br/>
     <b>脚本超时设置:</b> <%=Server.ScriptTimeout%><br/>
     <b>脚本解释引擎:</b> <%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %><br/>
     <b>服务器操作系统:</b> <%=Request.ServerVariables("OS")%><br/>
     <b>服务器IP地址:</b> <%=Request.ServerVariables("LOCAL_ADDR")%><br/>
     <b>客户端IP地址:</b> <%=Request.ServerVariables("REMOTE_ADDR")%><br/><br/>
     
     <b>关键组件:</b> (缺少关键组件的服务器会对PJBlog3运行有一定影响)<br/>
     <b>　－ Scripting.FileSystemObject 组件:</b> <%=DisI(CheckObjInstalled("Scripting.FileSystemObject"))%><br/>
     <b>　－ MSXML2.ServerXMLHTTP 组件:</b> <%=DisI(CheckObjInstalled("MSXML2.ServerXMLHTTP"))%><br/>
     <b>　－ Microsoft.XMLDOM 组件:</b> <%=DisI(CheckObjInstalled("Microsoft.XMLDOM"))%><br/>
     <b>　－ ADODB.Stream 组件:</b> <%=DisI(CheckObjInstalled("ADODB.Stream"))%><br/>
     <b>　－ Scripting.Dictionary 组件:</b> <%=DisI(CheckObjInstalled("Scripting.Dictionary"))%><br/>
     <br/>
     
     <b>其他组件: </b>(以下组件不影响PJBlog3运行)<br/>
     <b>　－ Msxml2.ServerXMLHTTP.5.0 组件:</b> <%=DisI(CheckObjInstalled("Msxml2.ServerXMLHTTP.5.0"))%><br/>
     <b>　－ Msxml2.DOMDocument.5.0 组件:</b> <%=DisI(CheckObjInstalled("Msxml2.DOMDocument.5.0"))%><br/>
     <b>　－ FileUp.upload 组件:</b> <%=DisI(CheckObjInstalled("FileUp.upload"))%><br/>
     <b>　－ JMail.SMTPMail 组件:</b> <%=DisI(CheckObjInstalled("JMail.SMTPMail"))%><br/>
     <b>　－ GflAx190.GflAx 组件:</b> <%=DisI(CheckObjInstalled("GflAx190.GflAx"))%><br/>
     <b>　－ easymail.Mailsend 组件:</b> <%=DisI(CheckObjInstalled("easymail.Mailsend"))%><br/>
    </div>
</td></tr></table>
<%

ElseIf Request.QueryString("Fmenu") = "Logout" Then '退出
    session(CookieName&"_System") = ""
    session(CookieName&"_disLink") = ""
    session(CookieName&"_disCount") = ""

%>
  <script>try{top.location="default.asp"}catch(e){location="default.asp"}</script>
  <%
ElseIf Request.QueryString("Fmenu") = "welcome" Then '欢迎
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
  <tr>
 	<th class="CTitle"><%=categoryTitle%></th>
  </tr>
  <tr>
    <td class="CPanel">
        <div id="updateInfo" style="background:ffffe1;border:1px solid #89441f;padding:4px;display:none"></div>
    <script>
      var CVersion="<%=blog_version%>";
      var CDate="<%=blog_UpdateDate%>";
    </script>
    <script src="http://www.pjhome.net/updateN.js"></script>
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
	 <tr>
		 <td valign="top" style="padding:5px;width:140px"><img src="images/Control/Icon/ControlPanel.png"/></td>
		 <td valign="top">
		    <div align="left" style="padding:5px;line-height:170%;clear:both;font-size:12px">
		     <b>当前软件版本:</b> PJBlog3 v<%=blog_version%><br/>
		     <b>软件更新日期:</b> <%=DateToStr(blog_UpdateDate,"mdy")%><br/>
		     <b>日志数量:</b> <%=blog_LogNums%> 篇<br/>
		     <b>评论数量:</b> <%=blog_CommNums%> 个<br/>
		     <b>引用数量:</b> <%=blog_TbCount%> 个<br/>
		     <b>留言数量:</b> <%=blog_MessageNums%>  个(需要留言本插件支持)<br/>
		     <b>会员数:</b> <%=blog_MemNums%> 人<br/>
		     <b>访问次数:</b> <%=blog_VisitNums%> 次<br/>
		    </div>    
		 </td>
	 </tr>
	</table>

</td></tr></table>
<%Else

    '==========================后台信息处理===============================
    Dim weblog
    Set weblog = Server.CreateObject("ADODB.RecordSet")
    '==========================基本信息处理===============================
    If Request.Form("action") = "General" Then
        If Request.Form("whatdo") = "General" Then
            '--------------------------基本信息处理-------------------
            SQL = "SELECT * FROM blog_Info"
            weblog.Open SQL, Conn, 1, 3
            weblog("blog_Name") = checkURL(CheckStr(Request.Form("SiteName")))
            weblog("blog_Title") = checkURL(CheckStr(Request.Form("blog_Title")))
            weblog("blog_master") = checkURL(CheckStr(Request.Form("blog_master")))
            weblog("blog_email") = checkURL(CheckStr(Request.Form("blog_email")))

            If Right(CheckStr(Request.Form("SiteURL")), 1)<>"/" Then
                weblog("blog_URL") = checkURL(CheckStr(Request.Form("SiteURL")))&"/"
            Else
                weblog("blog_URL") = checkURL(CheckStr(Request.Form("SiteURL")))
            End If

            weblog("blog_affiche") = ""
            weblog("blog_about") = CheckStr(Request.Form("blog_about"))
            weblog("blog_PerPage") = CheckStr(Request.Form("blogPerPage"))
            weblog("blog_commPage") = CheckStr(Request.Form("blogcommpage"))
            weblog("blog_BookPage") = 0 'CheckStr(Request.form("blogBookPage"))
            weblog("blog_commTimerout") = CheckStr(Request.Form("blog_commTimerout"))
            weblog("blog_commLength") = CheckStr(Request.Form("blog_commLength"))
            If CheckObjInstalled("ADODB.Stream") Then
                weblog("blog_postFile") = Request.Form("blog_postFile")
                ' if CheckStr(Request.form("blog_postFile"))="1" then weblog("blog_postFile")=1 else weblog("blog_postFile")=0
            Else
                weblog("blog_postFile") = 0
            End If
            If CheckStr(Request.Form("blog_Disregister")) = "1" Then weblog("blog_Disregister") = 1 Else weblog("blog_Disregister") = 0
            If CheckStr(Request.Form("blog_validate")) = "1" Then weblog("blog_validate") = 1 Else weblog("blog_validate") = 0
            If CheckStr(Request.Form("blog_commUBB")) = "1" Then weblog("blog_commUBB") = 1 Else weblog("blog_commUBB") = 0
            If CheckStr(Request.Form("blog_commIMG")) = "1" Then weblog("blog_commIMG") = 1 Else weblog("blog_commIMG") = 0
            If CheckStr(Request.Form("blog_ImgLink")) = "1" Then weblog("blog_ImgLink") = 1 Else weblog("blog_ImgLink") = 0
            weblog("blog_SplitType") = CBool(CheckStr(Request.Form("blog_SplitType")))
            weblog("blog_introChar") = CheckStr(Request.Form("blog_introChar"))
            weblog("blog_introLine") = CheckStr(Request.Form("blog_introLine"))
            weblog("blog_FilterName") = CheckStr(Request.Form("Register_UserNames"))
            weblog("blog_FilterIP") = CheckStr(Request.Form("FilterIPs"))
            weblog("blog_DisMod") = CheckStr(Request.Form("blog_DisMod"))
            If Not IsInteger(Request.Form("blog_CountNum")) Then
                weblog("blog_CountNum") = 0
            Else
                weblog("blog_CountNum") = Request.Form("blog_CountNum")
            End If
            weblog("blog_wapNum") = CheckStr(Request.Form("blog_wapNum"))
            If CheckStr(Request.Form("blog_wapImg")) = "1" Then weblog("blog_wapImg") = 1 Else weblog("blog_wapImg") = 0
            If CheckStr(Request.Form("blog_wapHTML")) = "1" Then weblog("blog_wapHTML") = 1 Else weblog("blog_wapHTML") = 0
            If CheckStr(Request.Form("blog_wapLogin")) = "1" Then weblog("blog_wapLogin") = 1 Else weblog("blog_wapLogin") = 0
            If CheckStr(Request.Form("blog_wapComment")) = "1" Then weblog("blog_wapComment") = 1 Else weblog("blog_wapComment") = 0
            If CheckStr(Request.Form("blog_wap")) = "1" Then weblog("blog_wap") = 1 Else weblog("blog_wap") = 0
            If CheckStr(Request.Form("blog_wapURL")) = "1" Then weblog("blog_wapURL") = 1 Else weblog("blog_wapURL") = 0

            Response.Cookies(CookieNameSetting)("ViewType") = ""
            weblog.update
            weblog.Close
            getInfo(2)
            If Int(Request.Form("SiteOpen")) = 1 Then
                Application.Lock
                Application(CookieName & "_SiteEnable") = 1
                Application(CookieName & "_SiteDisbleWhy") = ""
                Application.UnLock
            Else
                Application.Lock
                Application(CookieName & "_SiteEnable") = 0
                Application(CookieName & "_SiteDisbleWhy") = "抱歉!网站暂时关闭!"
                Application.UnLock
            End If
            session(CookieName&"_ShowMsg") = True
            session(CookieName&"_MsgText") = "基本信息修改成功!"
            RedirectUrl("ConContent.asp?Fmenu=General&Smenu=")
        ElseIf Request.Form("whatdo") = "Misc" Then
%>
    <!--#include file="common/ubbcode.asp" -->
    <%
If Request.Form("ReBulidArticle") = 1 or Request.Form("ReBulidArticle") = 2 Then           
    Application.Lock
    Application(CookieName & "_SiteEnable") = 0
    Application(CookieName & "_SiteDisbleWhy") = "抱歉!网站在初始化数据，请稍后在访问. :P"
    Application.UnLock
    
    Dim LoadArticle, LogLen
    LogLen = 0
    Set LoadArticle = conn.Execute("SELECT log_ID FROM blog_Content")
    Do Until LoadArticle.EOF
    	if Request.Form("ReBulidArticle") = 2 then 
     	   PostArticle LoadArticle("log_ID"), True
    	else
      	  PostArticle LoadArticle("log_ID"), False
    	end if
        LogLen = LogLen + 1
        LoadArticle.movenext
    Loop
    
    Application.Lock
    Application(CookieName & "_SiteEnable") = 1
    Application(CookieName & "_SiteDisbleWhy") = ""
    Application.UnLock
    
    session(CookieName&"_ShowMsg") = True
    session(CookieName&"_MsgText") = Session(CookieName&"_MsgText")&"共处理了 "&LogLen&" 篇日志文件! "
End If

If Request.Form("ReBulidIndex") = 1 Then
    Dim lArticle
    Set lArticle = New ArticleCache
    lArticle.SaveCache
    Set lArticle = Nothing
    session(CookieName&"_ShowMsg") = True
    session(CookieName&"_MsgText") = Session(CookieName&"_MsgText")&"重新输出日志索引! "
End If

If Request.Form("ReTatol") = 1 Then
    Dim blog_Content_count, blog_Comment_count, ContentCount, TBCount, Count_Member
    ContentCount = 0
    TBCount = conn.Execute("select count(*) from blog_Trackback")(0)
    Count_Member = conn.Execute("select count(*) from blog_Member")(0)
    conn.Execute("update blog_Info set blog_tbNums="&TBCount)
    conn.Execute("update blog_Info set blog_MemNums="&Count_Member)
    Set blog_Content_count = conn.Execute("SELECT log_CateID, Count(log_CateID) FROM blog_Content where log_IsDraft=false GROUP BY log_CateID")
    Set blog_Comment_count = conn.Execute("SELECT Count(*) FROM blog_Comment")
    Do While Not blog_Content_count.EOF
        ContentCount = ContentCount + blog_Content_count(1)
        conn.Execute("update blog_Category set cate_count="&blog_Content_count(1)&" where cate_ID="&blog_Content_count(0))
        blog_Content_count.movenext
    Loop
    conn.Execute("update blog_Info set blog_LogNums="&ContentCount)
    conn.Execute("update blog_Info set blog_CommNums="&blog_Comment_count(0))
    getInfo(2)
    CategoryList(2)
    session(CookieName&"_ShowMsg") = True
    session(CookieName&"_MsgText") = session(CookieName&"_MsgText")&"数据统计完成! "
End If
If Request.Form("CleanVisitor") = 1 Then
    conn.Execute("delete * from blog_Counter")
    session(CookieName&"_ShowMsg") = True
    session(CookieName&"_MsgText") = session(CookieName&"_MsgText")&"访客数据清除完成! "
End If
If Request.Form("ReBulid") = 1 Then
    FreeMemory
    session(CookieName&"_ShowMsg") = True
    session(CookieName&"_MsgText") = session(CookieName&"_MsgText")&"缓存重建成功! "
End If
RedirectUrl("ConContent.asp?Fmenu=General&Smenu=Misc")
Else
    session(CookieName&"_ShowMsg") = True
    session(CookieName&"_MsgText") = "非法提交内容"
    RedirectUrl("ConContent.asp?Fmenu=General&Smenu=")
End If
'==========================处理日志分类===============================
ElseIf Request.Form("action") = "Categories" Then
    '--------------------------处理日志批量移动----------------------------
    If Request.Form("whatdo") = "move" Then
        Dim Cate_source, Cate_target, Cate_source_name, Cate_source_count, Cate_target_name
        Cate_source = Int(CheckStr(Request.Form("source")))
        Cate_target = Int(CheckStr(Request.Form("target")))
        If Cate_source = Cate_target Then
            session(CookieName&"_ShowMsg") = True
            session(CookieName&"_MsgText") = "源分类和目标分类一致无法移动"
            RedirectUrl("ConContent.asp?Fmenu=Categories&Smenu=move")
        End If
        Cate_source_name = conn.Execute("select cate_Name from blog_Category where cate_ID="&Cate_source)(0)
        Cate_source_count = conn.Execute("select cate_count from blog_Category where cate_ID="&Cate_source)(0)
        Cate_target_name = conn.Execute("select cate_Name from blog_Category where cate_ID="&Cate_target)(0)
        conn.Execute ("update blog_Content set log_CateID="&Cate_target&" where log_CateID="&Cate_source)
        conn.Execute ("update blog_Category set cate_count=0 where cate_ID="&Cate_source)
        conn.Execute ("update blog_Category set cate_count=cate_count+"&Cate_source_count&" where cate_ID="&Cate_target)
        CategoryList(2)
        session(CookieName&"_ShowMsg") = True
        session(CookieName&"_MsgText") = "<span style=""color:#f00"">"&Cate_source_name&"</span> 移动到 <span style=""color:#f00"">"&Cate_target_name&"</span> 成功! 批量转移后，请到 <a href=""ConContent.asp?Fmenu=General&Smenu=Misc"" style=""color:#00f"">站点基本设置-初始化数据</a> ,重新生成所有日志到文件"
        RedirectUrl("ConContent.asp?Fmenu=Categories&Smenu=move")
        '--------------------------处理日志分类----------------------------
    ElseIf Request.Form("whatdo") = "Cate" Then
        '处理存在分类
        Dim LCate_ID, LCate_icons, LCate_Name, LCate_Intro, Lcate_URL, Lcate_Order, Lcate_count, LCate_local, Lcate_Secret
        Dim NCate_ID, NCate_icons, NCate_Name, NCate_Intro, Ncate_URL, Ncate_Order, NCate_local, Ncate_OutLink, Ncate_Secret
        LCate_ID = Split(Request.Form("Cate_ID"), ", ")
        LCate_Name = Split(Request.Form("Cate_Name"), ", ")
        LCate_icons = Split(Request.Form("Cate_icons"), ", ")
        LCate_Intro = Split(Request.Form("Cate_Intro"), ", ")
        Lcate_URL = Split(Request.Form("cate_URL"), ", ")
        Lcate_Order = Split(Request.Form("cate_Order"), ", ")
        Lcate_count = Split(Request.Form("cate_count"), ", ")
        LCate_local = Split(Request.Form("Cate_local"), ", ")
        Lcate_Secret = Split(Request.Form("cate_Secret"), ", ")
        For i = 0 To UBound(LCate_Name)
            SQL = "SELECT * FROM blog_Category where cate_ID="&Int(CheckStr(LCate_ID(i)))
            weblog.Open SQL, Conn, 1, 3
            weblog("cate_Name") = CheckStr(LCate_Name(i))
            weblog("cate_icon") = CheckStr(LCate_icons(i))
            weblog("Cate_Intro") = CheckStr(LCate_Intro(i))
            If Len(Trim(Lcate_URL(i)))>1 And Int(Lcate_count(i))<1 Then
                weblog("cate_URL") = Trim(CheckStr(Lcate_URL(i)))
                weblog("cate_OutLink") = True
            Else
                weblog("cate_URL") = Trim(CheckStr(Lcate_URL(i)))
                weblog("cate_OutLink") = False
            End If
            weblog("cate_Order") = Int(CheckStr(Lcate_Order(i)))
            weblog("Cate_local") = Int(CheckStr(LCate_local(i)))
            weblog("cate_Secret") = CBool(CheckStr(Lcate_Secret(i)))
            weblog.update
            weblog.Close
        Next
        '判断添加新日志
        NCate_Name = Trim(CheckStr(Request.Form("New_Cate_Name")))
        NCate_icons = CheckStr(Request.Form("New_Cate_icons"))
        NCate_Intro = Trim(CheckStr(Request.Form("New_Cate_Intro")))
        Ncate_URL = Trim(CheckStr(Request.Form("New_cate_URL")))
        Ncate_Order = CheckStr(Request.Form("New_cate_Order"))
        NCate_local = CheckStr(Request.Form("New_Cate_local"))
        Ncate_Secret = CheckStr(Request.Form("New_Cate_Secret"))
        If Len(NCate_Name)>0 Then
            If Len(Ncate_Order)<1 Then Ncate_Order = conn.Execute("select count(*) from blog_Category")(0)
            If Len(Ncate_URL)>0 Then Ncate_OutLink = True Else Ncate_OutLink = False
            Dim AddCateArray
            AddCateArray = Array(Array("cate_Name", NCate_Name), Array("cate_icon", NCate_icons), Array("Cate_Intro", NCate_Intro), Array("cate_URL", Ncate_URL), Array("cate_OutLink", Ncate_OutLink), Array("cate_Order", Int(Ncate_Order)), Array("Cate_local", NCate_local), Array("Cate_Secret", NCate_Secret))
            If DBQuest("blog_Category", AddCateArray, "insert") = 0 Then session(CookieName&"_MsgText") = "新日志分类添加成功，"
        End If
        FreeMemory
        session(CookieName&"_ShowMsg") = True
        session(CookieName&"_MsgText") = session(CookieName&"_MsgText")&"日志分类更新成功!"
        RedirectUrl("ConContent.asp?Fmenu=Categories&Smenu=")
        '--------------------------批量删除日志----------------------------
    ElseIf Request.Form("whatdo") = "batdel" Then
        Dim CID, Cids, tti, C1, C2
        C1 = 0
        C2 = 0
        CID = checkstr(request.Form("CID"))
        Cids = Split(CID, ", ")
        For tti = 0 To UBound(Cids)
            If DeleteLog(Cids(tti)) = 1 Then
                C1 = C1 + 1
            Else
                C2 = C2 + 1
            End If
        Next
        FreeMemory
        session(CookieName&"_ShowMsg") = True
        session(CookieName&"_MsgText") = "日志删除状态:"&C1&"篇成功,"&C2&"篇失败! 批量删除后，请到 <a href=""ConContent.asp?Fmenu=General&Smenu=Misc"" style=""color:#00f"">站点基本设置-初始化数据</a> ,重新生成所有日志到文件"
        RedirectUrl("ConContent.asp?Fmenu=Categories&Smenu=del")
        '--------------------------删除日志分类----------------------------
    ElseIf Request.Form("whatdo") = "DelCate" Then
        Dim DelCate, DelLog, DelID, conCount, comCount, P1, P2
        P1 = 0
        p2 = 0
        DelCate = Request.Form("DelCate")
        Set DelLog = Conn.Execute("select log_ID FROM blog_Content WHERE log_CateID="&DelCate)
        Do While Not DelLog.EOF
            DelID = DelLog("log_ID")
            If DeleteLog(DelID) = 1 Then
                P1 = P1 + 1
            Else
                P2 = P2 + 1
            End If
            DelLog.movenext
        Loop
        Conn.Execute("DELETE * FROM blog_Category WHERE cate_ID="&DelCate)
        FreeMemory
        session(CookieName&"_ShowMsg") = True
        session(CookieName&"_MsgText") = session(CookieName&"_MsgText")&"日志分类删除成功! 日志删除状态:"&P1&"篇成功,"&P2&"篇失败! 批量删除后，请到 <a href=""ConContent.asp?Fmenu=General&Smenu=Misc"" style=""color:#00f"">站点基本设置-初始化数据</a> ,重新生成所有日志到文件"
        RedirectUrl("ConContent.asp?Fmenu=Categories&Smenu=")
    ElseIf Request.Form("whatdo") = "Tag" Then
        Dim TagsID, TagName
        If Request.Form("doModule") = "DelSelect" Then
            TagsID = Split(Request.Form("selectTagID"), ", ")
            For i = 0 To UBound(TagsID)
                conn.Execute("DELETE * from blog_tag where tag_id="&TagsID(i))
            Next
            Tags(2)
            session(CookieName&"_ShowMsg") = True
            session(CookieName&"_MsgText") = session(CookieName&"_MsgText")&(UBound(TagsID) + 1)&" 个Tag被删除!"
            If blog_postFile>0 Then session(CookieName&"_MsgText") = session(CookieName&"_MsgText") + "由于你使用了静态日志功能，所以删除tag后建议到 <a href='ConContent.asp?Fmenu=General&Smenu=Misc' title='站点基本设置-初始化数据 '>初始化数据</a> 重新生成所有日志一次."
            RedirectUrl("ConContent.asp?Fmenu=Categories&Smenu=tag")
        Else
            TagsID = Split(Request.Form("TagID"), ", ")
            TagName = Split(Request.Form("tagName"), ", ")
            For i = 0 To UBound(TagsID)
                If Int(TagsID(i))<> -1 Then
                    conn.Execute("update blog_tag set tag_name='"&CheckStr(TagName(i))&"' where tag_id="&TagsID(i))
                Else
                    If Len(Trim(CheckStr(TagName(i))))>0 Then
                        conn.Execute("insert into blog_tag (tag_name,tag_count) values ('"&CheckStr(TagName(i))&"',0)")
                        session(CookieName&"_MsgText") = "新Tag添加成功! "
                    End If
                End If
            Next
            Tags(2)
            session(CookieName&"_ShowMsg") = True
            session(CookieName&"_MsgText") = session(CookieName&"_MsgText")&"Tag保存成功!"
            RedirectUrl("ConContent.asp?Fmenu=Categories&Smenu=tag")
        End If
    Else
        session(CookieName&"_ShowMsg") = True
        session(CookieName&"_MsgText") = "非法提交内容!"
        RedirectUrl("ConContent.asp?Fmenu=Categories&Smenu=")
    End If

    '==========================评论留言处理===============================
ElseIf Request.Form("action") = "Comment" Then

    Dim selCommID, doCommID, doTitle, doRedirect, t1, t2
    selCommID = Split(Request.Form("selectCommentID"), ", ")
    doCommID = Split(Request.Form("CommentID"), ", ")

    If Request.Form("doModule") = "updateKey" Then
        saveFilterKey Request.Form("keyList")
    ElseIf Request.Form("doModule") = "updateRegKey" Then
        saveReFilterKey
    ElseIf Request.Form("doModule") = "DelSelect" Then
        For i = 0 To UBound(selCommID)
            If Request.Form("whatdo") = "trackback" Then
                t1 = Int(Split(selCommID(i), "|")(0))
                t2 = Int(Split(selCommID(i), "|")(1))
                conn.Execute("UPDATE blog_Content SET log_QuoteNums=log_QuoteNums-1 WHERE log_ID="&t2)
                conn.Execute("DELETE * from blog_Trackback where tb_ID="&t1)
                conn.Execute("UPDATE blog_Info Set blog_tbNums=blog_tbNums-1")
                doTitle = "引用通告"
                PostArticle t2, False
                doRedirect = "trackback"
            ElseIf Request.Form("whatdo") = "msg" Then
                conn.Execute("DELETE * from blog_book where book_ID="&selCommID(i))
                doTitle = "留言"
                doRedirect = "msg"
            Else
                t1 = Int(Split(selCommID(i), "|")(0))
                t2 = Int(Split(selCommID(i), "|")(1))
                conn.Execute("update blog_Content set log_CommNums=log_CommNums-1 where log_ID="&t2)
                conn.Execute("update blog_Info set blog_CommNums=blog_CommNums-1")
                conn.Execute("DELETE * from blog_Comment where comm_ID="&t1)
                doTitle = "评论"
                doRedirect = ""
                PostArticle t2, False
            End If
        Next

        getInfo(2)
        NewComment(2)
        Application.Lock
        Application(CookieName&"_blog_Message") = ""
        Application.UnLock
        session(CookieName&"_ShowMsg") = True
        session(CookieName&"_MsgText") = (UBound(selCommID) + 1)&" 个"&doTitle&"记录被删除!"
        RedirectUrl("ConContent.asp?Fmenu=Comment&Smenu="&doRedirect)
    ElseIf Request.Form("doModule") = "Update" Then
        For i = 0 To UBound(doCommID)
            If Request.Form("whatdo") = "msg" Then
                If Int(Request.Form("edited_"&doCommID(i))) = 1 Then
                    conn.Execute("UPDATE blog_book SET book_Content='"&checkStr(Request.Form("message_"&doCommID(i)))&"',book_replyAuthor='"&memName&"',book_replyTime=#"&DateToStr(Now(), "Y-m-d H:I:S")&"#,book_reply='"&checkStr(Request.Form("reply_"&doCommID(i)))&"' WHERE book_ID="&doCommID(i))
                Else
                    conn.Execute("UPDATE blog_book SET book_Content='"&checkStr(Request.Form("message_"&doCommID(i)))&"',book_replyAuthor='"&memName&"',book_reply='"&checkStr(Request.Form("reply_"&doCommID(i)))&"' WHERE book_ID="&doCommID(i))
                End If
                doTitle = "留言"
                doRedirect = "msg"
            ElseIf Request.Form("whatdo") = "comment" Then
                conn.Execute("UPDATE blog_Comment SET comm_Content='"&checkStr(Request.Form("message_"&doCommID(i)))&"' WHERE comm_ID="&doCommID(i))
                doTitle = "评论"
                doRedirect = ""
            End If
        Next

        NewComment(2)
        Application.Lock
        Application(CookieName&"_blog_Message") = ""
        Application.UnLock
        session(CookieName&"_ShowMsg") = True
        session(CookieName&"_MsgText") = doTitle&"记录更新成功!"
        RedirectUrl("ConContent.asp?Fmenu=Comment&Smenu="&doRedirect)
    End If

    '==========================处理帐户和权限信息===============================
ElseIf Request.Form("action") = "Members" Then
    '--------------------------处理权限组信息----------------------------
    If Request.Form("whatdo") = "Group" Then
        Dim status_name, status_title, Rights, allCount, addCount
        status_name = Split(Request.Form("status_name"), ", ")
        status_title = Split(Request.Form("status_title"), ", ")
        allCount = UBound(status_name)
        Dim NS_Name, NS_Title, NS_Code, NS_UpSize, NS_UploadType, tmpNS
        If Len(status_name(allCount))>0 Then
            NS_Name = CheckStr(status_name(allCount))
            NS_Title = CheckStr(status_title(allCount))
            If Not IsValidChars(NS_Name) Then
                session(CookieName&"_ShowMsg") = True
                session(CookieName&"_MsgText") = "<span style=""color:#900"">添加新权限失败!权限标识不能为英文或数字以外的字符</span>"
                RedirectUrl("ConContent.asp?Fmenu=Members&Smenu=")
            End If
            Set tmpNS = conn.Execute("select stat_name from blog_status where stat_name='"&NS_Name&"'")
            If Not tmpNS.EOF Then
                session(CookieName&"_ShowMsg") = True
                session(CookieName&"_MsgText") = "<span style=""color:#900"">“"&NS_Name&"”</span> 权限标识已经存在无法添加新分组!"
                RedirectUrl("ConContent.asp?Fmenu=Members&Smenu=")
            End If
            conn.Execute("insert into blog_status (stat_name,stat_title,stat_Code,stat_attSize,stat_attType) values ('"&NS_Name&"','"&NS_Title&"','000000000000',0,'')")
            session(CookieName&"_MsgText") = "新分组添加成功!"
        End If
        For i = 0 To UBound(status_name) -1
            conn.Execute("update blog_status set stat_title='"&CheckStr(status_title(i))&"' where stat_name='"&CheckStr(status_name(i))&"'")
        Next
        UserRight(2)
        session(CookieName&"_ShowMsg") = True
        session(CookieName&"_MsgText") = session(CookieName&"_MsgText")&"权限组信息修改成功!"
        RedirectUrl("ConContent.asp?Fmenu=Members&Smenu=")
        '--------------------------处理帐户信息----------------------------
    ElseIf Request.Form("whatdo") = "User" Then

        '--------------------------编辑分组信息----------------------------
    ElseIf Request.Form("whatdo") = "EditGroup" Then
        Dim EditGroup, AddArticle, EditArticle, DelArticle, AddComment, DelComment, ShowHiddenCate, IsAdmin, CanUpload, UploadSize, UploadType, Group_title, SCode
        EditGroup = CheckStr(Request.Form("status_name"))
        AddArticle = CheckStr(Request.Form("AddArticle"))
        EditArticle = CheckStr(Request.Form("EditArticle"))
        DelArticle = CheckStr(Request.Form("DelArticle"))
        AddComment = CheckStr(Request.Form("AddComment"))
        DelComment = CheckStr(Request.Form("DelComment"))
        ShowHiddenCate = CheckStr(Request.Form("ShowHiddenCate"))
        IsAdmin = CheckStr(Request.Form("IsAdmin"))
        CanUpload = CheckStr(Request.Form("CanUpload"))
        UploadSize = CheckStr(Request.Form("UploadSize"))
        UploadType = CheckStr(Request.Form("UploadType"))
        Group_title = CheckStr(Request.Form("status_title"))
        SCode = AddArticle & EditArticle & DelArticle &_
        AddComment & DelComment & CanUpload & IsAdmin & ShowHiddenCate
        conn.Execute("update blog_status set stat_title='"&Group_title&"',stat_code='"&SCode&"',stat_attSize="&UploadSize&",stat_attType='"&UploadType&"' where stat_name='"&EditGroup&"'")
        UserRight(2)
        session(CookieName&"_ShowMsg") = True
        session(CookieName&"_MsgText") = "<span style=""color:#900"">“"&EditGroup&"”</span>权限分组 编辑成功!"
        RedirectUrl("ConContent.asp?Fmenu=Members&Smenu=EditRight&id="&EditGroup)
        '--------------------------删除分组信息----------------------------
    ElseIf Request.Form("whatdo") = "DelGroup" Then
        Dim DelGroup
        DelGroup = CheckStr(Request.Form("DelGroup"))
        If LCase(DelGroup)<>"supadmin" And LCase(DelGroup)<>"member" And LCase(DelGroup)<>"guest" Then
            conn.Execute ("update blog_Member set mem_Status='Member' where mem_Status='"&DelGroup&"'")
            Conn.Execute("DELETE * FROM blog_status WHERE stat_name='"&DelGroup&"'")
            UserRight(2)
            session(CookieName&"_ShowMsg") = True
            session(CookieName&"_MsgText") = "<span style=""color:#900"">“"&DelGroup&"”</span>权限分组 删除成功!"
            RedirectUrl("ConContent.asp?Fmenu=Members&Smenu=")
        Else
            session(CookieName&"_ShowMsg") = True
            session(CookieName&"_MsgText") = "特殊分组无法删除!"
            RedirectUrl("ConContent.asp?Fmenu=Members&Smenu=")
        End If
        '--------------------------保存用户权限----------------------------
    ElseIf Request.Form("whatdo") = "SaveUserRight" Then
        For i = 1 To Request.Form("mem_ID").Count
            conn.Execute("update blog_Member set mem_Status='"&Request.Form("mem_Status").Item(i)&"' where mem_ID="&Request.Form("mem_ID").Item(i))
        Next

        session(CookieName&"_ShowMsg") = True
        session(CookieName&"_MsgText") = "用户权限设置成功!"
        RedirectUrl("ConContent.asp?Fmenu=Members&Smenu=Users")
        '--------------------------删除用户----------------------------
    ElseIf Request.Form("whatdo") = "DelUser" Then
        Dim DelUserID, DelUserName, blogmemberNum, DelUserStatus
        DelUserID = Request.Form("DelID")
        blogmemberNum = conn.Execute("select count(mem_ID) from blog_Member where mem_Status='SupAdmin'")(0)

        DelUserStatus = conn.Execute("select mem_Status from blog_Member where mem_ID="&DelUserID)(0)
        If ((blogmemberNum = 1) And (DelUserStatus = "SupAdmin")) Then
            session(CookieName&"_ShowMsg") = True
            session(CookieName&"_MsgText") = "不能删除仅有的管理员权限!"
            RedirectUrl("ConContent.asp?Fmenu=Members&Smenu=Users")
        Else
            DelUserName = conn.Execute("select mem_Name from blog_Member where mem_ID="&DelUserID)(0)
            conn.Execute("delete * from blog_Member where mem_ID="&DelUserID)
            Conn.Execute("UPDATE blog_Info SET blog_MemNums=blog_MemNums-1")
            getInfo(2)
            session(CookieName&"_ShowMsg") = True
            session(CookieName&"_MsgText") = "<span style=""color:#900"">“"&DelUserName&"”</span> 删除成功!"
            RedirectUrl("ConContent.asp?Fmenu=Members&Smenu=Users")
        End If
    Else
        session(CookieName&"_ShowMsg") = True
        session(CookieName&"_MsgText") = "非法提交内容!"
        RedirectUrl("ConContent.asp?Fmenu=Members&Smenu=")
    End If
    '==========================友情链接管理===============================
ElseIf Request.Form("action") = "Links" Then
    Dim LinkID, LinkName, LinkURL, LinkLogo, LinkOrder, LinkMain
    '--------------------------友情链接过滤----------------------------
    If Request.Form("whatdo") = "Filter" Then
        session(CookieName&"_disLink") = CheckStr(Request.Form("disLink"))
        session(CookieName&"_disCount") = CheckStr(Request.Form("disCount"))
        RedirectUrl("ConContent.asp?Fmenu=Link&Smenu=")
        '--------------------------保存友情链接----------------------------
    ElseIf Request.Form("whatdo") = "SaveLink" Then
        Dim TLinkName, TLinkURL, TLinkLogo, TLinkOrder
        LinkID = Split(Request.Form("LinkID"), ", ")
        LinkName = Split(Request.Form("LinkName"), ", ")
        LinkURL = Split(Request.Form("LinkURL"), ", ")
        LinkLogo = Split(Request.Form("LinkLogo"), ", ")
        LinkOrder = Split(Request.Form("LinkOrder"), ", ")
        For i = 0 To UBound(LinkID)
            If UBound(LinkName)<0 Then TLinkName = "未知" Else TLinkName = LinkName(i)
            If UBound(LinkURL)<0 Then TLinkURL = "http://" Else TLinkURL = LinkURL(i)
            If UBound(LinkLogo)<0 Then TLinkLogo = "" Else TLinkLogo = LinkLogo(i)
            If UBound(LinkOrder)<0 Then TLinkOrder = "0" Else TLinkOrder = LinkOrder(i)
            conn.Execute("update blog_Links set link_Name='"&CheckStr(TLinkName)&"',link_URL='"&CheckStr(TLinkURL)&"',link_Image='"&CheckStr(TLinkLogo)&"',link_Order='"&CheckStr(TLinkOrder)&"' where link_ID="&LinkID(i))
        Next
        LinkID = Request.Form("new_LinkID")
        LinkName = Request.Form("new_LinkName")
        LinkURL = Request.Form("new_LinkURL")
        LinkLogo = Request.Form("new_LinkLogo")
        LinkOrder = Request.Form("new_LinkOrder")
        If Len(LinkOrder)<1 Then LinkOrder = conn.Execute("select count(*) from blog_Links")(0)
        If Len(Trim(CheckStr(LinkName)))>0 Then
            conn.Execute("insert into blog_Links (link_Name,link_URL,link_Image,link_Order,link_IsShow) values ('"&CheckStr(LinkName)&"','"&CheckStr(LinkURL)&"','"&CheckStr(LinkLogo)&"','"&CheckStr(LinkOrder)&"',true)")
            session(CookieName&"_MsgText") = "新友情链接添加成功! "
        End If
        Bloglinks(2)
        PostLink
        session(CookieName&"_ShowMsg") = True
        session(CookieName&"_MsgText") = session(CookieName&"_MsgText")&"保存链接成功!"
        RedirectUrl("ConContent.asp?Fmenu=Link&Smenu=&page="&Request.Form("page"))
        '--------------------------通过友情链接----------------------------
    ElseIf Request.Form("whatdo") = "ShowLink" Then
        conn.Execute ("update blog_Links set link_IsShow=true where link_ID="&CheckStr(Request.Form("ALinkID")))
        LinkName = conn.Execute ("select link_Name from blog_Links where link_ID="&CheckStr(Request.Form("ALinkID")))(0)
        session(CookieName&"_ShowMsg") = True
        session(CookieName&"_MsgText") = session(CookieName&"_MsgText")&"<span style=""color:#900"">“"&LinkName&"”</span> 通过验证!"
        Bloglinks(2)
        PostLink
        RedirectUrl("ConContent.asp?Fmenu=Link&Smenu=&page="&Request.Form("page"))
        '--------------------------置顶友情链接----------------------------
    ElseIf Request.Form("whatdo") = "TopLink" Then
        conn.Execute ("update blog_Links set link_IsMain=not link_IsMain where link_ID="&CheckStr(Request.Form("ALinkID")))
        LinkName = conn.Execute ("select link_Name from blog_Links where link_ID="&CheckStr(Request.Form("ALinkID")))(0)
        LinkMain = conn.Execute ("select link_IsMain from blog_Links where link_ID="&CheckStr(Request.Form("ALinkID")))(0)
        session(CookieName&"_ShowMsg") = True
        If LinkMain Then
            session(CookieName&"_MsgText") = session(CookieName&"_MsgText")&"<span style=""color:#900"">“"&LinkName&"”</span> 置顶成功!"
        Else
            session(CookieName&"_MsgText") = session(CookieName&"_MsgText")&"<span style=""color:#900"">“"&LinkName&"”</span> 取消首页置顶!"
        End If
        Bloglinks(2)
        PostLink
        RedirectUrl("ConContent.asp?Fmenu=Link&Smenu=&page="&Request.Form("page"))
        '--------------------------删除友情链接----------------------------
    ElseIf Request.Form("whatdo") = "DelLink" Then
        LinkName = conn.Execute ("select link_Name from blog_Links where link_ID="&CheckStr(Request.Form("ALinkID")))(0)
        conn.Execute ("DELETE * from blog_Links where link_ID="&CheckStr(Request.Form("ALinkID")))
        Session(CookieName&"_ShowMsg") = True
        Session(CookieName&"_MsgText") = session(CookieName&"_MsgText")&"<span style=""color:#900"">“"&LinkName&"”</span> 删除成功!"
        Bloglinks(2)
        PostLink
        RedirectUrl("ConContent.asp?Fmenu=Link&Smenu=&page="&Request.Form("page"))
    End If
    '==========================表情和关键字===============================
ElseIf Request.Form("action") = "smilies" Then
    Dim smilesID, smiles, smilesURL
    Dim KeyWordID, KeyWord, KeyWordURL
    '--------------------------处理表情符号----------------------------
    If Request.Form("whatdo") = "smilies" Then
        If Request.Form("doModule") = "DelSelect" Then
            smilesID = Split(Request.Form("selectSmiliesID"), ", ")
            For i = 0 To UBound(smilesID)
                conn.Execute("DELETE * from blog_Smilies where sm_ID="&smilesID(i))
            Next
            Smilies(2)
            session(CookieName&"_ShowMsg") = True
            session(CookieName&"_MsgText") = session(CookieName&"_MsgText")&(UBound(smilesID) + 1)&" 个表情被删除!"
            RedirectUrl("ConContent.asp?Fmenu=smilies&Smenu=")
        Else
            smilesID = Split(Request.Form("smilesID"), ", ")
            smiles = Split(Request.Form("smiles"), ", ")
            smilesURL = Split(Request.Form("smilesURL"), ", ")
            For i = 0 To UBound(smilesID)
                If Int(smilesID(i))<> -1 Then
                    conn.Execute("update blog_Smilies set sm_Text='"&CheckStr(smiles(i))&"',sm_Image='"&CheckStr(smilesURL(i))&"' where sm_ID="&smilesID(i))
                Else
                    If Len(Trim(CheckStr(smiles(i))))>0 Then
                        conn.Execute("insert into blog_Smilies (sm_Text,sm_Image) values ('"&CheckStr(smiles(i))&"','"&CheckStr(smilesURL(i))&"')")
                        session(CookieName&"_MsgText") = "新表情添加成功! "
                    End If
                End If
            Next
            Smilies(2)
            session(CookieName&"_ShowMsg") = True
            session(CookieName&"_MsgText") = session(CookieName&"_MsgText")&"表情保存成功!"
            RedirectUrl("ConContent.asp?Fmenu=smilies&Smenu=")
        End If
    ElseIf Request.Form("whatdo") = "KeyWord" Then
        If Request.Form("doModule") = "DelSelect" Then
            KeyWordID = Split(Request.Form("SelectKeyWordID"), ", ")
            For i = 0 To UBound(KeyWordID)
                conn.Execute("DELETE * from blog_Keywords where key_ID="&KeyWordID(i))
            Next
            Keywords(2)
            session(CookieName&"_ShowMsg") = True
            session(CookieName&"_MsgText") = session(CookieName&"_MsgText")&(UBound(KeyWordID) + 1)&"关键字被删除!"
            RedirectUrl("ConContent.asp?Fmenu=smilies&Smenu=KeyWord")
        Else
            KeyWordID = Split(Request.Form("KeyWordID"), ", ")
            KeyWord = Split(Request.Form("KeyWord"), ", ")
            KeyWordURL = Split(Request.Form("KeyWordURL"), ", ")
            For i = 0 To UBound(KeyWordID)
                If Int(KeyWordID(i))<> -1 Then
                    conn.Execute("update blog_Keywords set key_Text='"&CheckStr(KeyWord(i))&"',key_URL='"&CheckStr(KeyWordURL(i))&"' where key_ID="&KeyWordID(i))
                Else
                    If Len(Trim(CheckStr(KeyWord(i))))>0 Then
                        conn.Execute("insert into blog_Keywords (key_Text,key_URL) values ('"&CheckStr(KeyWord(i))&"','"&CheckStr(KeyWordURL(i))&"')")
                        session(CookieName&"_MsgText") = "新关键字添加成功! "
                    End If
                End If
            Next
            Keywords(2)
            session(CookieName&"_ShowMsg") = True
            session(CookieName&"_MsgText") = session(CookieName&"_MsgText")&"关键字保存成功!"
            RedirectUrl("ConContent.asp?Fmenu=smilies&Smenu=KeyWord")
        End If
    Else
        session(CookieName&"_ShowMsg") = True
        session(CookieName&"_MsgText") = "非法提交内容!"
        RedirectUrl("ConContent.asp?Fmenu=smilies&Smenu=")
    End If

    '==========================设置界面和模版===============================
ElseIf Request.Form("action") = "Skins" Then
    Dim skinpath, Skinname, moduleID, moduleName, moduleType, moduleTitle, moduleHidden, moduleTop, moduleOrder, moduleHtmlCode, mOrder
    '--------------------------设置默认界面----------------------------
    If Request.Form("whatdo") = "setDefaultSkin" Then
        skinpath = CheckStr(Request.Form("SkinPath"))
        Skinname = CheckStr(Request.Form("SkinName"))
        conn.Execute("update blog_Info set blog_DefaultSkin='"&skinpath&"'")
        getInfo(2)
        session(CookieName&"_ShowMsg") = True
        session(CookieName&"_MsgText") = "<span style=""color:#900"">“"&Skinname&"”</span> 设置为当前默认界面!"
        RedirectUrl("ConContent.asp?Fmenu=Skins&Smenu=")
        '--------------------------保存模块设置----------------------------
    ElseIf Request.Form("whatdo") = "UpdateModule" Then
        Dim selectID, doModule
        selectID = Split(Request.Form("selectID"), ", ")
        doModule = Request.Form("doModule")
        moduleID = Split(Request.Form("mID"), ", ")
        moduleName = Split(Request.Form("mName"), ", ")
        moduleType = Split(Request.Form("mType"), ", ")
        moduleTitle = Split(Request.Form("mTitle"), ", ")
        moduleHidden = Split(Request.Form("mHidden"), ", ")
        moduleTop = Split(Request.Form("mTop"), ", ")
        moduleOrder = Split(Request.Form("mOrder"), ", ")
        ',IsHidden="&CBool(moduleHidden(i))&",IndexOnly="&CBool(moduleTop(i))&"
        For i = 0 To UBound(moduleID)
            If Int(moduleID(i))<> -1 Then
                If Not conn.Execute ("select IsSystem from blog_module where id="&moduleID(i))(0) Then
                    conn.Execute("update blog_module set title='"&CheckStr(moduleTitle(i))&"',type='"&CheckStr(moduleType(i))&"',SortID="&CheckStr(moduleOrder(i))&" where id="&moduleID(i))
                Else
                    mOrder = CheckStr(moduleOrder(i))
                    If CheckStr(moduleName(i)) = "ContentList" Then mOrder = 0
                    conn.Execute("update blog_module set title='"&CheckStr(moduleTitle(i))&"',SortID="&mOrder&" where id="&moduleID(i))
                End If
            Else
                If Len(Trim(CheckStr(moduleName(i))))>0 Then
                    If Not conn.Execute("select name from blog_module where name='"&CheckStr(moduleName(i))&"'").EOF Then
                        session(CookieName&"_ShowMsg") = True
                        session(CookieName&"_MsgText") = "<span style=""color:#900"">“"&CheckStr(moduleName(i))&"”</span> 模块标识已经存在无法添加新模块!"
                        RedirectUrl("ConContent.asp?Fmenu=Skins&Smenu=module")
                    End If
                    If Not IsValidChars(CheckStr(moduleName(i))) Then
                        session(CookieName&"_ShowMsg") = True
                        session(CookieName&"_MsgText") = "<span style=""color:#900"">添加新模块失败!权限标识不能为英文或数字以外的字符</span>"
                        RedirectUrl("ConContent.asp?Fmenu=Skins&Smenu=module")
                    End If
                    If Len(CheckStr(moduleOrder(i)))<1 Then
                        mOrder = conn.Execute("select count(id) from blog_module")(0)
                    Else
                        If IsInteger(CheckStr(moduleOrder(i))) = False Then
                            session(CookieName&"_ShowMsg") = True
                            session(CookieName&"_MsgText") = "输入非法，添加失败!"
                            RedirectUrl("ConContent.asp?Fmenu=Skins&Smenu=module")
                        End If
                        mOrder = CheckStr(moduleOrder(i))
                    End If
                    conn.Execute("insert into blog_module (name,title,type,IsHidden,IndexOnly,SortID) values ('"&CheckStr(moduleName(i))&"','"&CheckStr(moduleTitle(i))&"','"&CheckStr(moduleType(i))&"',false,false,"&mOrder&")")
                    session(CookieName&"_MsgText") = "新模块添加成功! "
                End If
            End If
        Next
        For i = 0 To UBound(selectID)
            Select Case doModule
Case "dohidden":
                conn.Execute("update blog_module set IsHidden=true where id="&selectID(i))
Case "cancelhidden":
                conn.Execute("update blog_module set IsHidden=false where id="&selectID(i))
Case "doIndex":
                conn.Execute("update blog_module set IndexOnly=true where id="&selectID(i))
Case "cancelIndex":
                conn.Execute("update blog_module set IndexOnly=false where id="&selectID(i))
            End Select
        Next
        log_module(2)
        session(CookieName&"_ShowMsg") = True
        session(CookieName&"_MsgText") = session(CookieName&"_MsgText")&"模块保存成功!"
        RedirectUrl("ConContent.asp?Fmenu=Skins&Smenu=module")

        '--------------------------编辑模块HTML代码----------------------------
    ElseIf Request.Form("whatdo") = "editModule" Then
        moduleID = Request.Form("DoID")
        moduleName = Request.Form("DoName")
        moduleHtmlCode = ClearHTML(CheckStr(request.Form("HtmlCode")))
        SavehtmlCode moduleHtmlCode, moduleID
        log_module(2)
        session(CookieName&"_ShowMsg") = True
        session(CookieName&"_MsgText") = "<span style=""color:#900"">“"&moduleName&"”</span> 代码编辑成功!"
        If Request.Form("editType") = "normal" Then
            RedirectUrl("ConContent.asp?Fmenu=Skins&Smenu=editModuleNormal&miD="&moduleID)
        Else
            RedirectUrl("ConContent.asp?Fmenu=Skins&Smenu=editModule&miD="&moduleID)
        End If

        '-------------------------------删除模块------------------------------
    ElseIf Request.Form("whatdo") = "delModule" Then
        moduleID = Request.Form("DoID")
        If conn.Execute("select isSystem from blog_module where id="&moduleID)(0) Then
            session(CookieName&"_ShowMsg") = True
            session(CookieName&"_MsgText") = "<span style=""color:#900"">“"&conn.Execute("select title from blog_module where id="&moduleID)(0)&"”</span> 是内置模块无法删除！"
            RedirectUrl("ConContent.asp?Fmenu=Skins&Smenu=module")
        Else
            moduleName = conn.Execute("select title from blog_module where id="&moduleID)(0)
            delModule moduleID
            session(CookieName&"_ShowMsg") = True
            session(CookieName&"_MsgText") = "<span style=""color:#900"">“"&moduleName&"”</span> 删除成功！"
            log_module(2)
            RedirectUrl("ConContent.asp?Fmenu=Skins&Smenu=module")
        End If
        '-------------------------------保存插件配置------------------------------
    ElseIf Request.Form("whatdo") = "SavePluginsSetting" Then
        Dim GetPlugName, GetPlugSetItems, GetPlugSetItemName, GetPlugSetItemValue
        GetPlugName = Request.Form("PluginsName")
        Set GetPlugSetItems = conn.Execute ("Select * from blog_ModSetting where set_ModName='"&GetPlugName&"'")
        Do Until GetPlugSetItems.EOF
            GetPlugSetItemName = GetPlugSetItems("set_KeyName")
            GetPlugSetItemValue = checkstr(Request.Form(GetPlugSetItemName))
            conn.Execute ("update blog_ModSetting Set set_KeyValue='"&GetPlugSetItemValue&"' where set_ModName='"&GetPlugName&"'and set_KeyName='"&GetPlugSetItemName&"'")
            GetPlugSetItems.movenext
        Loop
        Dim ModSetTemp2
        Set ModSetTemp2 = New ModSet
        ModSetTemp2.Open GetPlugName
        ModSetTemp2.ReLoad()
        session(CookieName&"_ShowMsg") = True
        session(CookieName&"_MsgText") = "<span style=""color:#900"">“"&GetPlugName&"”</span> 设置保存成功！"
        RedirectUrl("ConContent.asp?Fmenu=Skins&Smenu=PluginsOptions&Plugins="&GetPlugName)
    End If
    '==========================附件管理===============================
ElseIf Request.Form("action") = "Attachments" Then
    '-------------------------------删除模块------------------------------
    If Request.Form("whatdo") = "DelFiles" Then
        Dim getFolders, getFiles, GetFolder, GetFile, getFolderCount, getFileCount
        Dim FSODel
        Set FSODel = Server.CreateObject("Scripting.FileSystemObject")
        getFolders = Split(Request.Form("folders"), ", ")
        getFiles = Split(Request.Form("Files"), ", ")
        getFolderCount = 0
        getFileCount = 0
        For Each GetFolder in getFolders
            If Len(getPathList(GetFolder)(1))>0 Then
                session(CookieName&"_ShowMsg") = True
                session(CookieName&"_MsgText") = "<span style=""color:#900"">“"&GetFolder&"”</span> 文件夹内含有文件，无法删除!"
                RedirectUrl("ConContent.asp?Fmenu=SQLFile&Smenu=Attachments")
            End If
            If FSODel.FolderExists(Server.MapPath(GetFolder)) Then
                FSODel.DeleteFolder Server.MapPath(GetFolder), True
                getFolderCount = getFolderCount + 1
            End If
        Next
        For Each GetFile in getFiles
            If FSODel.FileExists(Server.MapPath(GetFile)) Then
                FSODel.DeleteFile Server.MapPath(GetFile), True
                getFileCount = getFileCount + 1
            End If
        Next
        session(CookieName&"_ShowMsg") = True
        session(CookieName&"_MsgText") = "有 <span style=""color:#900"">"&getFileCount&" 文件, "&getFolderCount&" 个文件夹</span> 被删除!"
        RedirectUrl("ConContent.asp?Fmenu=SQLFile&Smenu=Attachments")
    End If
Else'登录欢迎

End If

'----------------------End if--------------------
End If
%>
</div>
</body>
</html>
<%

Else
    RedirectUrl("default.asp")
End If
%>
