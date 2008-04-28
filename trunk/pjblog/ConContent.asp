<!--#include file="conCommon.asp" -->
<!--#include file="common/function.asp" -->
<!--#include file="common/library.asp" -->
<!--#include file="common/cache.asp" -->
<!--#include file="common/checkUser.asp" -->
<!--#include file="FCKeditor/fckeditor.asp" -->
<!--#include file="common/XML.asp" -->
<!--#include file="common/ModSet.asp" -->
<!--#include file="class/cls_logAction.asp" -->
<!--#include file="class/cls_control.asp" -->
<%
'***************PJblog2 后台管理页面*******************
' PJblog2 Copyright 2005
' Update:2006-6-10
'**************************************************
if not ChkPost() then 
  session(CookieName&"_System")=""
  session(CookieName&"_disLink")=""
  session(CookieName&"_disCount")=""
  %>
   <script>try{top.location="default.asp"}catch(e){location="default.asp"}</script>
  <%
end if
if session(CookieName&"_System")=true and memName<>empty and stat_Admin=true then 
dim i
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
if Request.QueryString("Fmenu")="General" then '站点基本设置
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
  <tr>
      <th class="CTitle"><%=categoryTitle%></th>
  </tr>
  <tr>
    <td class="CPanel">
    <div class="SubMenu"><a href="?Fmenu=General">设置基本信息</a> | <a href="?Fmenu=General&Smenu=visitors">查看访客记录</a> | <a href="?Fmenu=General&Smenu=Misc">初始化数据</a></div>
<%
if Request.QueryString("Smenu")="visitors" then
%>
	   <table border="0" cellpadding="2" cellspacing="1" class="TablePanel">
<%
 If CheckStr(Request.QueryString("Page"))<>Empty Then
	Curpage=CheckStr(Request.QueryString("Page"))
	If IsInteger(Curpage)=False OR Curpage<0 Then Curpage=1
 Else
	Curpage=1
 End If
 dim bCounter,PageC
 Set bCounter=Server.CreateObject("ADODB.RecordSet")
 SQL="SELECT * FROM blog_Counter order by coun_Time desc"
 bCounter.Open SQL,Conn,1,1
 PageC=0
 IF not bCounter.EOF Then
	bCounter.PageSize=30
	bCounter.AbsolutePage=CurPage
	Dim bCounter_nums
	bCounter_nums=bCounter.RecordCount
    response.write "<tr><td colspan=""6"" style=""border-bottom:1px solid #999;""><div class=""pageContent"">"&MultiPage(bCounter_nums,30,CurPage,"?Fmenu=General&Smenu=visitors&","","float:left")&"</div><div class=""Content-body"" style=""line-height:200%""></td></tr>"
%>
        <tr align="center">
          <td width="100" nowrap="nowrap" class="TDHead">访客IP</td>
          <td width="120" nowrap="nowrap" class="TDHead">访客操作系统</td>
          <td nowrap="nowrap" class="TDHead">访客浏览器</td>
          <td width="300" nowrap="nowrap" class="TDHead">访客介入地址</td>
          <td class="TDHead">访客访问时间</td>
	   </tr>
	   <%
	   	Do Until bCounter.EOF OR PageC=bCounter.PageSize
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
	    PageC=PageC+1
	loop
    bCounter.Close
    Set bCounter=Nothing
    response.write ("</table>")
 end if	   
elseif Request.QueryString("Smenu")="Misc" then%>
<form action="ConContent.asp" method="post" style="margin:0px">
	<input type="hidden" name="action" value="General"/>
	<input type="hidden" name="whatdo" value="Misc"/>
    <div align="left" style="padding:5px;"><%getMsg%>
    <input type="checkbox" name="ReBulid" value="1" id="T1"/> <label for="T1">重建数据缓存</label><br/>
    <input type="checkbox" name="ReTatol" value="1" id="T2"/> <label for="T2">重新统计网站数据</label><br/>
    <input type="checkbox" name="ReBulidArticle" value="1" id="T3"/> <label for="T3">重新生成所有日志到文件</label> <span style="color:#f00">(这个过程可能会花很长时间,由你的日志数量来决定)</span><br/>
    <input type="checkbox" name="ReBulidIndex" value="1" id="T4"/> <label for="T4">重新建立日志索引</label><br/>
    <input type="checkbox" name="CleanVisitor" value="1" id="T5"/> <label for="T5">清除访客记录</label><br/>
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
 	  <tr><td width="180" align="right">静态日志模式</td><td><input name="blog_postFile" type="checkbox" value="1" <%if blog_postFile then response.write ("checked=""checked""")%>  <%if not CheckObjInstalled("ADODB.Stream") then response.write "disabled"%>/> <span style="color:#999">把文章保存成文件, 减轻数据库负担. <%if not CheckObjInstalled("ADODB.Stream") then response.write "<b style='color:#f00'>需要 ADODB.Stream 组件支持</b>"%></span> </td></tr>
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
		<tr><td width="180" align="right">发表评论必须输入验证码</td><td width="300"><input name="blog_validate" type="checkbox" value="1" <%if blog_validate then response.write ("checked=""checked""")%>  /> <span style="color:#999">可以让会员不写验证码也可以发表评论</span> </td></tr>
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
elseif Request.QueryString("Fmenu")="Categories" then '日志分类管理
	    dim Arr_Category,blog_Cate,blog_Cate_Item,Icons_Lists,Icons_List,CateInOpstions
        dim CategoryListDB
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
  <tr>
    <th class="CTitle"><%=categoryTitle%></th>
  </tr>
  <tr>
    <td class="CPanel">
    <div class="SubMenu"><a href="?Fmenu=Categories">设置日志分类</a> | <a href="?Fmenu=Categories&Smenu=move">批量移动日志</a> | <a href="?Fmenu=Categories&Smenu=del">批量删除日志</a> | <a href="?Fmenu=Categories&Smenu=tag">Tag管理</a></div>
<%
if Request.QueryString("Smenu")="tag" then
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
	  dim BTag
      Set BTag=conn.execute("select * from blog_tag order by tag_id desc")
	  do until BTag.eof 
	   %>
	   <tr align="center">
		  <td><input name="selectTagID" type="checkbox" value="<%=BTag("tag_id")%>"/></td>
          <td><input name="TagID" type="hidden" value="<%=BTag("tag_id")%>"/><input name="tagName" type="text" size="14" class="text" value="<%=BTag("tag_name")%>"/></td>
          <td><input name="tagCount" type="text" size="2" class="text" value="<%=BTag("tag_count")%>" readonly style="background:#ffe"/> 篇</td>
	   </tr>
	   <%
	   BTag.movenext
	   loop
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
elseif Request.QueryString("Smenu")="move" then
        set CategoryListDB=conn.execute("select * from blog_Category order by cate_local asc, cate_Order desc")
	     do while not CategoryListDB.eof
	      if not CategoryListDB("cate_OutLink") then
	       CateInOpstions=CateInOpstions&"<option value="""&CategoryListDB("cate_ID")&""">&nbsp;&nbsp;"&CategoryListDB("cate_Name")&" ["&CategoryListDB("cate_count")&"]</option>"
	      end if
	      CategoryListDB.movenext
         loop
         set CategoryListDB=nothing
%>
  <form action="ConContent.asp" method="post" style="margin:0px" onsubmit="return CheckMove()">
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
	 elseif  Request.QueryString("Smenu")="del" then 
	 dim TempOsel,FilterO
	     set CategoryListDB=conn.execute("select * from blog_Category order by cate_local asc, cate_Order desc")
	     FilterO=checkstr(request.QueryString("filter"))
	     if len(FilterO)<1 then FilterO=-1
	     do while not CategoryListDB.eof
	      if not CategoryListDB("cate_OutLink") then
	       if int(FilterO)=CategoryListDB("cate_ID") then TempOsel= "selected"
	       CateInOpstions=CateInOpstions&"<option value="""&CategoryListDB("cate_ID")&""" "&TempOsel&">&nbsp;&nbsp;&nbsp; - "&CategoryListDB("cate_Name")&" ["&CategoryListDB("cate_count")&"]</option>"
	       TempOsel=""
	      end if
	      CategoryListDB.movenext
         loop
         set CategoryListDB=nothing
	 %>
  <form action="ConContent.asp" method="post" style="margin:0px">
   <input type="hidden" name="action" value="Categories"/>
   <input type="hidden" name="whatdo" value="batdel"/>
    <div align="left" style="padding:5px;"><%getMsg%>
		&nbsp;过滤器: <select name="filter" onchange="location='?Fmenu=Categories&Smenu=del&filter='+this.value">
			<option value="-1">显示所有日志</option>
				<%=CateInOpstions%>
			<option value="-3" <%if int(FilterO)=-3 then response.write "selected"%>>&nbsp;&nbsp;显示所有隐藏日志</option>
			<option value="-2" <%if int(FilterO)=-2 then response.write "selected"%>>&nbsp;&nbsp;显示所有草稿</option>
		</select>
		
		<table style="font-size:12px;margin:8px 0px 8px 0px">
		<%
		 dim DelContent
		  if int(FilterO)=-1 then 
		   SQL="select * from blog_content order by log_posttime desc"
		  elseif int(FilterO)=-2 then
		   SQL="select * from blog_content where log_IsDraft=true order by log_posttime desc"
		  elseif int(FilterO)=-3 then
		   SQL="select * from blog_content where log_IsShow=false order by log_posttime desc"
		  else
		   SQL="select * from blog_content where log_CateID="&int(FilterO)&" and log_IsDraft=false order by log_posttime desc"
		  end if
		 set DelContent=conn.execute(SQL)
			 if DelContent.eof and DelContent.bof then
				 %>
				 <tr><td>没有找到符合条件的查询</td></tr>
				 <%else
				 Dim TempImg
				 do until DelContent.eof
				 if DelContent("log_IsShow")=false then TempImg="<img src=""images/icon_lock.gif"" alt="""" border=""0"" style=""margin:0px 0px -3px 2px""/>"
					  if int(FilterO)=-2 then
						 %>
						 <tr><td><input name="CID" type="checkbox" value="<%=DelContent("log_ID")%>"/></td><td><a href="blogedit.asp?id=<%=DelContent("log_ID")%>" target="_blank"><%=DelContent("log_ID")%>. <%=DelContent("log_Title")%> <%=TempImg%></a></td></tr>
						 <%
					  else
						 %>
						 <tr><td><input name="CID" type="checkbox" value="<%=DelContent("log_ID")%>"/></td><td><a href="article.asp?id=<%=DelContent("log_ID")%>" target="_blank"><%=DelContent("log_ID")%>. <%=DelContent("log_Title")%> <%=TempImg%></a></td></tr>
						 <%
					  end if
				  TempImg=""
				  DelContent.movenext
				  loop
			  end if
			  DelContent.close
			  set DelContent=nothing
			 %>
		</table>
&nbsp;<input type="checkbox" name="checkbox" style="margin-bottom:-1px;" onclick="checkAll(this)"/> 全选
<input type="submit" name="Submit" value="删除选中的日志" class="button" style="margin-bottom:-1px;"/>
    </div>
     </form>
     <br/>
	 <%else
        set CategoryListDB=conn.execute("select * from blog_Category order by cate_local asc, cate_Order asc")
        if CheckObjInstalled("Scripting.FileSystemObject") then Icons_Lists=split(getPathList("images\icons")(1),"*")
	 %>
	 <script>
	 	var il = [];
        <%
        for each Icons_List in Icons_Lists
          response.write ("il.push('"&Icons_List&"');" & chr(13))
        next
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
	    do while not CategoryListDB.eof
        %>
        <tr id="Catetr_<%=CategoryListDB("cate_ID")%>" style="background:<%
         if int(CategoryListDB("cate_local"))=1 then
          response.write ("#a9c9e9")
         elseif int(CategoryListDB("cate_local"))=2 then
          response.write ("#bcf39e")
         else
          
         end if
        %>">
          <td align="center" nowrap>
          <img name="CateImg_<%=CategoryListDB("cate_ID")%>" src="<%=CategoryListDB("cate_icon")%>" width="16" height="16" />
         <%if CheckObjInstalled("Scripting.FileSystemObject") then%>
          <select name="Cate_icons" defaultValue="<%=CategoryListDB("cate_icon")%>" onchange="document.images['CateImg_<%=CategoryListDB("cate_ID")%>'].src=this.value;" style="width:120px;"></select>
          <%else%>
          <input name="Cate_icons" type="text" class="text" value="<%=CategoryListDB("cate_icon")%>" size="18" onchange="document.images['CateImg_<%=CategoryListDB("cate_ID")%>'].src=this.value"/>
          <%end if%>
          </td>
          <td><input name="Cate_Name" type="text" class="text" value="<%=CategoryListDB("cate_Name")%>" size="14"/></td>
          <td align="left"><input name="Cate_Intro" type="text" class="text" value="<%=CategoryListDB("cate_Intro")%>" size="30"/></td>
          <td align="left"><input name="cate_URL" type="text" value="<%=CategoryListDB("cate_URL")%>" size="30" class="text" <%if CategoryListDB("cate_count")>0 Then response.write "readonly=""readonly"" style=""background:#e5e5e5"""%>/></td>
          <td align="left"><input name="cate_Order" type="text" value="<%=CategoryListDB("cate_Order")%>" size="2" class="text"/></td>
          <td align="center">
           <select name="Cate_local" onchange="getElementById('Catetr_<%=CategoryListDB("cate_ID")%>').style.backgroundColor=(this.value==1)?'#a9c9e9':(this.value==2)?'#bcf39e':''">
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
        loop
        set CategoryListDB=nothing
        %>
        <tr align="center" bgcolor="#D5DAE0">
         <td colspan="9" class="TDHead" align="left" style="border-top:1px solid #999"><img src="images/add.gif" style="margin:0px 2px -3px 2px"/>添加日志分类</td>
        </tr>
        <tr>
          <td align="center" nowrap><img name="CateImg" src="images/icons/1.gif" width="16" height="16" />
          <%if CheckObjInstalled("Scripting.FileSystemObject") then%>
	          <select name="New_Cate_icons" onchange="document.images['CateImg'].src=this.value" style="width:120px;"></select>
          <%else%>
          <input name="New_Cate_icons" type="text" class="text" value="images/icons/1.gif" size="18" onchange="document.images['CateImg'].src=this.value"/>
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
elseif Request.QueryString("Fmenu")="Comment" Then%>
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
   if Request.QueryString("Smenu")="spam" then
     dim spamXml,spamList
     set spamXml=new PXML
     	 spamXml.XmlPath="spam.xml"
     	 spamXml.open
     	 %>
     	 <div align="left" style="padding-top:5px;"><%getMsg%>
     	 <form action="ConContent.asp" method="post" style="margin:0px" onsubmit="return addSpanKey()" name="filter">
			    <input type="hidden" name="action" value="Comment"/>
			    <input type="hidden" name="doModule" value="updateKey"/>
			    <input type="hidden" name="keyList" id="keyList" value=""/>
     	 <table border="0" cellpadding="2" cellspacing="1" class="TablePanel">
     	 <tr><td class="TDHead">过滤关键字</td></tr>
     	 <tr><td><div style="width:394px;overflow:hidden">
     	 <%
     	 spamList = "<select name=""spamList"" id=""spamList"" size=""20"" multiple=""multiple"" style=""width:400px;margin:-3px 0 -3px -3px"">"
     	 for i=0 to spamXml.GetXmlNodeLength("key")-1
	     	 spamList = spamList & "<option value=""" & spamXml.SelectXmlNodeItemText("key",i) & """>" & spamXml.SelectXmlNodeItemText("key",i) & "</option>"
     	 next
      	 response.write spamList
	     set spamXml= nothing
      %>
       </div></td></tr>
       <tr><td style="padding-bottom:5px;background:#FAE1AF"><img src="images/add.gif" alt="" style="margin:0 5px -3px"/>添加关键字：<input id="keyWord" type="text" size="27" class="text" onkeypress="this.style.backgroundColor='#fff';"/>
       <input type="Submit" name="Submit" value="添加" class="button" style="margin-bottom:-2px"/><input type="button" name="button" value="移除" class="button" style="margin-bottom:-2px" onclick="clearKey()"/>
       </td></tr>
       </table>
       		<div class="SubButton" style="text-align:left;padding:5px;margin:0px">
		       	<button accesskey="s" class="button" style="margin-bottom:0px;margin-left:-5px;" onclick="submitList()" >保存关键字列表(<u>S</u>)</button>
		       
	        </div>          
            <div style="color:#f00"><b>友情提示:</b><br/> - 添加或清除关键字后必须 <b>保存关键字列表</b>，垃圾关键字列表才生效。<br/> - 使用逗号或者空格输入字符串可以一次添加多个关键字<br/> - enter键直接插入关键字 ,用ctrl或shift键多选清除关键字</div>
      </div>
      </form>
      <%
	   elseif Request.QueryString("Smenu")="reg" then
		     dim reSpamXml,reSpamList
		     set reSpamXml=new PXML
		     	 reSpamXml.XmlPath="reg.xml"
		     	 reSpamXml.open

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
		     	 			<input name="des" type="button" class="button" value="添加" onclick="addRules()" style="font-size:12px;margin:0px;margin-top:-2px;margin-bottom:-1px;height:20px;"/>
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
		       	<button accesskey="t" class="button" style="margin-bottom:0px;margin-left:-5px;" onclick="testRules()">测试(<u>T</u>)</button>
		        <button accesskey="s" class="button" style="margin-bottom:0px" onclick="submitReList()">保存(<u>S</u>)</button>
	        </div>          
            <div style="color:#f00"><b>友情提示:</b>
	            <br/> - <b>允许匹配次数:</b> 允许规则匹配的最大次数
	            <br/> - 当匹配次数设置为0时,说明评论中不允许有次规则的文字
            </div>
			</form>
			<script>
			<%
			     for i=0 to reSpamXml.GetXmlNodeLength("key")-1
		     		 response.write "addRule('" & replace(reSpamXml.GetAttributes("key","des",i),"'","\'") & "','" & replace(reSpamXml.GetAttributes("key","re",i),"'","\'") & "','" & reSpamXml.GetAttributes("key","times",i) & "');" & chr(13)
		     	 next
			%>
			</script>
		</div>
		<%
	     	 set reSpamXml= nothing
	   else
   %>
		<form action="ConContent.asp" method="post" style="margin:0px">
		   <input type="hidden" name="action" value="Comment"/>
		   <input type="hidden" name="doModule" value="DelSelect"/>
		      <div align="left" style="padding-top:5px;"><%getMsg%>
		      <%
		      		dim blog_Comment,comm_Num,commArr,commArrLen,Pcount,aUrl,pSize,saveButton
					Pcount=0
					saveButton = "<input type=""button"" value=""保存全部"" class=""button"" onclick=""updateComment()"" style=""font-weight:bold;margin:0px;margin-bottom:5px;float:right;margin-right:6px""/>"
				    Set blog_Comment=Server.CreateObject("Adodb.RecordSet")
						if Request.QueryString("Smenu")="trackback" then
							SQL="SELECT tb_ID,tb_Intro,tb_Site,tb_PostTime,tb_Title,blog_ID,tb_URL,C.log_Title FROM blog_Content C,blog_Trackback T WHERE T.blog_ID=C.log_ID ORDER BY tb_PostTime desc"
							aUrl="?Fmenu=Comment&Smenu=trackback&"
							pSize = 100
							saveButton=""
							response.write "<input type=""hidden"" name=""whatdo"" value=""trackback""/>"
						elseif Request.QueryString("Smenu")="msg" then
							SQL="SELECT book_ID,book_Content,book_Messager,book_PostTime,book_IP,book_reply FROM blog_book ORDER BY book_PostTime desc"
							aUrl="?Fmenu=Comment&Smenu=msg&"
							pSize = 12
							response.write "<input type=""hidden"" name=""whatdo"" value=""msg""/>"
					    else '评论
							SQL="SELECT comm_ID,comm_Content,comm_Author,comm_PostTime,comm_PostIP,blog_ID,T.log_Title from blog_Comment C,blog_Content T WHERE C.blog_ID=T.log_ID ORDER BY C.comm_PostTime desc"
							aUrl="?Fmenu=Comment&"
							pSize = 15
							response.write "<input type=""hidden"" name=""whatdo"" value=""comment""/>"
						end if
		      %>
			       <div style="height:24px;">
				       <input type="button" value="删除所选内容" onclick="DelComment()" class="button" style="margin:0px;margin-bottom:5px;float:right;"/> 
				       <input type="button" value="全选" onclick="checkAll()" class="button" style="margin:0px;margin-bottom:5px;float:right;margin-right:6px"/>
				       <%=saveButton%>
				       <div id="page1" class="pageContent"></div>
			       </div>
			       <div class="msgDiv">
					<%
						blog_Comment.Open SQL,Conn,1,1
						IF blog_Comment.EOF AND blog_Comment.BOF Then
							response.write "</div>"
						else
							blog_Comment.PageSize=pSize
							blog_Comment.AbsolutePage=CurPage
							comm_Num=blog_Comment.RecordCount
										   
							commArr=blog_Comment.GetRows(comm_Num)
							blog_Comment.close
							set blog_Comment = nothing
							commArrLen=Ubound(commArr,2)
							'commArr(3,Pcount)
							do until Pcount = commArrLen+1 or Pcount = pSize
								if Request.QueryString("Smenu")="trackback" then
							      %>
							        <div class="item"><input type="hidden" name="CommentID" value="<%=commArr(0,Pcount)%>"/>
								        <div class="title"><span class="blogTitle"><a href="article.asp?id=<%=commArr(5,Pcount)%>" target="_blank" title="<%=commArr(7,Pcount)%>"><%=CutStr(commArr(7,Pcount),25)%></a></span><input type="checkbox" name="selectCommentID" value="<%=commArr(0,Pcount)%>|<%=commArr(5,Pcount)%>" onclick="highLight(this)"/><img src="images/icon_trackback.gif" alt=""/><b><a href="<%=commArr(6,Pcount)%>" target="_blank"><%=commArr(2,Pcount)%></a></b> <span class="date">[<%=DateToStr(commArr(3,Pcount),"Y-m-d H:I:S")%>]</span></div>
								        <div class="contentTB">
								         <b>标题:</b> <%=checkURL(HTMLDecode(commArr(4,Pcount)))%><br/>
								         <b>链接:</b> <a href="<%=commArr(6,Pcount)%>" target="_blank"><%=commArr(6,Pcount)%></a><br/>
								         <b>摘要:</b> <%=checkURL(HTMLDecode(commArr(1,Pcount)))%><br/>
								        </div>
								    </div>
							      <%
							      elseif Request.QueryString("Smenu")="msg" then
							      %>
							        <div class="item"><input type="hidden" name="CommentID" value="<%=commArr(0,Pcount)%>"/>
								        <div class="title"><input type="checkbox" name="selectCommentID" value="<%=commArr(0,Pcount)%>" onclick="highLight(this)"/><img src="images/reply.gif" alt=""/><b><%=HtmlEncode(commArr(2,Pcount))%></b> <span class="date">[<%=DateToStr(commArr(3,Pcount),"Y-m-d H:I:S")%> | <%=commArr(4,Pcount)%>]</span></div>
								        <div class="content"><textarea name="message_<%=commArr(0,Pcount)%>" onfocus="focusMe(this)" onblur="blurMe(this)" onmouseover="overMe(this)" onmouseout="outMe(this)"><%=UnCheckStr(commArr(1,Pcount))%></textarea></div>
								        <div class="reply"><h5>回复内容:<%if len(trim(commArr(5,Pcount)))<1 or IsNull(commArr(5,Pcount)) then response.write "<span class=""tip"">(无回复留言)</span>"%></h5><textarea name="reply_<%=commArr(0,Pcount)%>" onfocus="focusMe(this)" onblur="blurMe(this)" onmouseover="overMe(this)" onmouseout="outMe(this)" onchange="checkMe(this);document.getElementById('edited_<%=commArr(0,Pcount)%>').value=1"><%=UnCheckStr(commArr(5,Pcount))%></textarea><input id="edited_<%=commArr(0,Pcount)%>" name="edited_<%=commArr(0,Pcount)%>" type="hidden" value="0"></div>
								    </div>
							      <%
							       else '评论
							      %>
							        <div class="item"><input type="hidden" name="CommentID" value="<%=commArr(0,Pcount)%>"/>
								        <div class="title"><span class="blogTitle"><a href="article.asp?id=<%=commArr(5,Pcount)%>" target="_blank" title="<%=commArr(6,Pcount)%>"><%=CutStr(commArr(6,Pcount),25)%></a></span><input type="checkbox" name="selectCommentID" value="<%=commArr(0,Pcount)%>|<%=commArr(5,Pcount)%>" onclick="highLight(this)"/><img src="images/icon_quote.gif" alt=""/><b><%=HtmlEncode(commArr(2,Pcount))%></b> <span class="date">[<%=DateToStr(commArr(3,Pcount),"Y-m-d H:I:S")%> | <%=commArr(4,Pcount)%>]</span></div>
								        <div class="content"><textarea name="message_<%=commArr(0,Pcount)%>" onfocus="focusMe(this)" onblur="blurMe(this)" onmouseover="overMe(this)" onmouseout="outMe(this)"><%=UnCheckStr(commArr(1,Pcount))%></textarea></div>
								    </div>
							      <%
								end if
								Pcount=Pcount+1
							loop
							%>
						</div>
						<div style="height:24px;">
						       <input type="button" value="删除所选内容" onclick="DelComment()" class="button" style="margin:0px;margin-bottom:5px;float:right;"/> 
						       <input type="button" value="全选" onclick="checkAll()" class="button" style="margin:0px;margin-bottom:5px;float:right;margin-right:6px"/>
						       <%=saveButton%>
						       <div id="page2" class="pageContent"><%=MultiPage(comm_Num,pSize,CurPage,aUrl,"","")%></div>
						       <script>document.getElementById("page1").innerHTML=document.getElementById("page2").innerHTML</script>
					    </div>
			  </div>
		 </form>
		   <%
		end if
		set blog_Comment=nothing
   end if	
	%>
  </td></tr>
  </table>
<%
'============================================
'      界面设置
'============================================
elseif Request.QueryString("Fmenu")="Skins" Then 
 dim bmID,bMInfo
 Dim blog_module
 dim PluginsFolders,PluginsFolder,Bmodules,Bmodule,tempB,SubItemLen,tempI
 Dim PluginsXML,DBXML,TypeArray
 TypeArray=Array("sidebar","content","function")
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
  <tr>
    <th class="CTitle"><%=categoryTitle%></th>
  </tr>
  <tr>
    <td class="CPanel">
	<div class="SubMenu"><a href="?Fmenu=Skins">设置外观</a> | <a href="?Fmenu=Skins&Smenu=module">设置模块</a> | <a href="?Fmenu=Skins&Smenu=Plugins">已装插件管理</a> | <a href="?Fmenu=Skins&Smenu=PluginsInstall">安装模块插件</a></div>

<%
 if Request.QueryString("Smenu")="module" then
 %>
   <div align="left" style="padding:5px;"><%getMsg%>
   <form action="ConContent.asp" method="post" style="margin:0px">
       <input type="hidden" name="action" value="Skins"/>
       <input type="hidden" name="whatdo" value="UpdateModule"/>
       <input type="hidden" name="DoID" value=""/>
      <table border="0" cellpadding="2" cellspacing="1" class="TablePanel">
        <tr align="center">
          <td width="18" class="TDHead"><nobr>&nbsp;</nobr></td>
          <td width="18" class="TDHead">&nbsp;</td>
          <td width="18" class="TDHead">&nbsp;</td>
          <td width="18" class="TDHead"><nobr>&nbsp;</nobr></td>
          <td width="68" class="TDHead"><nobr>类型</nobr></td>
		  <td width="82" class="TDHead"><nobr>模块标识</nobr></td>
		  <td width="160" class="TDHead"><nobr>模块名称</nobr></td>
          <td width="42" class="TDHead"><nobr>排序</nobr></td>
          <td  class="TDHead"><nobr>模块操作</nobr></td>
          </tr>
<%
 dim blogModule
 set blogModule=conn.execute("select * from blog_module where type<>'function' order by type desc,SortID asc")
 do until blogModule.eof 
%>        <tr id="tr_<%=blogModule("id")%>" align="center" style="background:<%if blogModule("type")="content" then response.write ("#a9c9e9")%>">
          <td><input type="checkbox" name="selectID" value="<%=blogModule("id")%>"/></td>
          <td><%if blogModule("IsHidden") then response.write "<img src=""images/icon_lock.gif"" alt=""隐藏模块""/>" %></td>
	      <td><%if blogModule("IndexOnly") then response.write "<img src=""images/urlInTop.gif"" alt=""只在首页出现""/>" %></td>
          <td><img name="Mimg_<%=blogModule("id")%>" src="images/<%=blogModule("type")%>.gif" width="16" height="16"/></td>
          <td><input type="hidden" name="mID" value="<%=blogModule("id")%>"/>
	          <select name="mType" onchange="document.getElementById('tr_<%=blogModule("id")%>').style.backgroundColor=(this.value=='content')?'#a9c9e9':'';document.images['Mimg_<%=blogModule("id")%>'].src='images/'+this.value+'.gif'" <%if blogModule("IsSystem") then response.write "disabled"%>>
	           <option value="sidebar">侧边模块</option>
	           <option value="content" <%if blogModule("type")="content" then response.write ("selected=""selected""")%>>内容模块</option>
	          </select>
           <%if blogModule("IsSystem") then response.write "<input name=""mType"" type=""hidden"" value="""&blogModule("type")&""">"%>
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
      loop
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
           <option value="sidebar">侧边模块</option>
           <option value="content">内容模块</option>
          </select></td>
          <td><input name="mName" type="text" size="12" class="text" /></td>
          <td><input name="mTitle" type="text" size="24" class="text" /></td>
          <td><input name="mOrder" type="text" size="3" class="text" /></td>
          <td></td>
          </tr>
          </table>
		<div class="SubButton" style="text-align:left;padding:5px;margin:0px">
		<select name="doModule">
			 <option>所选模块可用操作</option>
			 <option>------------------------</option>
			 <option value="dohidden">&nbsp;&nbsp;- 设置隐藏</option>
			 <option value="cancelhidden">&nbsp;&nbsp;- 取消隐藏</option>
			 <option>------------------------</option>
			 <option value="doIndex">&nbsp;&nbsp;- 设置首页独享</option>
			 <option value="cancelIndex">&nbsp;&nbsp;- 取消首页独享</option>
		 </select>
			<input type="submit" name="Submit" value="保存模块" class="button" style="margin-bottom:0px"/> 
     </div>          
               <div style="color:#f00">模块标识是唯一标记.一旦确定就无法修改.系统自带的模块不允许删除,内容模块只在首页有效.<br/><b>ContentList</b> 是系统自带的日志列表模块，不允许做任何修改</div>
    </div>
  </form>
 <%
'========================================================
' 可视化编辑模块HTML代码
'========================================================
  elseif Request.QueryString("Smenu")="editModule" Then
 %>

     <div align="center" style="padding:5px;"><%getMsg%>
   <form action="ConContent.asp" method="post" style="margin:0px" onsubmit="return checkSystem()">
     <%
     bmID=Request.QueryString("miD")
     if IsInteger(bmID)=True then
     set bMInfo=conn.execute("select * from blog_module where id="&bmID)
     if bMInfo.eof or bMInfo.bof then
          session(CookieName&"_ShowMsg")=true
          session(CookieName&"_MsgText")="没找到符合条件的模块!"
          Response.Redirect("ConContent.asp?Fmenu=Skins&Smenu=module")
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
          oFCKeditor.BasePath	= sBasePath
          oFCKeditor.Config("AutoDetectLanguage") = true
          oFCKeditor.Config("DefaultLanguage")    = "zh-cn"
          oFCKeditor.Config("FormatSource") = true
          oFCKeditor.Config("FormatOutput") = true
          oFCKeditor.Config("EnableXHTML") = true
          oFCKeditor.Config("EnableSourceXHTML") = true
          oFCKeditor.Value	= UnCheckStr(bMInfo("HtmlCode"))
          oFCKeditor.Create "HtmlCode"

       end if
      else
     session(CookieName&"_ShowMsg")=true
     session(CookieName&"_MsgText")="你提交了非法字符!"
     Response.Redirect("ConContent.asp?Fmenu=Skins&Smenu=module")
     end if
     %>
	<div class="SubButton">
      <input type="submit" name="Submit" value="保存模块代码" class="button" /> 
      <input type="button" name="Submit" value="返回模块列表" class="button" onclick="location='ConContent.asp?Fmenu=Skins&Smenu=module'"/> 
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
 elseif Request.QueryString("Smenu")="editModuleNormal" Then %>
     <div align="left" style="padding:5px;"><%getMsg%>
   <form action="ConContent.asp" method="post" style="margin:0px" onsubmit="return checkSystem()">
     <%
     bmID=Request.QueryString("Mid")
     if IsInteger(bmID)=True then
     set bMInfo=conn.execute("select * from blog_module where id="&bmID)
     if bMInfo.eof or bMInfo.bof then
          session(CookieName&"_ShowMsg")=true
          session(CookieName&"_MsgText")="没找到符合条件的模块!"
          Response.Redirect("ConContent.asp?Fmenu=Skins&Smenu=module")
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
       end if
      else
     session(CookieName&"_ShowMsg")=true
     session(CookieName&"_MsgText")="你提交了非法字符!"
     Response.Redirect("ConContent.asp?Fmenu=Skins&Smenu=module")
     end if
     %>
	<div class="SubButton" style="text-align:left">
      <input type="submit" name="Submit" value="保存HTML代码" class="button" <%If Not CheckObjInstalled("ADODB.Stream") Then Response.Write("disabled")%>/> 
      <input type="button" name="Submit" value="返回模块列表" class="button" onclick="location='ConContent.asp?Fmenu=Skins&Smenu=module'"/> 
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
 elseif Request.QueryString("Smenu")="PluginsInstall" then 
  Bmodule=getModName
  PluginsFolders=split(getPathList("Plugins")(0),"*")
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
	if CheckObjInstalled(getXMLDOM()) then
	 if CheckObjInstalled("Scripting.FileSystemObject") then
     	set PluginsXML=new PXML
     	for each PluginsFolder in PluginsFolders
     	 PluginsXML.XmlPath="Plugins/"&PluginsFolder&"/install.xml"
     	 PluginsXML.open
     	  if PluginsXML.getError=0 then
     	   if len(PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginName"))>0 and IsvalidPlugins(PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginName")) and (not IsvalidValue(Bmodule,PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginName"))) and IsvalidValue(TypeArray,PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginType")) then
     	 %>		 
             <tr>
               <td align="center"><img src="images/<%=PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginType")%>.gif" width="16" height="16"/></td>
     		   <td align="left"><%=PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginTitle")%></td>
     		   <td align="center"><%=PluginsXML.SelectXmlNodeText("PluginInstall/main/Author")%></td>
               <td align="center"><%=PluginsXML.SelectXmlNodeText("PluginInstall/main/pubDate")%></td>
               <td align="center"><a href="?Fmenu=Skins&Smenu=InstallPlugins&Plugins=<%=PluginsFolder%>">安装此插件</a> | <a href="javascript:alert('<%=PluginsXML.SelectXmlNodeText("PluginInstall/main/About")%>')">关于</a></td>
             </tr>
     	<% 
     	   end If
     	  end if
      	 PluginsXML.CloseXml()
     	next
     else
    response.write ("<tr><td colspan=""6"" align=""center""><div style=""background:#ffffe8;border:1px solid #95801c;padding:3px;text-align:left;"">你的系统不支持 <b>Scripting.FileSystemObject</b> 只能手动输入插件的文件夹名称</div>")
    response.write ("<div style=""text-align:left;padding:3px""><b>插件路径:</b> Plugins / <input id=""SPath"" type=""text"" size=""16"" class=""text"" value=""""/> <input type=""button"" value=""安装插件"" class=""button"" style=""margin-bottom:-2px"" onclick=""if (document.getElementById('SPath').value.length>0) {location='ConContent.asp?Fmenu=Skins&Smenu=InstallPlugins&Plugins='+document.getElementById('SPath').value}else{alert('请输入插件路径!')}""/></div>")
    response.write ("</td></tr>")
     end if
   else
    response.write ("<tr><td colspan=""6"" align=""center""><div style=""background:#ffffe8;border:1px solid #95801c;padding:3px;text-align:left;"">你的系统不支持 <b>"&getXMLDOM()&"</b>，无法使用插件管理功能，请与服务商联系！</div></td></tr>")
   end if
  	%>
      </table>
  <div style="color:#f00;text-align:left">
  此处列出系统找到的合法的PJBlog2插件，安装插件前需要把插件连同其目录一起上传到Plugins文件夹内
  <br/><b>注意:这里只列出没有安装的插件。</b>
  <br/><%If Not CheckObjInstalled("ADODB.Stream") Then %><b>你的服务器不支持</b> <b><a href="http://www.google.com/search?hl=zh-CN&q=ADODB.Stream&btnG=Google+%E6%90%9C%E7%B4%A2&lr=" target="_blank"><b>ADODB.Stream</b></a> 组件,那将意味着大部分插件的无法正常工作</b><%End If%>
  </div>    
</div>
<%


'============================================================
'    安装插件
'============================================================
 Elseif Request.QueryString("Smenu")="InstallPlugins" then 
    InstallPlugins
'============================================================
'    显示已经安装插件
'============================================================
 elseif Request.QueryString("Smenu")="Plugins" then 
  dim Blog_Plugins
  set Blog_Plugins=conn.execute("select * from blog_module where IsInstall=true order by id desc")
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
       loop
       set Blog_Plugins=nothing
       %>
     </table>
       <div style="color:#f00;text-align:left">
       <%If Not CheckObjInstalled("ADODB.Stream") Then %>
	       <b>你的服务器不支持</b> <b><a href="http://www.google.com/search?hl=zh-CN&q=ADODB.Stream&btnG=Google+%E6%90%9C%E7%B4%A2&lr=" target="_blank"><b>ADODB.Stream</b></a> 组件,那将意味着大部分插件的无法正常工作</b>
       <%else%>
	       <input type="button" name="button" value="修复插件" class="button" onclick="FixPlugins()"/>
       <%End If%>
<br/>
 假如插件反安装不成功，请到 <b>数据库与附件-数据库管理</b> 压缩修复数据再反安装
</div>
  <%
  
  
  
'============================================================
'    反安装插件
'============================================================
 elseif Request.QueryString("Smenu")="UnInstallPlugins" then
   Dim UnPlugName,getCateID,DropTable,KeepTable,ModSetTemp1,getMod
   PluginsFolder=CheckStr(Request.QueryString("Plugins"))
   KeepTable=CBool(Request.QueryString("Keep"))
   set PluginsXML=new PXML
   PluginsXML.XmlPath="Plugins/"&PluginsFolder&"/install.xml"
   PluginsXML.open
   if PluginsXML.getError=0 Then
     UnPlugName=PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginName")
     Set ModSetTemp1=New ModSet
     ModSetTemp1.Open UnPlugName
     ModSetTemp1.RemoveApplication
     DropTable=PluginsXML.SelectXmlNodeText("PluginInstall/main/DropTable")
     set getMod=conn.Execute("select CateID from blog_module where name='"&UnPlugName&"'")
     if getMod.eof then
          session(CookieName&"_ShowMsg")=true
          session(CookieName&"_MsgText")="<font color=""#ff0000"">"&UnPlugName&"</font> 无法反安装,数据库没有找到相应的信息!"
          Response.Redirect("ConContent.asp?Fmenu=Skins&Smenu=Plugins")
      else
        getCateID=getMod(0)
        if len(getCateID)>0 then conn.Execute("delete * from blog_Category where cate_ID="&getCateID)
        delPlugins UnPlugName
        if len(DropTable)>0 And KeepTable=False Then conn.Execute("DROP TABLE "&DropTable)
        SubItemLen = int(PluginsXML.GetXmlNodeLength("PluginInstall/SubItem/item"))
        
        for tempI=0 to SubItemLen-1 
          if not PluginsXML.SelectXmlNodeText("PluginInstall/SubItem/item/PluginType")="function" then
             delPlugins UnPlugName&"SubItem"&(tempI+1)
          end If
       Next
         If len(PluginsXML.SelectXmlNodeText("PluginInstall/main/SettingFile"))>0 Then
          if KeepTable=False Then InstallPlugingSetting "",UnPlugName,"delete"
       End If
     end if
   end If
   
   PluginsXML.CloseXml()
   log_module(2)
   CategoryList(2)
   FixPlugins(0)
   session(CookieName&"_ShowMsg")=true
   session(CookieName&"_MsgText")="<font color=""#ff0000"">"&UnPlugName&"</font> 插件反安装完成!"
   Response.Redirect("ConContent.asp?Fmenu=Skins&Smenu=Plugins")
  
'============================================================
'    修复插件
'============================================================  
elseif Request.QueryString("Smenu")="FixPlugins" then
  FixPlugins(1)
'============================================================
'    插件配置
'============================================================
elseif Request.QueryString("Smenu")="PluginsOptions" then
 dim PluginsSetting,LoadSetXML,KeyLen,Si,LoadModSet,SelectTemp
 set PluginsSetting=conn.execute("select top 1 * from blog_module where name='"+checkstr(Request.QueryString("Plugins"))+"'")
 set LoadSetXML=new PXML
 Set LoadModSet=New ModSet
 LoadModSet.Open(PluginsSetting("name"))
 LoadSetXML.XmlPath="Plugins/"&PluginsSetting("InstallFolder")&"/"&PluginsSetting("SettingXML")
 LoadSetXML.Open
   if LoadSetXML.getError=0 then
    KeyLen = int(LoadSetXML.GetXmlNodeLength("PluginOptions/Key"))
    getMsg
    Response.Write ("<div align=""center""><form action=""ConContent.asp"" method=""post"" style=""margin:0px"">")
    Response.Write ("<input type=""hidden"" name=""action"" value=""Skins""/>")
    Response.Write ("<input type=""hidden"" name=""whatdo"" value=""SavePluginsSetting""/>")
    Response.Write ("<input type=""hidden"" name=""PluginsName"" value="""&PluginsSetting("name")&"""/>")
    response.write  "<table border=""0"" cellpadding=""2"" cellspacing=""1"" class=""TablePanel"" style=""margin:6px"">"
     response.write("<tr><td colspan=""2"" align=""left"" style=""background:#e5e5e5;padding:6px""><div style=""font-weight:bold;font-size:14px;"">"&PluginsSetting("title")&" 的基本设置</div></td></tr>")
       For tempI=0 To KeyLen-1
          response.write "<tr><td align=""right"" width=""200"" valign=""top"" style=""padding-top:6px"">"&LoadSetXML.GetAttributes("PluginOptions/Key","description",TempI)&"</td><td width=""300"">"
          if Lcase(LoadSetXML.GetAttributes("PluginOptions/Key","type",TempI))="select" then
             response.write "<select name="""&LoadSetXML.GetAttributes("PluginOptions/Key","name",TempI)&""">"
                for Si=0 to LoadSetXML.SelectXmlNode("PluginOptions/Key",TempI).getElementsByTagName("option").length-1
                     If LoadModSet.getKeyValue(LoadSetXML.GetAttributes("PluginOptions/Key","name",tempI)) = LoadSetXML.SelectXmlNode("PluginOptions/Key",TempI).getElementsByTagName("option").item(Si).attributes(0).value Then SelectTemp="selected"
                    If LoadSetXML.SelectXmlNode("PluginOptions/Key",TempI).getElementsByTagName("option").item(Si).attributes.length>0 Then
                        Response.write "<option "&SelectTemp&" value="""&LoadSetXML.SelectXmlNode("PluginOptions/Key",TempI).getElementsByTagName("option").item(Si).attributes(0).value&""">"&LoadSetXML.SelectXmlNode("PluginOptions/Key",TempI).getElementsByTagName("option").item(Si).text&"</option>"
                     else
                        Response.write "<option "&SelectTemp&">"&LoadSetXML.SelectXmlNode("PluginOptions/Key",TempI).getElementsByTagName("option").item(Si).text&"</option>"
                    end If
                    SelectTemp=""
                Next
             response.write "</select></td></tr>"
          elseif Lcase(LoadSetXML.GetAttributes("PluginOptions/Key","type",TempI))="textarea" Then
             response.write "<textarea name="""&LoadSetXML.GetAttributes("PluginOptions/Key","name",tempI)&""" rows="""&LoadSetXML.GetAttributes("PluginOptions/Key","rows",TempI)&""" cols="""&LoadSetXML.GetAttributes("PluginOptions/Key","cols",TempI)&""">"&LoadModSet.getKeyValue(LoadSetXML.GetAttributes("PluginOptions/Key","name",tempI))&"</textarea></td></tr>"
          Else
             response.write "<input name="""&LoadSetXML.GetAttributes("PluginOptions/Key","name",TempI)&""" type="""&LoadSetXML.GetAttributes("PluginOptions/Key","type",TempI)&""" size="""&LoadSetXML.GetAttributes("PluginOptions/Key","size",TempI)&""" value="""&LoadModSet.getKeyValue(LoadSetXML.GetAttributes("PluginOptions/Key","name",tempI))&"""/></td></tr>"
          end If
       next
    response.write "<tr><td colspan=""2"" align=""center""><input type=""submit"" name=""Submit"" value=""保存设置"" class=""button""/><input type=""button"" value=""放弃返回"" class=""button"" onclick=""history.go(-1)""/></td></tr>"
    response.write "</table>"
    response.write "</form></div>"
   else
    response.write "无法找到配置文件"
   end if
 set LoadSetXML=nothing
 set PluginsSetting=nothing
'============================================================
'    设置外观 
'============================================================
 else
 dim SkinFolders,SkinFolder
  SkinFolders=split(getPathList("skins")(0),"*")
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
	if CheckObjInstalled(getXMLDOM()) and CheckObjInstalled("Scripting.FileSystemObject") then
      dim SkinXML,k,SkinPreview
      k=2
     set SkinXML=new PXML
	 for each SkinFolder in  SkinFolders
      SkinXML.XmlPath="skins/"&SkinFolder&"/skin.xml"
      SkinXML.open
	  if SkinXML.getError=0 then
	  if k/2=int(k/2) then response.write "<tr>"
	  SkinPreview="images/Control/skin.jpg"
	  if FileExist("skins/"&SkinFolder&"/Preview.jpg") then SkinPreview="skins/"&SkinFolder&"/Preview.jpg"
	 %>	
 		<td width="50%" style='border-bottom:1px dotted #ccc'>
 			<div class="<%if Lcase(blog_DefaultSkin)<>Lcase(SkinFolder) then response.write ("un")%>selectskin">
 			<img src="<%=SkinPreview%>" alt="" border="0" class="skinimg"/>
 			  <div class="skinDes">
 			  <div style="height:38px;overflow:hidden"><b style="color:#004000"><%=SkinXML.SelectXmlNodeText("SkinName")%></b></div>
 			  <span style="height:16px;overflow:hidden;cursor:default" title="设计者:<%=SkinXML.SelectXmlNodeText("SkinDesigner")%>"><B>设计者:</B> <%=SkinXML.SelectXmlNodeText("SkinDesigner")%></span><br/>
 			  <B>发布时间:</B> <%=SkinXML.SelectXmlNodeText("pubDate")%><br/></div>
		 <%
		    if Lcase(blog_DefaultSkin)=Lcase(SkinFolder) then 
			  response.write ("<div class=""cskin""><img src=""images/Control/select.gif"" alt="""" border=""0"" /></div>当前界面")
			 else
			  response.write ("<a href=""javascript:setSkin('"&SkinFolder&"','"&SkinXML.SelectXmlNodeText("SkinName")&"')"">设置为当前主题</a>")
			end if
		  %>
		  </div>
 		</td>
	<%
	  end if
	 SkinXML.CloseXml
	 if k/2<>int(k/2) then response.write "</tr>"
	 k=k+1
	next
	 if k/2<>int(k/2) then response.write "</tr>"
	
	set SkinXML=nothing
   else
    response.write ("<tr><td colspan=""2"" align=""center""><div style=""background:#ffffe8;border:1px solid #95801c;padding:3px;text-align:left;"">你的系统不支持 <b>"&getXMLDOM()&"</b> 或 <b>Scripting.FileSystemObject</b> 只能手动输入Skin的文件夹名称</div>")
    response.write ("<div style=""text-align:left;padding:3px""><b>界面路径:</b> Skins / <input id=""SPath"" type=""text"" size=""16"" class=""text"" value="""+blog_DefaultSkin+"""/> <input type=""button"" value=""保存界面"" class=""button"" style=""margin-bottom:-2px"" onclick=""if (document.getElementById('SPath').value.length>0) {setSkin(document.getElementById('SPath').value,document.getElementById('SPath').value)}else{alert('请输入界面路径!')}""/></div>")
    response.write ("</td></tr>")
   end if
	%>
      </table>
</div>
<%end if%>
 </td>
  </tr></table>
<%

elseif Request.QueryString("Fmenu")="SQLFile" then '数据库与文件
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
     if Request.QueryString("Smenu")="Attachments" then
     %>
   <form action="ConContent.asp" method="post" style="margin:0px" onsubmit="if (confirm('是否删除选择的文件或文件夹')) {return true} else {return false}">
   <input type="hidden" name="action" value="Attachments"/>
   <input type="hidden" name="whatdo" value="DelFiles"/>
     <%
     dim AttPath,ArrFolder,Arrfile,ArrFolders,Arrfiles,arrUpFolders,arrUpFolder,TempF
     TempF=""
     AttPath=Request.QueryString("AttPath")
    if len(AttPath)<1 then 
       AttPath="attachments"
     elseif bc(server.mapPath(AttPath),server.mapPath("attachments")) then
       AttPath="attachments"
     end If
     ArrFolders=split(getPathList(AttPath)(0),"*")
     Arrfiles=split(getPathList(AttPath)(1),"*")
     response.write "<div style=""font-weight:bold;font-size:14px;margin:3px;margin-left:0px"">"&AttPath&"</div><div style=""margin:3px;margin-left:0px;"">"
     if AttPath<>"attachments" then
       arrUpFolders=split(AttPath,"/")
       for i=0 to ubound(arrUpFolders)-1
        arrUpFolder=arrUpFolder&TempF&arrUpFolders(i)
        TempF="/"
       next
     end if
     if len(arrUpFolder)>0 then
      response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=""?Fmenu=SQLFile&Smenu=Attachments&AttPath="&arrUpFolder&"""><img border=""0"" src=""images/file/folder_up.gif"" style=""margin:4px 3px -3px 0px""/>返回上级目录</a><br>"
     end if
     for each ArrFolder in ArrFolders
      response.write "<input name=""folders"" type=""checkbox"" value="""&AttPath&"/"&ArrFolder&"""/>&nbsp;<a href=""?Fmenu=SQLFile&Smenu=Attachments&AttPath="&AttPath&"/"&ArrFolder&"""><img border=""0"" src=""images/file/folder.gif"" style=""margin:4px 3px -3px 0px""/>"&ArrFolder&"</a><br>"
     next 
     for each Arrfile in Arrfiles
      response.write "<input name=""Files"" type=""checkbox"" value="""&AttPath&"/"&Arrfile&"""/>&nbsp;<a href="""&AttPath&"/"&Arrfile&""" target=""_blank"">"&getFileIcons(getFileInfo(AttPath&"/"&Arrfile)(1))&Arrfile&"</a>&nbsp;&nbsp;"&getFileInfo(AttPath&"/"&Arrfile)(0)&" | "&getFileInfo(AttPath&"/"&Arrfile)(2)&" | "&getFileInfo(AttPath&"/"&Arrfile)(3)&"<br>"
     next 
     response.write "</div>"
     %>
    <div style="color:#f00">如果文件夹内存在文件，那么该文件夹将无法删除!</div>
	  	<div class="SubButton">
      <input type="button" value="全选" class="button" onclick="checkAll()"/>  <input type="submit" name="Submit" value="删除所选的文件或文件夹" class="button"/> 
     </div>
     </form>
	  <%else%>
	  <b>数据库路径:</b> <%=Server.MapPath(AccessFile)%><br/>
	  <b>数据库大小:</b> <span id="accessSize"><%=getFileInfo(AccessFile)(0)%></span><br/>
	  <b>数据库操作:</b> <a href="?Fmenu=SQLFile&do=Compact">压缩修复</a> | <a href="?Fmenu=SQLFile&do=Backup">备份</a><br/>
	  <%
	  Dim AccessFSO,AccessEngine,AccessSource
'-------------压缩数据库-----------------
	  if Request.QueryString("do")="Compact" then
    	Set AccessFSO=Server.CreateObject("Scripting.FileSystemObject")
      	IF AccessFSO.FileExists(Server.Mappath(AccessFile)) Then
		  Response.Write "<div style='padding:4px 0px 4px 10px;border: 1px dotted #999;margin:2px;background:#ffffee'>"
		  Response.Write "压缩数据库开始，网站暂停一切用户的前台操作...<br/>"
		  Response.Write "关闭数据库操作...<br/>"
		  call CloseDB
		  Application.Lock
		  FreeApplicationMemory
		  Application(CookieName & "_SiteEnable") = 0
		  Application(CookieName & "_SiteDisbleWhy") = "网站暂停中，请稍候几分钟后再来..."
	      Application.UnLock
	      Set AccessEngine = CreateObject("JRO.JetEngine")
	      AccessEngine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath(AccessFile), "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath(AccessFile & ".temp")
	      AccessFSO.CopyFile Server.Mappath(AccessFile & ".temp"),Server.Mappath(AccessFile)
	      AccessFSO.DeleteFile(Server.Mappath(AccessFile & ".temp"))
	      Set AccessFSO = Nothing
	      Set AccessEngine = Nothing
		  Application.Lock
		  Application(CookieName & "_SiteEnable") = 1
		  Application(CookieName & "_SiteDisbleWhy") = ""
		  Application.UnLock
		  Response.write "压缩数据库完成...<br/>"
		  Application.Lock
		  Application(CookieName & "_SiteEnable") = 1
		  Application(CookieName & "_SiteDisbleWhy") = ""
		  Application.UnLock
		  Response.Write "网站恢复正常访问..."
		  Response.Write "</div>"
		  Response.write "<script>document.getElementById('accessSize').innerText='"&getFileInfo(AccessFile)(0)&"'</script>"
      	end if
	  end if
'-------------备份数据库数据库-----------------
	  if Request.QueryString("do")="Backup" then
    	Set AccessFSO=Server.CreateObject("Scripting.FileSystemObject")
      	IF AccessFSO.FileExists(Server.Mappath(AccessFile)) Then
		  Response.Write "<div style='padding:4px 0px 4px 10px;border: 1px dotted #999;margin:2px;background:#ffffee'>"
		  Response.Write "备份数据库开始，网站暂停一切用户的前台操作...<br/>"
		  Response.Write "关闭数据库操作...<br/>"
		  call CloseDB
		  Application.Lock
		  FreeApplicationMemory
		  Application(CookieName & "_SiteEnable") = 0
		  Application(CookieName & "_SiteDisbleWhy") = "网站暂停中，请稍候几分钟后再来..."
	      Application.UnLock
          CopyFiles Server.Mappath(AccessFile),Server.Mappath("backup/Backup_" & DateToStr(Now(),"YmdHIS") & "_" & randomStr(8) &".mbk")
		  Application.Lock
		  Application(CookieName & "_SiteEnable") = 1
		  Application(CookieName & "_SiteDisbleWhy") = ""
		  Application.UnLock
		  Response.write "压缩数据库完成...<br/>"
		  Application.Lock
		  Application(CookieName & "_SiteEnable") = 1
		  Application(CookieName & "_SiteDisbleWhy") = ""
		  Application.UnLock
		  Response.Write "网站恢复正常访问..."
		  Response.Write "</div>"
		  Response.write "<script>document.getElementById('accessSize').innerText='"&getFileInfo(AccessFile)(0)&"'</script>"
      	end if
       Set AccessFSO=Nothing
	  end if
'---------------还原数据库------------
	  if Request.QueryString("do")="Restore" then
	   AccessSource=Request.QueryString("source")
    	Set AccessFSO=Server.CreateObject("Scripting.FileSystemObject")
      	IF AccessFSO.FileExists(Server.Mappath(AccessSource)) Then
		  Response.Write "<div style='padding:4px 0px 4px 10px;border: 1px dotted #999;margin:2px;background:#ffffee'>"
		  Response.Write "还原数据库开始，网站暂停一切用户的前台操作...<br/>"
		  Response.Write "关闭数据库操作...<br/>"
		  call CloseDB
		  Application.Lock
		  FreeApplicationMemory
		  Application(CookieName & "_SiteEnable") = 0
		  Application(CookieName & "_SiteDisbleWhy") = "网站暂停中，请稍候几分钟后再来..."
	      Application.UnLock
          CopyFiles Server.Mappath(AccessFile),Server.Mappath(AccessFile & ".TEMP")
          if DeleteFiles(Server.Mappath(AccessFile)) then response.write ("原数据库删除成功<br/>")
          response.write CopyFiles(Server.Mappath(AccessSource),Server.Mappath(AccessFile))
          if DeleteFiles(Server.MapPath(AccessSource)) then response.write ("数据库备份删除成功<br/>")
	  	  if DeleteFiles(Server.Mappath(AccessFile & ".TEMP")) then response.write ("Temp备份删除成功<br/>")
		  Application.Lock
		  Application(CookieName & "_SiteEnable") = 1
		  Application(CookieName & "_SiteDisbleWhy") = ""
		  Application.UnLock
		  Response.write "数据库还原完成...<br/>"
		  Application.Lock
		  Application(CookieName & "_SiteEnable") = 1
		  Application(CookieName & "_SiteDisbleWhy") = ""
		  Application.UnLock
		  Response.Write "网站恢复正常访问..."
		  Response.Write "</div>"
		  Response.write "<script>document.getElementById('accessSize').innerText='"&getFileInfo(AccessFile)(0)&"'</script>"
      	end if
       Set AccessFSO=Nothing
	  end if
'---------------删除备份数据库------------
	  if Request.QueryString("do")="DelFile" then
	   AccessSource=Request.QueryString("source")
	    Set AccessFSO=Server.CreateObject("Scripting.FileSystemObject")
      	IF AccessFSO.FileExists(Server.Mappath(AccessSource)) Then
		  Response.Write "<div style='padding:4px 0px 4px 10px;border: 1px dotted #999;margin:2px;background:#ffffee'>"
          if DeleteFiles(Server.MapPath(AccessSource)) then response.write ("数据库备份删除成功<br/>")
		  Response.Write "</div>"
      	end if
       Set AccessFSO=Nothing
      end if
	  %>
	  	  <br/><b>数据库备份列表:</b> <br/>
	  <%
	   dim BackUpFiles,BackUpFile
	   BackUpFiles=split(getPathList("backup")(1),"*")
       for each BackUpFile in BackUpFiles
        response.write "<a href=""backup/"&BackUpFile&""" target=""_blank"">"&getFileIcons(getFileInfo("backup/"&BackUpFile)(1))&BackUpFile&"</a>"
        response.write "&nbsp;&nbsp;<a href=""?Fmenu=SQLFile&do=DelFile&source=backup/"&BackUpFile&""" title=""删除备份文件"">删除</a> | <a href=""?Fmenu=SQLFile&do=Restore&source=backup/"&BackUpFile&""" title=""删除备份文件"">还原数据库</a>"
        response.write " | "&getFileInfo("backup/"&BackUpFile)(0)&" | "&getFileInfo("backup/"&BackUpFile)(2)&"<br/>"
       next
       %>
	 <%end if%>
	 </div>
 </td>
  </tr></table>
<%
elseif Request.QueryString("Fmenu")="Members" then '帐户与权限
Dim blog_Status,blog_Statu,StatusItem,blog_Status_Len
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
  <tr>
    <th class="CTitle"><%=categoryTitle%></th>
  </tr>
  <tr>
    <td class="CPanel">
	<div class="SubMenu"><a href="?Fmenu=Members">权限管理</a> | <a href="?Fmenu=Members&Smenu=Users">帐户管理</a></div>
<%
if Request.QueryString("Smenu")="Users" then

%>
<%getMsg%>
 <form action="ConContent.asp" method="post" style="margin:0px">
   <input type="hidden" name="action" value="Members"/>
   <input type="hidden" name="whatdo" value="SaveUserRight"/>
   <input type="hidden" name="DelID" value=""/>
<table border="0" cellpadding="2" cellspacing="1" class="TablePanel" style="margin:5px;">
<%
blog_Status=Application(CookieName&"_blog_rights")
dim FindUser,FindUserFilter
FindUser=Request.QueryString("User")
if len(FindUser)<1 then 
 FindUserFilter=""
else
 FindUserFilter=" AND M.mem_Name='" & FindUser & "'"
end if
 If CheckStr(Request.QueryString("Page"))<>Empty Then
	Curpage=CheckStr(Request.QueryString("Page"))
	If IsInteger(Curpage)=False OR Curpage<0 Then Curpage=1
 Else
	Curpage=1
 End If
 dim bMember,PageCM
 Set bMember=Server.CreateObject("ADODB.RecordSet")
 SQL="SELECT M.*,S.stat_name,S.stat_title FROM blog_Member as M,blog_status as S where M.mem_Status=S.stat_name"&FindUserFilter&" order by mem_RegTime desc"
 bMember.Open SQL,Conn,1,1
 PageCM=0
    response.write ("<tr><td colspan=""8"" style=""border-bottom:1px solid #999;background:#fae1af;height:36px"">&nbsp;用户名&nbsp;<input id=""FindUser"" type=""text"" class=""text"" size=""16""/><input type=""button"" value=""查找用户"" class=""button"" style=""margin-bottom:-2px"" onclick=""location='ConContent.asp?Fmenu=Members&Smenu=Users&User='+escape(document.getElementById('FindUser').value)""/></td></tr>")
 IF not bMember.EOF Then
	bMember.PageSize=30
	bMember.AbsolutePage=CurPage
	Dim bMember_nums
	bMember_nums=bMember.RecordCount
    response.write "<tr><td colspan=""8"" style=""border-bottom:1px solid #999;""><div class=""pageContent"">"&MultiPage(bMember_nums,30,CurPage,"?Fmenu=Members&Smenu=Users&","","float:left")&"</div><div class=""Content-body"" style=""line-height:200%""></td></tr>"

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
        blog_Status_Len=ubound(blog_Status,2)
	   	Do Until bMember.EOF OR PageCM=bMember.PageSize
   %>
        <tr align="center">
          <td nowrap><%=bMember("mem_ID")%>
          <%
          	if bMember("mem_Name") <> memName then
          		response.write "<input type=""hidden"" name=""mem_ID"" value="""&bMember("mem_ID")&"""/>"
          	end if
          %>
		</td>
          <td nowrap align="left"><a href="member.asp?action=view&memName=<%=Server.URLEncode(bMember("mem_Name"))%>" target="_blank"><%=bMember("mem_Name")%></a></td>
          <td nowrap>&nbsp;<span id="RightStr_<%=bMember("mem_ID")%>"><%=bMember("stat_title")%></span>&nbsp;</td>
          <td nowrap>&nbsp;<%=DateToStr(bMember("mem_RegTime"),"Y-m-d")%>&nbsp;</td>
          <td nowrap>&nbsp;<%=DateToStr(bMember("mem_lastVisit"),"Y-m-d H:I A")%>&nbsp;</td>
          <td nowrap>&nbsp;<%=bMember("mem_lastIP")%>&nbsp;</td>
          <td>
	          <select name="mem_Status" onchange="ChValue(this.value,'RightStr_<%=bMember("mem_ID")%>')" <%if bMember("mem_Name") = memName then response.write "disabled"%>><%
				For i=0 to blog_Status_Len
	              if bMember("stat_name")=blog_Status(0,i) then
		                response.write "<option value="""&blog_Status(0,i)&""" selected=""selected"">"&blog_Status(0,i)&"</option>"
	               else
		                response.write "<option value="""&blog_Status(0,i)&""">"&blog_Status(0,i)&"</option>"
	              end if
	           next
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
	    PageCM=PageCM+1
	loop
    bMember.Close
    Set bMember=Nothing
 	else
   response.write ("<tr><td colspan=""8"" align=""center"" >你所查询的用户不存在！</td></tr>")
 end if	 
   %></table>
 	<div class="SubButton">
      <input type="submit" name="Submit" value="保存用户权限" class="button"/> 
     </div>
     </form>
 <script>
  function ChValue(str,obj){
   <%
		   For i=0 to blog_Status_Len
	            response.write "if (str=='"&blog_Status(0,i)&"') {document.getElementById(obj).innerText='"&blog_Status(1,i)&"'};"
           next
   %>
   }
 </script>
 <%
     elseif Request.QueryString("Smenu")="EditRight" then
     dim RightDB
     sql="select * from blog_status where stat_name='" & checkstr(Request.QueryString("id")) & "'"
     set RightDB=conn.execute(sql)
	     if RightDB.eof then 
	      response.write "没找到该权限组.请重新更新blog缓存信息"
	     else

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
	         <tr><td align="right">附件大小</td><td ><input name="UploadSize" type="text" size="20" class="text" title="<%=RightDB("stat_attSize")%>字节" value="<%=RightDB("stat_attSize")%>" style="font-size:11px" onchange="this.title=this.value+' 字节'"/></td></tr>
             <tr><td align="right">附件类型</td><td ><input name="UploadType" type="text" size="50" class="text" title="<%=RightDB("stat_attType")%>" value="<%=RightDB("stat_attType")%>" style="font-size:11px" onchange="this.title=this.value"/></td></tr>
             <tr><td colspan="2" align="center"><input type="submit" name="Submit" value="保存设置" class="button"/><input type="button" value="放弃返回" class="button" onclick="history.go(-1)"/></td></tr>
	         
	     </table>
	     </div>
	    </form>
<%
	     end if
	 else
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
blog_Status=Application(CookieName&"_blog_rights")
blog_Status_Len=ubound(blog_Status,2)

For i=0 to blog_Status_Len
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
	 end if
elseif Request.QueryString("Fmenu")="Link" then '友情链接管理
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
  <tr>
	<th class="CTitle"><%=categoryTitle%></th>
  </tr>
  <tr>
    <td class="CPanel"><form name="filter" action="ConContent.asp" method="post" style="margin:0px">
	<%
	dim disLink,disCount
	 if session(CookieName&"_disLink")="" then session(CookieName&"_disLink")="All"
     if session(CookieName&"_disCount")="" then session(CookieName&"_disCount")="10"
	 disLink=session(CookieName&"_disLink")
	 disCount=session(CookieName&"_disCount")
dim FilterWhere
If CheckStr(Request.QueryString("Page"))<>Empty Then
	Curpage=CheckStr(Request.QueryString("Page"))
	If IsInteger(Curpage)=False OR Curpage<0 Then Curpage=1
Else
	Curpage=1
End If	
	%>
	<div class="SubMenu">过滤器: 
   <input type="hidden" name="action" value="Links"/>
   <input type="hidden" name="whatdo" value="Filter"/>
  	<select name="disLink" onchange="document.forms['filter'].submit()">
	  <option value="All">显示所有链接</option>
	  <option value="Allow" <%if disLink="Allow" then response.write ("selected=""selected""")%>>已通过验证链接</option>
	  <option value="NoAllow" <%if disLink="NoAllow" then response.write ("selected=""selected""")%>>未通过验证链接</option>
	  <option value="Top" <%if disLink="Top" then response.write ("selected=""selected""")%>>置顶链接</option>
	  <option value="NoTop" <%if disLink="NoTop" then response.write ("selected=""selected""")%>>未置顶链接</option>
	</select> 每页显示 	<select name="disCount" onchange="document.forms['filter'].submit()">
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
select case disLink
  case "All"
   FilterWhere=""
  case "Allow"
   FilterWhere=" where link_IsShow=true"
  case "NoAllow"
   FilterWhere=" where link_IsShow=false"
  case "Top"
   FilterWhere=" where link_IsMain=true"
  case "NoTop"
   FilterWhere=" where link_IsMain=false"
  case else
   FilterWhere=""  
end select
 dim bLink
 Set bLink=Server.CreateObject("ADODB.RecordSet")
 SQL="SELECT * FROM blog_Links"&FilterWhere&" order by link_IsShow desc,link_Order desc"
 bLink.Open SQL,Conn,1,1
 IF not bLink.EOF Then
	bLink.PageSize=disCount
	bLink.AbsolutePage=CurPage
	Dim bLink_nums
	bLink_nums=bLink.RecordCount
	Dim MultiPages,PageCount
    response.write "<tr><td colspan=""6"" style=""border-bottom:1px solid #999;""><div class=""pageContent"">"&MultiPage(bLink_nums,disCount,CurPage,"?Fmenu=Link&Smenu=&","","float:left")&"</div><div class=""Content-body"" style=""line-height:200%""></td></tr>"
  end if
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
  IF not bLink.EOF Then
	   	Do Until bLink.EOF OR PageCount=bLink.PageSize
		if not bLink("link_IsShow") then 
   %>
        <tr align="center" bgcolor="#FCF4BC">
          <td><img src="images/slink.gif" alt="没有通过验证链接"/></td>
          <td><input name="LinkID" type="hidden" value="<%=bLink("link_ID")%>"/><input name="LinkName" type="text" size="18" class="text" value="<%=bLink("link_Name")%>"/></td>
          <td><input name="LinkURL" type="text" size="30" class="text" value="<%=bLink("link_URL")%>"/></td>
          <td><input name="LinkLogo" type="text" size="40" class="text" value="<%=bLink("link_Image")%>"/></td>
          <td><input name="LinkOrder" type="text" class="text" size="2" value="<%=bLink("link_Order")%>"/></td>
          <td><a href="#" onclick="ShowLink(<%=bLink("link_ID")%>)" title="通过该链接的验证"><img border="0" src="images/alink.gif" width="16" height="16" style="margin:0px 2px -3px 0px"/>通过</a> <a href="<%=bLink("link_URL")%>" target="_blank" title="查看该链接"><img border="0" src="images/icon_trackback.gif" width="16" height="16" style="margin:0px 2px -3px 0px"/>查看</a> <a href="#" onclick="Dellink(<%=bLink("link_ID")%>)" title="删除该链接"><img border="0" src="images/icon_del.gif" width="16" height="16" style="margin:0px 2px -3px 0px"/>删除</a> </td>
        </tr>
		<%else%>
        <tr align="center">
          <td><%if bLink("link_IsMain") then response.write ("<img src=""images/urlInTop.gif"" alt=""置顶链接""/>") else response.write ("&nbsp;")%></td>
          <td><input name="LinkID" type="hidden" value="<%=bLink("link_ID")%>"/><input name="LinkName" type="text" size="18" class="text" value="<%=bLink("link_Name")%>"/></td>
          <td><input name="LinkURL" type="text" size="30" class="text" value="<%=bLink("link_URL")%>"/></td>
          <td><input name="LinkLogo" type="text" size="40" class="text" value="<%=bLink("link_Image")%>"/></td>
          <td><input name="LinkOrder" type="text" class="text" size="2" value="<%=bLink("link_Order")%>"/></td>
          <td><%if bLink("link_IsMain") then response.write ("<a href=""#"" onclick=""Toplink("&bLink("link_ID")&")"" title=""取消该链接在首页置顶""><img border=""0"" src=""images/ct.gif"" style=""margin:0px 2px -3px 0px""/>取消</a> ") else response.write ("<a href=""#"" onclick=""Toplink("&bLink("link_ID")&")"" title=""把该链接在首页置顶""><img border=""0"" src=""images/it.gif"" style=""margin:0px 2px -3px 0px""/>置顶</a> ")%>
		  <a href="<%=bLink("link_URL")%>" target="_blank" title="查看该链接"><img border="0" src="images/icon_trackback.gif" width="16" height="16" style="margin:0px 2px -3px 0px"/>查看</a> <a href="#" onclick="Dellink(<%=bLink("link_ID")%>)" title="删除该链接"><img border="0" src="images/icon_del.gif" width="16" height="16" style="margin:0px 2px -3px 0px"/>删除</a> </td>
	   </tr>
<%      end if
     	bLink.MoveNext
	    PageCount=PageCount+1
	loop
    bLink.Close
    Set bLink=Nothing
end if%>	   
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
elseif Request.QueryString("Fmenu")="smilies" then '表情与关键字
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
	  dim bKeyWord
      Set bKeyWord=conn.execute("select * from blog_Keywords order by key_ID desc")
	  do until bKeyWord.eof 
	   %>
        <tr align="center">
          <td><input name="SelectKeyWordID" type="checkbox" value="<%=bKeyWord("key_ID")%>"/></td>
          <td><input name="KeyWordID" type="hidden" value="<%=bKeyWord("key_ID")%>"/><input name="KeyWord" type="text" size="18" class="text" value="<%=bKeyWord("key_Text")%>"/></td>
          <td><input name="KeyWordURL" type="text" size="34" class="text" value="<%=bKeyWord("key_URL")%>"/></td>
	   </tr>
	   <%
	   bKeyWord.movenext
	   loop
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
	  dim bSmile
      Set bSmile=conn.execute("select * from blog_Smilies order by sm_ID desc")
	  do until bSmile.eof 
	   %>
	   <tr align="center">
		  <td><input name="selectSmiliesID" type="checkbox" value="<%=bSmile("sm_ID")%>"/></td>
          <td><img src="images/smilies/<%=bSmile("sm_Image")%>" alt="<%=bSmile("sm_Image")%>"/></td>
          <td><input name="smilesID" type="hidden" value="<%=bSmile("sm_ID")%>"/><input name="smiles" type="text" size="14" class="text" value="<%=bSmile("sm_Text")%>"/></td>
          <td><input name="smilesURL" type="text" size="27" class="text" value="<%=bSmile("sm_Image")%>"/></td>
	   </tr>
	   <%
	   bSmile.movenext
	   loop
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
elseif Request.QueryString("Fmenu")="Status" then '服务器配置
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
  <tr>
	<th class="CTitle"><%=categoryTitle%></th>
  </tr>
  <tr>
    <td class="CPanel">
    <div align="left" style="padding:5px;line-height:150%">
     <b>软件版本:</b> PJBlog2 v<%=blog_version%> - <%=DateToStr(blog_UpdateDate,"mdy")%><br/>
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
     
     <b>关键组件:</b> (缺少关键组件的服务器会对PJBlog2运行有一定影响)<br/>
     <b>　－ Scripting.FileSystemObject 组件:</b> <%=DisI(CheckObjInstalled("Scripting.FileSystemObject"))%><br/>
     <b>　－ MSXML2.ServerXMLHTTP 组件:</b> <%=DisI(CheckObjInstalled("MSXML2.ServerXMLHTTP"))%><br/> 
     <b>　－ Microsoft.XMLDOM 组件:</b> <%=DisI(CheckObjInstalled("Microsoft.XMLDOM"))%><br/>
     <b>　－ ADODB.Stream 组件:</b> <%=DisI(CheckObjInstalled("ADODB.Stream"))%><br/> 
     <b>　－ Scripting.Dictionary 组件:</b> <%=DisI(CheckObjInstalled("Scripting.Dictionary"))%><br/>
     <br/>
     
     <b>其他组件: </b>(以下组件不影响PJBlog2运行)<br/>
     <b>　－ Msxml2.ServerXMLHTTP.5.0 组件:</b> <%=DisI(CheckObjInstalled("Msxml2.ServerXMLHTTP.5.0"))%><br/> 
     <b>　－ Msxml2.DOMDocument.5.0 组件:</b> <%=DisI(CheckObjInstalled("Msxml2.DOMDocument.5.0"))%><br/> 
     <b>　－ FileUp.upload 组件:</b> <%=DisI(CheckObjInstalled("FileUp.upload"))%><br/>
     <b>　－ JMail.SMTPMail 组件:</b> <%=DisI(CheckObjInstalled("JMail.SMTPMail"))%><br/>
     <b>　－ GflAx190.GflAx 组件:</b> <%=DisI(CheckObjInstalled("GflAx190.GflAx"))%><br/>
     <b>　－ easymail.Mailsend 组件:</b> <%=DisI(CheckObjInstalled("easymail.Mailsend"))%><br/>
    </div>
</td></tr></table>
<%
elseif Request.QueryString("Fmenu")="Logout" then '退出
  session(CookieName&"_System")=""
  session(CookieName&"_disLink")=""
  session(CookieName&"_disCount")=""
  %>
  <script>try{top.location="default.asp"}catch(e){location="default.asp"}</script>
  <%
elseif Request.QueryString("Fmenu")="welcome" then '欢迎%>
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
		     <b>当前软件版本:</b> PJBlog2 v<%=blog_version%><br/>
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
<%else 

'==========================后台信息处理===============================
dim weblog
Set weblog=Server.CreateObject("ADODB.RecordSet")
'==========================基本信息处理===============================
 if Request.form("action")="General" then 
    if Request.form("whatdo")="General" then 
'--------------------------基本信息处理-------------------
    SQL="SELECT * FROM blog_Info"
    weblog.Open SQL,Conn,1,3
    weblog("blog_Name")=checkURL(CheckStr(Request.form("SiteName")))
    weblog("blog_Title")=checkURL(CheckStr(Request.form("blog_Title")))
    weblog("blog_master")=checkURL(CheckStr(Request.form("blog_master")))
    weblog("blog_email")=checkURL(CheckStr(Request.form("blog_email")))
    
    if right(CheckStr(Request.form("SiteURL")),1)<>"/" then
		weblog("blog_URL")=checkURL(CheckStr(Request.form("SiteURL")))&"/"
     else
		weblog("blog_URL")=checkURL(CheckStr(Request.form("SiteURL")))
    end if
    
    weblog("blog_affiche")=""
    weblog("blog_about")=CheckStr(Request.form("blog_about"))
    weblog("blog_PerPage")=CheckStr(Request.form("blogPerPage"))
    weblog("blog_commPage")=CheckStr(Request.form("blogcommpage"))
    weblog("blog_BookPage")=0 'CheckStr(Request.form("blogBookPage"))
    weblog("blog_commTimerout")=CheckStr(Request.form("blog_commTimerout"))
    weblog("blog_commLength")=CheckStr(Request.form("blog_commLength"))
    if CheckObjInstalled("ADODB.Stream") then
	    if CheckStr(Request.form("blog_postFile"))="1" then weblog("blog_postFile")=1 else weblog("blog_postFile")=0
	else
		weblog("blog_postFile")=0
    end if
    if CheckStr(Request.form("blog_Disregister"))="1" then weblog("blog_Disregister")=1 else weblog("blog_Disregister")=0
    if CheckStr(Request.form("blog_validate"))="1" then weblog("blog_validate")=1 else weblog("blog_validate")=0
    if CheckStr(Request.form("blog_commUBB"))="1" then weblog("blog_commUBB")=1 else weblog("blog_commUBB")=0
    if CheckStr(Request.form("blog_commIMG"))="1" then weblog("blog_commIMG")=1 else weblog("blog_commIMG")=0
    if CheckStr(Request.form("blog_ImgLink"))="1" then weblog("blog_ImgLink")=1 else weblog("blog_ImgLink")=0
    weblog("blog_SplitType")=CBool(CheckStr(Request.form("blog_SplitType")))
    weblog("blog_introChar")=CheckStr(Request.form("blog_introChar"))
    weblog("blog_introLine")=CheckStr(Request.form("blog_introLine"))
    weblog("blog_FilterName")=CheckStr(Request.form("Register_UserNames"))
    weblog("blog_FilterIP")=CheckStr(Request.form("FilterIPs"))
    weblog("blog_DisMod")=CheckStr(Request.form("blog_DisMod"))
    if not IsInteger(Request.form("blog_CountNum")) then
    	weblog("blog_CountNum")=0
    else
	    weblog("blog_CountNum")=Request.form("blog_CountNum")
    end if
    weblog("blog_wapNum")=CheckStr(Request.form("blog_wapNum"))
    if CheckStr(Request.form("blog_wapImg"))="1" then weblog("blog_wapImg")=1 else weblog("blog_wapImg")=0
    if CheckStr(Request.form("blog_wapHTML"))="1" then weblog("blog_wapHTML")=1 else weblog("blog_wapHTML")=0
    if CheckStr(Request.form("blog_wapLogin"))="1" then weblog("blog_wapLogin")=1 else weblog("blog_wapLogin")=0
    if CheckStr(Request.form("blog_wapComment"))="1" then weblog("blog_wapComment")=1 else weblog("blog_wapComment")=0
    if CheckStr(Request.form("blog_wap"))="1" then weblog("blog_wap")=1 else weblog("blog_wap")=0
    if CheckStr(Request.form("blog_wapURL"))="1" then weblog("blog_wapURL")=1 else weblog("blog_wapURL")=0

    Response.Cookies(CookieNameSetting)("ViewType")="" 
    weblog.update
    weblog.close
    getInfo(2)
    if int(Request.form("SiteOpen"))=1 then
	   Application.Lock
	   Application(CookieName & "_SiteEnable") = 1
	   Application(CookieName & "_SiteDisbleWhy") = ""
	   Application.UnLock
	  Else
	   Application.Lock
	   Application(CookieName & "_SiteEnable") = 0
	   Application(CookieName & "_SiteDisbleWhy") = "抱歉!网站暂时关闭!"
	   Application.UnLock
    end if
    session(CookieName&"_ShowMsg") = true
    session(CookieName&"_MsgText") = "基本信息修改成功!"
    Response.Redirect("ConContent.asp?Fmenu=General&Smenu=")
    elseif Request.form("whatdo")="Misc" Then %>
    <!--#include file="common/ubbcode.asp" -->
    <%
     if Request.form("ReBulidArticle")=1 Then
      Dim LoadArticle,LogLen
      LogLen=0
      Set LoadArticle=conn.Execute("SELECT log_ID FROM blog_Content")
      Do Until LoadArticle.eof
        PostArticle LoadArticle("log_ID")
        LogLen=LogLen+1
        LoadArticle.movenext
      Loop
      session(CookieName&"_ShowMsg")=True
      session(CookieName&"_MsgText")=Session(CookieName&"_MsgText")&"共转换 "&LogLen&" 篇日志到文件! "
     End If
     
    if Request.form("ReBulidIndex")=1 Then
	     dim lArticle
		 set lArticle = new ArticleCache
		 lArticle.SaveCache
		 set lArticle = nothing
      session(CookieName&"_ShowMsg")=True
      session(CookieName&"_MsgText")=Session(CookieName&"_MsgText")&"重新输出日志索引! "
    end if 
     
     if Request.form("ReTatol")=1 Then
      dim blog_Content_count,blog_Comment_count,ContentCount,TBCount,Count_Member
      ContentCount=0
      TBCount=conn.execute("select count(*) from blog_Trackback")(0)
      Count_Member=conn.execute("select count(*) from blog_Member")(0)
      conn.execute("update blog_Info set blog_tbNums="&TBCount)
      conn.execute("update blog_Info set blog_MemNums="&Count_Member)
      set blog_Content_count=conn.execute("SELECT log_CateID, Count(log_CateID) FROM blog_Content where log_IsDraft=false GROUP BY log_CateID")
      set blog_Comment_count=conn.execute("SELECT Count(*) FROM blog_Comment")
      do while not blog_Content_count.eof
       ContentCount=ContentCount+blog_Content_count(1)
       conn.execute("update blog_Category set cate_count="&blog_Content_count(1)&" where cate_ID="&blog_Content_count(0))
       blog_Content_count.movenext
      Loop
      conn.execute("update blog_Info set blog_LogNums="&ContentCount)
      conn.execute("update blog_Info set blog_CommNums="&blog_Comment_count(0))
      getInfo(2):CategoryList(2)
      session(CookieName&"_ShowMsg")=True
      session(CookieName&"_MsgText")=session(CookieName&"_MsgText")&"数据统计完成! "
     end if   
     if Request.form("CleanVisitor")=1 then
      conn.execute("delete * from blog_Counter")
      session(CookieName&"_ShowMsg")=true
      session(CookieName&"_MsgText")=session(CookieName&"_MsgText")&"访客数据清除完成! "
     end if  
     if Request.form("ReBulid")=1 then 
      FreeMemory
      session(CookieName&"_ShowMsg")=true
      session(CookieName&"_MsgText")=session(CookieName&"_MsgText")&"缓存重建成功! "
     end If
     Response.Redirect("ConContent.asp?Fmenu=General&Smenu=Misc")
    else
     session(CookieName&"_ShowMsg")=true
     session(CookieName&"_MsgText")="非法提交内容"
     Response.Redirect("ConContent.asp?Fmenu=General&Smenu=")
    end if
'==========================处理日志分类===============================
 elseif Request.form("action")="Categories" then 
'--------------------------处理日志批量移动----------------------------
   if Request.form("whatdo")="move" then 
   dim Cate_source,Cate_target,Cate_source_name,Cate_source_count,Cate_target_name
    Cate_source=int(CheckStr(Request.form("source")))
    Cate_target=int(CheckStr(Request.form("target")))
    if Cate_source=Cate_target then 
     session(CookieName&"_ShowMsg")=true
     session(CookieName&"_MsgText")="源分类和目标分类一致无法移动"
     Response.Redirect("ConContent.asp?Fmenu=Categories&Smenu=move")
    end if
    Cate_source_name=conn.execute("select cate_Name from blog_Category where cate_ID="&Cate_source)(0)
    Cate_source_count=conn.execute("select cate_count from blog_Category where cate_ID="&Cate_source)(0)
    Cate_target_name=conn.execute("select cate_Name from blog_Category where cate_ID="&Cate_target)(0)
    conn.execute ("update blog_Content set log_CateID="&Cate_target&" where log_CateID="&Cate_source)
    conn.execute ("update blog_Category set cate_count=0 where cate_ID="&Cate_source)
    conn.execute ("update blog_Category set cate_count=cate_count+"&Cate_source_count&" where cate_ID="&Cate_target)
    CategoryList(2)
    session(CookieName&"_ShowMsg")=true
    session(CookieName&"_MsgText")="<span style=""color:#f00"">"&Cate_source_name&"</span> 移动到 <span style=""color:#f00"">"&Cate_target_name&"</span> 成功! 批量转移后，请到 <a href=""ConContent.asp?Fmenu=General&Smenu=Misc"" style=""color:#00f"">站点基本设置-初始化数据</a> ,重新生成所有日志到文件"
    Response.Redirect("ConContent.asp?Fmenu=Categories&Smenu=move")
'--------------------------处理日志分类----------------------------
   elseif Request.form("whatdo")="Cate" then 
    '处理存在分类
    dim LCate_ID,LCate_icons,LCate_Name,LCate_Intro,Lcate_URL,Lcate_Order,Lcate_count,LCate_local,Lcate_Secret
    dim NCate_ID,NCate_icons,NCate_Name,NCate_Intro,Ncate_URL,Ncate_Order,NCate_local,Ncate_OutLink,Ncate_Secret
    LCate_ID=split(Request.form("Cate_ID"),", ")
    LCate_Name=split(Request.form("Cate_Name"),", ")
    LCate_icons=split(Request.form("Cate_icons"),", ")
    LCate_Intro=split(Request.form("Cate_Intro"),", ")
    Lcate_URL=split(Request.form("cate_URL"),", ")
    Lcate_Order=split(Request.form("cate_Order"),", ")
    Lcate_count=split(Request.form("cate_count"),", ")
    LCate_local=Split(Request.form("Cate_local"),", ")
    Lcate_Secret=Split(Request.form("cate_Secret"),", ")
    For i=0 To UBound(LCate_Name)
     SQL="SELECT * FROM blog_Category where cate_ID="&int(CheckStr(LCate_ID(i)))
     weblog.Open SQL,Conn,1,3
     weblog("cate_Name")=CheckStr(LCate_Name(i))
     weblog("cate_icon")=CheckStr(LCate_icons(i))
     weblog("Cate_Intro")=CheckStr(LCate_Intro(i))
     if len(trim(Lcate_URL(i)))>1 and int(Lcate_count(i))<1 then 
      weblog("cate_URL")=trim(CheckStr(Lcate_URL(i)))
      weblog("cate_OutLink")=true
     else
      weblog("cate_URL")=trim(CheckStr(Lcate_URL(i)))
      weblog("cate_OutLink")=false
     end if
     weblog("cate_Order")=int(CheckStr(Lcate_Order(i)))
     weblog("Cate_local")=int(CheckStr(LCate_local(i)))
     weblog("cate_Secret")=CBool(CheckStr(Lcate_Secret(i)))
     weblog.update
     weblog.close
    next
    '判断添加新日志
    NCate_Name=trim(CheckStr(Request.form("New_Cate_Name")))
    NCate_icons=CheckStr(Request.form("New_Cate_icons"))
    NCate_Intro=trim(CheckStr(Request.form("New_Cate_Intro")))
    Ncate_URL=trim(CheckStr(Request.form("New_cate_URL")))
    Ncate_Order=CheckStr(Request.form("New_cate_Order"))
    NCate_local=CheckStr(Request.form("New_Cate_local"))
    Ncate_Secret=CheckStr(Request.form("New_Cate_Secret"))
    if len(NCate_Name)>0 then
     if len(Ncate_Order)<1 then Ncate_Order=conn.execute("select count(*) from blog_Category")(0)
     if len(Ncate_URL)>0 then Ncate_OutLink=true else Ncate_OutLink=false
     dim AddCateArray
     AddCateArray=array(array("cate_Name",NCate_Name),array("cate_icon",NCate_icons),array("Cate_Intro",NCate_Intro),array("cate_URL",Ncate_URL),array("cate_OutLink",Ncate_OutLink),array("cate_Order",int(Ncate_Order)),array("Cate_local",NCate_local),Array("Cate_Secret",NCate_Secret))
     if DBQuest("blog_Category",AddCateArray,"insert")=0 then session(CookieName&"_MsgText")="新日志分类添加成功，"
    end if
    FreeMemory
    session(CookieName&"_ShowMsg")=true
    session(CookieName&"_MsgText")=session(CookieName&"_MsgText")&"日志分类更新成功!"
    Response.Redirect("ConContent.asp?Fmenu=Categories&Smenu=")
'--------------------------批量删除日志----------------------------
   elseif Request.form("whatdo")="batdel" then 
     dim CID,Cids,tti,C1,C2
     C1=0
     C2=0
     CID=checkstr(request.form("CID"))
     Cids=split(CID,", ")
     for tti=0 to ubound(Cids)
	     if DeleteLog(Cids(tti))=1 then 
	       C1=C1+1
	      else
	       C2=C2+1
	     end if
     next
    FreeMemory
    session(CookieName&"_ShowMsg")=true
    session(CookieName&"_MsgText")="日志删除状态:"&C1&"篇成功,"&C2&"篇失败! 批量删除后，请到 <a href=""ConContent.asp?Fmenu=General&Smenu=Misc"" style=""color:#00f"">站点基本设置-初始化数据</a> ,重新生成所有日志到文件"
    Response.Redirect("ConContent.asp?Fmenu=Categories&Smenu=del")
'--------------------------删除日志分类----------------------------
   elseif Request.form("whatdo")="DelCate" then 
    Dim DelCate,DelLog,DelID,conCount,comCount,P1,P2
    P1=0
    p2=0
    DelCate=Request.form("DelCate")
    set DelLog=Conn.Execute("select log_ID FROM blog_Content WHERE log_CateID="&DelCate)
    do while not DelLog.eof 
     DelID=DelLog("log_ID")
     if DeleteLog(DelID)=1 then 
       P1=P1+1
      else
       P2=P2+1
     end if
     DelLog.movenext
    loop
    Conn.ExeCute("DELETE * FROM blog_Category WHERE cate_ID="&DelCate)
    FreeMemory
    session(CookieName&"_ShowMsg")=true
    session(CookieName&"_MsgText")=session(CookieName&"_MsgText")&"日志分类删除成功! 日志删除状态:"&P1&"篇成功,"&P2&"篇失败! 批量删除后，请到 <a href=""ConContent.asp?Fmenu=General&Smenu=Misc"" style=""color:#00f"">站点基本设置-初始化数据</a> ,重新生成所有日志到文件"
    Response.Redirect("ConContent.asp?Fmenu=Categories&Smenu=")
   elseif Request.form("whatdo")="Tag" then 
       Dim TagsID,TagName
	   if Request.form("doModule")="DelSelect" then
		    TagsID=split(Request.form("selectTagID"),", ")
		    for i=0 to ubound(TagsID)
				conn.execute("DELETE * from blog_tag where tag_id="&TagsID(i))
			next
		    Tags(2)
		    session(CookieName&"_ShowMsg")=true
		    session(CookieName&"_MsgText")=session(CookieName&"_MsgText")&(ubound(TagsID)+1)&" 个Tag被删除!"
		    if blog_postFile then session(CookieName&"_MsgText")=session(CookieName&"_MsgText")+"由于你使用了静态日志功能，所以删除tag后建议到 <a href='ConContent.asp?Fmenu=General&Smenu=Misc' title='站点基本设置-初始化数据 '>初始化数据</a> 重新生成所有日志一次."
		    Response.Redirect("ConContent.asp?Fmenu=Categories&Smenu=tag")
    	   else
		     TagsID=split(Request.form("TagID"),", ")
		     TagName=split(Request.form("tagName"),", ")
			 for i=0 to ubound(TagsID)
		     if int(TagsID(i))<>-1 then
		        conn.execute("update blog_tag set tag_name='"&CheckStr(TagName(i))&"' where tag_id="&TagsID(i))
		       else
		         if len(trim(CheckStr(TagName(i))))>0 then
		          conn.execute("insert into blog_tag (tag_name,tag_count) values ('"&CheckStr(TagName(i))&"',0)")
		          session(CookieName&"_MsgText")="新Tag添加成功! "
		         end if
		     end if
		    next
		    Tags(2)
		    session(CookieName&"_ShowMsg")=true
		    session(CookieName&"_MsgText")=session(CookieName&"_MsgText")&"Tag保存成功!"
		    Response.Redirect("ConContent.asp?Fmenu=Categories&Smenu=tag")
     end if
   else
    session(CookieName&"_ShowMsg")=true
    session(CookieName&"_MsgText")="非法提交内容!"
    Response.Redirect("ConContent.asp?Fmenu=Categories&Smenu=")
  end if
  
'==========================评论留言处理===============================
 elseif Request.form("action")="Comment" then
 
  dim selCommID,doCommID,doTitle,doRedirect,t1,t2
  selCommID=split(Request.form("selectCommentID"),", ")
  doCommID=split(Request.form("CommentID"),", ")

  if Request.form("doModule")="updateKey" then
	    saveFilterKey Request.form("keyList")
  elseif Request.form("doModule")="updateRegKey" then
  		saveReFilterKey
  elseif  Request.form("doModule")="DelSelect" then
		     for i=0 to ubound(selCommID)
				 if Request.form("whatdo")="trackback" then
				        t1=int(split(selCommID(i),"|")(0)): t2=int(split(selCommID(i),"|")(1))
						conn.execute("UPDATE blog_Content SET log_QuoteNums=log_QuoteNums-1 WHERE log_ID="&t2)
						conn.execute("DELETE * from blog_Trackback where tb_ID="&t1)
						conn.Execute("UPDATE blog_Info Set blog_tbNums=blog_tbNums-1")
						doTitle="引用通告"
					    PostArticle t2
					    doRedirect="trackback"
				 elseif Request.form("whatdo")="msg" then
						conn.execute("DELETE * from blog_book where book_ID="&selCommID(i))
						doTitle="留言"
					    doRedirect="msg"
				 else
				        t1=int(split(selCommID(i),"|")(0)): t2=int(split(selCommID(i),"|")(1))
				        conn.execute("update blog_Content set log_CommNums=log_CommNums-1 where log_ID="&t2)
					    conn.ExeCute("update blog_Info set blog_CommNums=blog_CommNums-1")
						conn.execute("DELETE * from blog_Comment where comm_ID="&t1)
						doTitle="评论"
					    doRedirect=""
					    PostArticle t2
				end if
			next
			
			getInfo(2)
		    NewComment(2)
			Application.Lock
		    Application(CookieName&"_blog_Message")=""
			Application.UnLock
		    session(CookieName&"_ShowMsg")=true
		    session(CookieName&"_MsgText")=(ubound(selCommID)+1)&" 个"&doTitle&"记录被删除!"
 		    Response.Redirect("ConContent.asp?Fmenu=Comment&Smenu="&doRedirect)
   elseif  Request.form("doModule")="Update" then
		     for i=0 to ubound(doCommID)
				if Request.form("whatdo")="msg" then
						if int(Request.form("edited_"&doCommID(i)))=1 then
							conn.execute("UPDATE blog_book SET book_Content='"&checkStr(Request.form("message_"&doCommID(i)))&"',book_replyAuthor='"&memName&"',book_replyTime=#"&DateToStr(now(),"Y-m-d H:I:S")&"#,book_reply='"&checkStr(Request.form("reply_"&doCommID(i)))&"' WHERE book_ID="&doCommID(i))
						else
							conn.execute("UPDATE blog_book SET book_Content='"&checkStr(Request.form("message_"&doCommID(i)))&"',book_replyAuthor='"&memName&"',book_reply='"&checkStr(Request.form("reply_"&doCommID(i)))&"' WHERE book_ID="&doCommID(i))
						end if
						doTitle="留言"
					    doRedirect="msg"
				elseif Request.form("whatdo")="comment" then
						conn.execute("UPDATE blog_Comment SET comm_Content='"&checkStr(Request.form("message_"&doCommID(i)))&"' WHERE comm_ID="&doCommID(i))
						doTitle="评论"
					    doRedirect=""
				end if
		     next
		     
		    NewComment(2)
			Application.Lock
		    Application(CookieName&"_blog_Message")=""
			Application.UnLock
		    session(CookieName&"_ShowMsg")=true
		    session(CookieName&"_MsgText")=doTitle&"记录更新成功!"
 		    Response.Redirect("ConContent.asp?Fmenu=Comment&Smenu="&doRedirect)
   end if
    
'==========================处理帐户和权限信息===============================
 elseif Request.form("action")="Members" then 
'--------------------------处理权限组信息----------------------------
  if Request.form("whatdo")="Group" then
  dim status_name,status_title,Rights,allCount,addCount
  status_name=split(Request.form("status_name"),", ")
  status_title=split(Request.form("status_title"),", ")
  allCount=ubound(status_name)
  dim NS_Name,NS_Title,NS_Code,NS_UpSize,NS_UploadType,tmpNS
  if len(status_name(allCount))>0 then
   NS_Name=CheckStr(status_name(allCount))
   NS_Title=CheckStr(status_title(allCount))
  if not IsValidChars(NS_Name) then 
    session(CookieName&"_ShowMsg")=true
    session(CookieName&"_MsgText")="<span style=""color:#900"">添加新权限失败!权限标识不能为英文或数字以外的字符</span>"
    Response.Redirect("ConContent.asp?Fmenu=Members&Smenu=")
  end if
  set tmpNS=conn.execute("select stat_name from blog_status where stat_name='"&NS_Name&"'")
  if not tmpNS.eof then
    session(CookieName&"_ShowMsg")=true
    session(CookieName&"_MsgText")="<span style=""color:#900"">“"&NS_Name&"”</span> 权限标识已经存在无法添加新分组!"
    Response.Redirect("ConContent.asp?Fmenu=Members&Smenu=")
  end if
   conn.execute("insert into blog_status (stat_name,stat_title,stat_Code,stat_attSize,stat_attType) values ('"&NS_Name&"','"&NS_Title&"','000000000000',0,'')")
   session(CookieName&"_MsgText")="新分组添加成功!"
  end if
  for i=0 to ubound(status_name)-1
  conn.execute("update blog_status set stat_title='"&CheckStr(status_title(i))&"' where stat_name='"&CheckStr(status_name(i))&"'")
  next
    UserRight(2) 
    session(CookieName&"_ShowMsg")=true
    session(CookieName&"_MsgText")=session(CookieName&"_MsgText")&"权限组信息修改成功!"
    Response.Redirect("ConContent.asp?Fmenu=Members&Smenu=")  
'--------------------------处理帐户信息----------------------------
  elseif Request.form("whatdo")="User" then
  
'--------------------------编辑分组信息----------------------------
  elseif Request.form("whatdo")="EditGroup" then
    dim EditGroup,AddArticle,EditArticle,DelArticle,AddComment,DelComment,ShowHiddenCate,IsAdmin,CanUpload,UploadSize,UploadType,Group_title,SCode
    EditGroup=CheckStr(Request.form("status_name"))
    AddArticle=CheckStr(Request.form("AddArticle"))
    EditArticle=CheckStr(Request.form("EditArticle"))
    DelArticle=CheckStr(Request.form("DelArticle"))
    AddComment=CheckStr(Request.form("AddComment"))
    DelComment=CheckStr(Request.form("DelComment"))
    ShowHiddenCate=CheckStr(Request.form("ShowHiddenCate"))
    IsAdmin=CheckStr(Request.form("IsAdmin"))
    CanUpload=CheckStr(Request.form("CanUpload"))
    UploadSize=CheckStr(Request.form("UploadSize"))
    UploadType=CheckStr(Request.form("UploadType"))
    Group_title=CheckStr(Request.form("status_title"))
    SCode=AddArticle & EditArticle & DelArticle &_
          AddComment & DelComment & CanUpload & IsAdmin & ShowHiddenCate
    conn.execute("update blog_status set stat_title='"&Group_title&"',stat_code='"&SCode&"',stat_attSize="&UploadSize&",stat_attType='"&UploadType&"' where stat_name='"&EditGroup&"'")
    UserRight(2)
    session(CookieName&"_ShowMsg")=true
    session(CookieName&"_MsgText")="<span style=""color:#900"">“"&EditGroup&"”</span>权限分组 编辑成功!"
    Response.Redirect("ConContent.asp?Fmenu=Members&Smenu=EditRight&id="&EditGroup)
'--------------------------删除分组信息----------------------------
  elseif Request.form("whatdo")="DelGroup" then
   dim DelGroup
   DelGroup=CheckStr(Request.form("DelGroup"))
   if lcase(DelGroup)<>"supadmin" and lcase(DelGroup)<>"member" and lcase(DelGroup)<>"guest" then
    conn.execute ("update blog_Member set mem_Status='Member' where mem_Status='"&DelGroup&"'")
    Conn.ExeCute("DELETE * FROM blog_status WHERE stat_name='"&DelGroup&"'")
    UserRight(2)
    session(CookieName&"_ShowMsg")=true
    session(CookieName&"_MsgText")="<span style=""color:#900"">“"&DelGroup&"”</span>权限分组 删除成功!"
    Response.Redirect("ConContent.asp?Fmenu=Members&Smenu=")
   else
    session(CookieName&"_ShowMsg")=true
    session(CookieName&"_MsgText")="特殊分组无法删除!"
    Response.Redirect("ConContent.asp?Fmenu=Members&Smenu=")
   end if
'--------------------------保存用户权限----------------------------
  elseif Request.form("whatdo")="SaveUserRight" then
	for i=1 to Request.form("mem_ID").count
      conn.execute("update blog_Member set mem_Status='"&Request.form("mem_Status").item(i)&"' where mem_ID="&Request.form("mem_ID").item(i))
    next
    
    session(CookieName&"_ShowMsg") = true
    session(CookieName&"_MsgText") = "用户权限设置成功!"
    Response.Redirect("ConContent.asp?Fmenu=Members&Smenu=Users")
'--------------------------删除用户----------------------------
  elseif Request.form("whatdo")="DelUser" then
    dim DelUserID,DelUserName,blogmemberNum, DelUserStatus
    DelUserID=Request.form("DelID")
        blogmemberNum=conn.execute("select count(mem_ID) from blog_Member where mem_Status='SupAdmin'")(0)        
        
        DelUserStatus=conn.execute("select mem_Status from blog_Member where mem_ID="&DelUserID)(0)
            if ((blogmemberNum = 1) and (DelUserStatus = "SupAdmin")) then 
                        session(CookieName&"_ShowMsg")=true
                   session(CookieName&"_MsgText")="不能删除仅有的管理员权限!"
                   Response.Redirect("ConContent.asp?Fmenu=Members&Smenu=Users")
                else 
                DelUserName=conn.execute("select mem_Name from blog_Member where mem_ID="&DelUserID)(0)
                    conn.execute("delete * from blog_Member where mem_ID="&DelUserID)
                        Conn.ExeCute("UPDATE blog_Info SET blog_MemNums=blog_MemNums-1")
                        getInfo(2)
                    session(CookieName&"_ShowMsg")=true
                    session(CookieName&"_MsgText")="<span style=""color:#900"">“"&DelUserName&"”</span> 删除成功!"
                    Response.Redirect("ConContent.asp?Fmenu=Members&Smenu=Users")        
                  end if   
   else
    session(CookieName&"_ShowMsg")=true
    session(CookieName&"_MsgText")="非法提交内容!"
    Response.Redirect("ConContent.asp?Fmenu=Members&Smenu=")
  end if
'==========================友情链接管理===============================
 elseif Request.form("action")="Links" then 
    dim LinkID,LinkName,LinkURL,LinkLogo,LinkOrder,LinkMain
'--------------------------友情链接过滤----------------------------
   if Request.form("whatdo")="Filter" then
     session(CookieName&"_disLink")=CheckStr(Request.form("disLink"))
     session(CookieName&"_disCount")=CheckStr(Request.form("disCount"))
     Response.Redirect("ConContent.asp?Fmenu=Link&Smenu=")
'--------------------------保存友情链接----------------------------
   elseif Request.form("whatdo")="SaveLink" then
    dim TLinkName,TLinkURL,TLinkLogo,TLinkOrder
    LinkID=split(Request.form("LinkID"),", ")
    LinkName=split(Request.form("LinkName"),", ")
    LinkURL=split(Request.form("LinkURL"),", ")
    LinkLogo=split(Request.form("LinkLogo"),", ")
    LinkOrder=split(Request.form("LinkOrder"),", ")
    for i=0 to ubound(LinkID)
        if ubound(LinkName)<0 then TLinkName="未知" else TLinkName=LinkName(i)
        if ubound(LinkURL)<0 then TLinkURL="http://" else TLinkURL=LinkURL(i)
        if ubound(LinkLogo)<0 then TLinkLogo="" else TLinkLogo=LinkLogo(i)
        if ubound(LinkOrder)<0 then TLinkOrder="0" else TLinkOrder=LinkOrder(i)
        conn.execute("update blog_Links set link_Name='"&CheckStr(TLinkName)&"',link_URL='"&CheckStr(TLinkURL)&"',link_Image='"&CheckStr(TLinkLogo)&"',link_Order='"&CheckStr(TLinkOrder)&"' where link_ID="&LinkID(i))
    next
    LinkID=Request.form("new_LinkID")
    LinkName=Request.form("new_LinkName")
    LinkURL=Request.form("new_LinkURL")
    LinkLogo=Request.form("new_LinkLogo")
    LinkOrder=Request.form("new_LinkOrder")
       if len(LinkOrder)<1 then LinkOrder=conn.execute("select count(*) from blog_Links")(0)
       if len(trim(CheckStr(LinkName)))>0 then
          conn.execute("insert into blog_Links (link_Name,link_URL,link_Image,link_Order,link_IsShow) values ('"&CheckStr(LinkName)&"','"&CheckStr(LinkURL)&"','"&CheckStr(LinkLogo)&"','"&CheckStr(LinkOrder)&"',true)")
          session(CookieName&"_MsgText")="新友情链接添加成功! "
        end if
    Bloglinks(2)
    PostLink
    session(CookieName&"_ShowMsg")=true
    session(CookieName&"_MsgText")=session(CookieName&"_MsgText")&"保存链接成功!"
    Response.Redirect("ConContent.asp?Fmenu=Link&Smenu=&page="&Request.form("page"))
'--------------------------通过友情链接----------------------------
   elseif Request.form("whatdo")="ShowLink" then
    conn.execute ("update blog_Links set link_IsShow=true where link_ID="&CheckStr(Request.form("ALinkID")))
    LinkName=conn.execute ("select link_Name from blog_Links where link_ID="&CheckStr(Request.form("ALinkID")))(0)
    session(CookieName&"_ShowMsg")=true
    session(CookieName&"_MsgText")=session(CookieName&"_MsgText")&"<span style=""color:#900"">“"&LinkName&"”</span> 通过验证!"
    Bloglinks(2)
    PostLink
    Response.Redirect("ConContent.asp?Fmenu=Link&Smenu=&page="&Request.form("page"))
'--------------------------置顶友情链接----------------------------
   elseif Request.form("whatdo")="TopLink" then
    conn.execute ("update blog_Links set link_IsMain=not link_IsMain where link_ID="&CheckStr(Request.form("ALinkID")))
    LinkName=conn.execute ("select link_Name from blog_Links where link_ID="&CheckStr(Request.form("ALinkID")))(0)
    LinkMain=conn.execute ("select link_IsMain from blog_Links where link_ID="&CheckStr(Request.form("ALinkID")))(0)
    session(CookieName&"_ShowMsg")=true
    if LinkMain then
      session(CookieName&"_MsgText")=session(CookieName&"_MsgText")&"<span style=""color:#900"">“"&LinkName&"”</span> 置顶成功!"
     else
      session(CookieName&"_MsgText")=session(CookieName&"_MsgText")&"<span style=""color:#900"">“"&LinkName&"”</span> 取消首页置顶!"
    end if
    Bloglinks(2)
    PostLink
    Response.Redirect("ConContent.asp?Fmenu=Link&Smenu=&page="&Request.form("page"))
'--------------------------删除友情链接----------------------------
   elseif Request.form("whatdo")="DelLink" then
    LinkName=conn.execute ("select link_Name from blog_Links where link_ID="&CheckStr(Request.form("ALinkID")))(0)
    conn.execute ("DELETE * from blog_Links where link_ID="&CheckStr(Request.form("ALinkID")))
    Session(CookieName&"_ShowMsg")=true
    Session(CookieName&"_MsgText")=session(CookieName&"_MsgText")&"<span style=""color:#900"">“"&LinkName&"”</span> 删除成功!"
    Bloglinks(2)
    PostLink
    Response.Redirect("ConContent.asp?Fmenu=Link&Smenu=&page="&Request.form("page"))
   end if
'==========================表情和关键字===============================
 elseif Request.form("action")="smilies" then 
 dim smilesID,smiles,smilesURL
 dim KeyWordID,KeyWord,KeyWordURL
'--------------------------处理表情符号----------------------------
   if Request.form("whatdo")="smilies" then
	   if Request.form("doModule")="DelSelect" then
		    smilesID=split(Request.form("selectSmiliesID"),", ")
		    for i=0 to ubound(smilesID)
				conn.execute("DELETE * from blog_Smilies where sm_ID="&smilesID(i))
			next
		    Smilies(2)
		    session(CookieName&"_ShowMsg")=true
		    session(CookieName&"_MsgText")=session(CookieName&"_MsgText")&(ubound(smilesID)+1)&" 个表情被删除!"
		    Response.Redirect("ConContent.asp?Fmenu=smilies&Smenu=")
    	   else
		     smilesID=split(Request.form("smilesID"),", ")
		     smiles=split(Request.form("smiles"),", ")
		     smilesURL=split(Request.form("smilesURL"),", ")
			 for i=0 to ubound(smilesID)
		     if int(smilesID(i))<>-1 then
		        conn.execute("update blog_Smilies set sm_Text='"&CheckStr(smiles(i))&"',sm_Image='"&CheckStr(smilesURL(i))&"' where sm_ID="&smilesID(i))
		       else
		         if len(trim(CheckStr(smiles(i))))>0 then
		          conn.execute("insert into blog_Smilies (sm_Text,sm_Image) values ('"&CheckStr(smiles(i))&"','"&CheckStr(smilesURL(i))&"')")
		          session(CookieName&"_MsgText")="新表情添加成功! "
		         end if
		     end if
		    next
		    Smilies(2)
		    session(CookieName&"_ShowMsg")=true
		    session(CookieName&"_MsgText")=session(CookieName&"_MsgText")&"表情保存成功!"
		    Response.Redirect("ConContent.asp?Fmenu=smilies&Smenu=")
     end if
   Elseif Request.form("whatdo")="KeyWord" then
	if Request.form("doModule")="DelSelect" then
		KeyWordID=split(Request.form("SelectKeyWordID"),", ")
		for i=0 to ubound(KeyWordID)
		conn.execute("DELETE * from blog_Keywords where key_ID="&KeyWordID(i))
		next
		Keywords(2)
	    session(CookieName&"_ShowMsg")=true
	    session(CookieName&"_MsgText")=session(CookieName&"_MsgText")&(ubound(KeyWordID)+1)&"关键字被删除!"
	    Response.Redirect("ConContent.asp?Fmenu=smilies&Smenu=KeyWord")
	 else
	     KeyWordID=split(Request.form("KeyWordID"),", ")
	     KeyWord=split(Request.form("KeyWord"),", ")
	     KeyWordURL=split(Request.form("KeyWordURL"),", ")
		 for i=0 to ubound(KeyWordID)
	     if int(KeyWordID(i))<>-1 then
	        conn.execute("update blog_Keywords set key_Text='"&CheckStr(KeyWord(i))&"',key_URL='"&CheckStr(KeyWordURL(i))&"' where key_ID="&KeyWordID(i))
	       else
	         if len(trim(CheckStr(KeyWord(i))))>0 then
	          conn.execute("insert into blog_Keywords (key_Text,key_URL) values ('"&CheckStr(KeyWord(i))&"','"&CheckStr(KeyWordURL(i))&"')")
	          session(CookieName&"_MsgText")="新关键字添加成功! "
	         end if
	     end if
	    next
		Keywords(2)
	    session(CookieName&"_ShowMsg")=true
	    session(CookieName&"_MsgText")=session(CookieName&"_MsgText")&"关键字保存成功!"
	    Response.Redirect("ConContent.asp?Fmenu=smilies&Smenu=KeyWord")
    end if
   else
    session(CookieName&"_ShowMsg")=true
    session(CookieName&"_MsgText")="非法提交内容!"
    Response.Redirect("ConContent.asp?Fmenu=smilies&Smenu=")
   end if
   
'==========================设置界面和模版===============================
 elseif Request.form("action")="Skins" then 
 dim skinpath,Skinname,moduleID,moduleName,moduleType,moduleTitle,moduleHidden,moduleTop,moduleOrder,moduleHtmlCode,mOrder
'--------------------------设置默认界面----------------------------
   if Request.form("whatdo")="setDefaultSkin" then
    skinpath=CheckStr(Request.form("SkinPath"))
    Skinname=CheckStr(Request.form("SkinName"))
    conn.execute("update blog_Info set blog_DefaultSkin='"&skinpath&"'")
    getInfo(2)
    session(CookieName&"_ShowMsg")=true
    session(CookieName&"_MsgText")="<span style=""color:#900"">“"&Skinname&"”</span> 设置为当前默认界面!"
    Response.Redirect("ConContent.asp?Fmenu=Skins&Smenu=")
'--------------------------保存模块设置----------------------------
   elseif Request.form("whatdo")="UpdateModule" then
    Dim selectID,doModule
    selectID=split(Request.form("selectID"),", ")
    doModule=Request.form("doModule")
    moduleID=split(Request.form("mID"),", ")
    moduleName=split(Request.form("mName"),", ")
    moduleType=split(Request.form("mType"),", ")
    moduleTitle=split(Request.form("mTitle"),", ")
    moduleHidden=split(Request.form("mHidden"),", ")
    moduleTop=split(Request.form("mTop"),", ")
    moduleOrder=split(Request.form("mOrder"),", ")
    ',IsHidden="&CBool(moduleHidden(i))&",IndexOnly="&CBool(moduleTop(i))&"
	 for i=0 to ubound(moduleID)
     if int(moduleID(i))<>-1 then
      if not conn.execute ("select IsSystem from blog_module where id="&moduleID(i))(0) then
        conn.execute("update blog_module set title='"&CheckStr(moduleTitle(i))&"',type='"&CheckStr(moduleType(i))&"',SortID="&CheckStr(moduleOrder(i))&" where id="&moduleID(i))
        else
        mOrder=CheckStr(moduleOrder(i))
        if CheckStr(moduleName(i))="ContentList" then mOrder=0
        conn.execute("update blog_module set title='"&CheckStr(moduleTitle(i))&"',SortID="&mOrder&" where id="&moduleID(i))
       end if
       else
         if len(trim(CheckStr(moduleName(i))))>0 then
			  if not conn.execute("select name from blog_module where name='"&CheckStr(moduleName(i))&"'").eof then
			    session(CookieName&"_ShowMsg")=true
			    session(CookieName&"_MsgText")="<span style=""color:#900"">“"&CheckStr(moduleName(i))&"”</span> 模块标识已经存在无法添加新模块!"
			    Response.Redirect("ConContent.asp?Fmenu=Skins&Smenu=module")
			  end if
			  if not IsValidChars(CheckStr(moduleName(i))) then 
			    session(CookieName&"_ShowMsg")=true
			    session(CookieName&"_MsgText")="<span style=""color:#900"">添加新模块失败!权限标识不能为英文或数字以外的字符</span>"
			    Response.Redirect("ConContent.asp?Fmenu=Skins&Smenu=module")
			  end if
			  if len(CheckStr(moduleOrder(i)))<1 then 
				    mOrder=conn.execute("select count(id) from blog_module")(0)
			   else
				     if IsInteger(CheckStr(moduleOrder(i)))=false then 
					       session(CookieName&"_ShowMsg")=true
					       session(CookieName&"_MsgText")="输入非法，添加失败!"
					       Response.Redirect("ConContent.asp?Fmenu=Skins&Smenu=module")
				     end if
			    mOrder=CheckStr(moduleOrder(i))
			  end if
          conn.execute("insert into blog_module (name,title,type,IsHidden,IndexOnly,SortID) values ('"&CheckStr(moduleName(i))&"','"&CheckStr(moduleTitle(i))&"','"&CheckStr(moduleType(i))&"',false,false,"&mOrder&")")
          session(CookieName&"_MsgText")="新模块添加成功! "
         end if
     end if
    next
    for i=0 to ubound(selectID)
	    Select case doModule
	      Case "dohidden":
	        conn.execute("update blog_module set IsHidden=true where id="&selectID(i))
	      Case "cancelhidden":
	        conn.execute("update blog_module set IsHidden=false where id="&selectID(i))
	      Case "doIndex":
	        conn.execute("update blog_module set IndexOnly=true where id="&selectID(i))
	      Case "cancelIndex":
	        conn.execute("update blog_module set IndexOnly=false where id="&selectID(i))
	    end Select
    next
    log_module(2)
    session(CookieName&"_ShowMsg")=true
    session(CookieName&"_MsgText")=session(CookieName&"_MsgText")&"模块保存成功!"
    Response.Redirect("ConContent.asp?Fmenu=Skins&Smenu=module")
    
'--------------------------编辑模块HTML代码----------------------------
   elseif Request.form("whatdo")="editModule" Then
    moduleID=Request.form("DoID")
    moduleName=Request.form("DoName")
    moduleHtmlCode=ClearHTML(CheckStr(request.form("HtmlCode")))
    SavehtmlCode moduleHtmlCode,moduleID
    log_module(2)
    session(CookieName&"_ShowMsg")=True
    session(CookieName&"_MsgText")="<span style=""color:#900"">“"&moduleName&"”</span> 代码编辑成功!"
    if Request.form("editType")="normal" then
		    Response.Redirect("ConContent.asp?Fmenu=Skins&Smenu=editModuleNormal&miD="&moduleID)
     else
		    Response.Redirect("ConContent.asp?Fmenu=Skins&Smenu=editModule&miD="&moduleID)
    end if
   
'-------------------------------删除模块------------------------------
   elseif Request.form("whatdo")="delModule" Then
    moduleID=Request.form("DoID")
    if conn.execute("select isSystem from blog_module where id="&moduleID)(0) Then
      session(CookieName&"_ShowMsg")=True
      session(CookieName&"_MsgText")="<span style=""color:#900"">“"&conn.execute("select title from blog_module where id="&moduleID)(0)&"”</span> 是内置模块无法删除！"
      Response.Redirect("ConContent.asp?Fmenu=Skins&Smenu=module")
    Else
      moduleName=conn.execute("select title from blog_module where id="&moduleID)(0)
      delModule moduleID
      session(CookieName&"_ShowMsg")=True
      session(CookieName&"_MsgText")="<span style=""color:#900"">“"&moduleName&"”</span> 删除成功！"
      log_module(2)
      Response.Redirect("ConContent.asp?Fmenu=Skins&Smenu=module")
    end If
'-------------------------------保存插件配置------------------------------
   elseif Request.form("whatdo")="SavePluginsSetting" Then
    Dim GetPlugName,GetPlugSetItems,GetPlugSetItemName,GetPlugSetItemValue
    GetPlugName=Request.Form("PluginsName")
    Set GetPlugSetItems=conn.Execute ("Select * from blog_ModSetting where set_ModName='"&GetPlugName&"'")
    Do Until GetPlugSetItems.eof 
     GetPlugSetItemName=GetPlugSetItems("set_KeyName")
     GetPlugSetItemValue=checkstr(Request.Form(GetPlugSetItemName))
     conn.Execute ("update blog_ModSetting Set set_KeyValue='"&GetPlugSetItemValue&"' where set_ModName='"&GetPlugName&"'and set_KeyName='"&GetPlugSetItemName&"'")
     GetPlugSetItems.movenext
    Loop
    Dim ModSetTemp2
     Set ModSetTemp2=New ModSet
     ModSetTemp2.Open GetPlugName
     ModSetTemp2.ReLoad()
      session(CookieName&"_ShowMsg")=True
      session(CookieName&"_MsgText")="<span style=""color:#900"">“"&GetPlugName&"”</span> 设置保存成功！"
      Response.Redirect("ConContent.asp?Fmenu=Skins&Smenu=PluginsOptions&Plugins="&GetPlugName)
   end If
'==========================附件管理===============================
 elseif Request.form("action")="Attachments" then 
'-------------------------------删除模块------------------------------
   if Request.form("whatdo")="DelFiles" then
    dim getFolders,getFiles,getFolder,getFile,getFolderCount,getFileCount
    Dim FSODel
    Set FSODel=Server.CreateObject("Scripting.FileSystemObject")
    getFolders=split(Request.form("folders"),", ")
    getFiles=split(Request.form("Files"),", ")
    getFolderCount=0
    getFileCount=0
    for each getFolder in getFolders
     if len(getPathList(getFolder)(1))>0 then
       session(CookieName&"_ShowMsg")=true
       session(CookieName&"_MsgText")="<span style=""color:#900"">“"&getFolder&"”</span> 文件夹内含有文件，无法删除!"
       Response.Redirect("ConContent.asp?Fmenu=SQLFile&Smenu=Attachments")
     end if
     if FSODel.FolderExists(Server.MapPath(getFolder)) then
      FSODel.DeleteFolder Server.MapPath(getFolder),true
      getFolderCount=getFolderCount+1
     end if
    next
    for each getFile in getFiles
     if FSODel.FileExists(Server.MapPath(getFile)) then
      FSODel.DeleteFile Server.MapPath(getFile),true
      getFileCount=getFileCount+1
     end if
    next
    session(CookieName&"_ShowMsg")=true
    session(CookieName&"_MsgText")="有 <span style=""color:#900"">"&getFileCount&" 文件, "&getFolderCount&" 个文件夹</span> 被删除!"
    Response.Redirect("ConContent.asp?Fmenu=SQLFile&Smenu=Attachments")
   end if
  else'登录欢迎
  
 end if

'----------------------End if--------------------
end if
%>
</div>
</body>
</html>
<%

Else
 Response.Redirect("default.asp")
end if
%>