<%
'=================================
' 站点基本设置
'=================================
Public Sub c_ceneral
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
	<tr>
    	<th class="CTitle"><%=control.categoryTitle%></th>
  	</tr>
  	<tr>
    	<td class="CPanel">
            <div class="SubMenu">
                <a href="?Fmenu=General">设置基本信息</a> | 
                <a href="?Fmenu=General&Smenu=visitors">查看访客记录</a> | 
                <a href="?Fmenu=General&Smenu=Misc">初始化数据</a> | 
                <a href="?Fmenu=General&Smenu=clear">清理服务器缓存</a>
            </div>
<%
	If Request.QueryString("Smenu") = "visitors" Then
	ElseIf Request.QueryString("Smenu") = "clear" Then
		Response.Write "<div style='padding:4px 0px 4px 10px;border: 1px dotted #999;margin:2px;background:#ffffee'>"
		Application.Lock
		Control.FreeApplicationMemory
		Application.UnLock
		Response.Write "<br/><span><b style='color:#040'>缓存清理完毕...	</b></span>"
		Response.Write "</div>"
	Else
' ****************************************************************
'	设置基本信息
' ****************************************************************
%>
			<form action="../pjblog.logic/control/log_general.asp?action=SaveGeneral" method="post">
            	<%Control.getMsg%>
                <fieldset>
                	<legend> 站点基本信息</legend>
                    <div align="left">
                    	<table border="0" cellpadding="2" cellspacing="1">
                            <tr>
                            	<td width="180"><div align="right"> BLOG 名称 </div></td>
                              	<td align="left"><input name="SiteName" type="text" size="30" class="text" value="<%=SiteName%>"/></td>
                                <td class="shuom">将显示在网页的标题处</td>
                            </tr>
                            <tr>
		          				<td width="180"><div align="right"> BLOG 副标题 </div></td>
		          				<td align="left"><input name="blog_Title" type="text" size="50" class="text" value="<%=blog_Title%>"/></td>
                                <td class="shuom">将显示在页面的LOGO处</td>
		        			</tr>
							<tr>
		          				<td width="180"><div align="right"> 站长昵称 </div></td>
		          				<td align="left"><input name="blog_master" type="text" size="10" class="text" value="<%=blog_master%>" maxlength="10"/></td>
                                <td class="shuom"></td>
		        			</tr>
                            <tr>
		          				<td width="180"><div align="right"> BLOG 地址</div></td>
		          				<td align="left"><input name="SiteURL" type="text" size="50" class="text" value="<%=SiteURL%>" onblur="if (this.value.reg(REGEXP.REG_WEBURL)){$('checkweburl').innerHTML = '<font color=green>博客地址格式正确</font>';}else{$('checkweburl').innerHTML = '<font color=red>您所填写的博客地址格式不正确!</font>';}"/></td>
                                <td class="shuom">关系到<strong>RSS</strong>地址的可读性和静态页面<strong>CSS</strong>的正确加载<span id="checkweburl" class="Demo"></span></td>
		        			</tr>
                            <tr>
                                <td width="180"><div align="right"> 站长邮件地址 </div></td>
                                <td align="left"><input name="blog_email" type="text" size="50" class="text" value="<%=blog_email%>" onblur="if (this.value.reg(REGEXP.REG_EMAIL)){$('checkemail').innerHTML = '<font color=green>邮箱格式正确</font>';}else{$('checkemail').innerHTML = '<font color=red>您所填写的邮箱格式不正确!</font>';}"/></td>
                                <td class="shuom">关系到发送邮件的默认账户和GRA头像的显示<span id="checkemail" class="Demo"></span></td>
                            </tr>
							<tr>
                              	<td width="180"><div align="right"> 关键字设置 </div></td>
                              	<td align="left"><input name="blog_KeyWords" type="text" class="text" value="<%=blog_KeyWords%>" maxlength="123" size="50" /></td>
                                <td class="shuom">关系到<strong>SEO</strong>对网页首页的关键字的读取</td>
                            </tr>
                            <tr>
                              	<td width="180"><div align="right"> 页面描述设置 </div></td>
                              	<td align="left"><input name="blog_Description" type="text" class="text" value="<%=blog_Description%>" maxlength="176" size="50" /></td>
                                <td class="shuom">关系到<strong>SEO</strong>对网页首页的描述的读取</td>
                            </tr>
                            <tr>
                              	<td width="180"><div align="right"> 网站备案信息 </div></td>
                              	<td align="left"><input name="blog_about" type="text" size="50" class="text" value="<%=blogabout%>"/></td>
                                <td class="shuom">请将你的<strong>备案信息</strong>填写,如还未备案,请留空.</td>
                            </tr>  
                            <tr>
                              	<td><div align="right">Blog对外开放</div></td>
                              	<td align="left"><input name="SiteOpen" type="checkbox" value="1" <%if Application(Sys.CookieName & "_SiteEnable") = 1 then response.write ("checked=""checked""")%>/></td>
                                <td class="shuom">如果关闭,用户将无法浏览您的网站.</td>
                            </tr>
                        </table>
                    </div>
                </fieldset>
                <fieldset>
		    		<legend> 日志保存设置</legend>
                    <div align="left">
                        <table border="0" cellpadding="2" cellspacing="1">
                            <tr>
                                <td width="180" align="right" valign="top" style="padding-top:8px">日志后缀列表</td>
                                <td><input type="text" name="blog_html" value="<%=blog_html%>" class="text" size="50" /></td>
                                <td class="shuom">自定义静态日志后缀列表<sup style="font-size:10px; color:#ff0000"><i>new</i></sup></td>
                            </tr>
                        </table>
                    </div>
                </fieldset>
		    <div align="left">
                <div class="SubButton">
		      		<input type="submit" name="Submit" value="保存配置" class="button"/>
		     	</div>
            </form>
<%
	End If
%>
		</td>
	</tr>
</table>	 
<%
end Sub
%>