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
		          				<td width="180"><div align="right"> BLOG 地址
		                  		<div class="shuom">关系到<strong>RSS</strong>地址的可读性</div>
		          				</div></td>
		          				<td align="left"><input name="SiteURL" type="text" size="50" class="text" value="<%=SiteURL%>"/></td>
		        			</tr>
                            <tr>
                                <td width="180"><div align="right"> 站长邮件地址 </div></td>
                                <td align="left"><input name="blog_email" type="text" size="50" class="text" value="<%=blog_email%>"/></td>
                            </tr>
							<tr>
                              	<td width="180"><div align="right"> 首页 KeyWords 设置 </div></td>
                              	<td align="left"><input name="blog_KeyWords" type="text" class="text" value="<%=blog_KeyWords%>" size="100" maxlength="123"/></td>
                            </tr>
                            <tr>
                              	<td width="180"><div align="right"> 首页 Description 设置 </div></td>
                              	<td align="left"><input name="blog_Description" type="text" class="text" value="<%=blog_Description%>" size="100" maxlength="176"/></td>
                            </tr>
                            <tr>
                              	<td width="180"><div align="right"> 网站备案信息 </div></td>
                              	<td align="left"><input name="blog_about" type="text" size="50" class="text" value="<%=blogabout%>"/></td>
                            </tr>  
                            <tr>
                              	<td><div align="right">Blog对外开放</div></td>
                              	<td align="left"><input name="SiteOpen" type="checkbox" value="1" <%if Application(Sys.CookieName & "_SiteEnable") = 1 then response.write ("checked=""checked""")%>/></td>
                            </tr>
                        </table>
                    </div>
                </fieldset>
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