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
                <a href="?Fmenu=General&Smenu=Misc">初始化数据</a>
            </div>
<%
	If Request.QueryString("Smenu") = "visitors" Then
	ElseIf Request.QueryString("Smenu") = "Misc" Then
%>
			<table cellpadding="3" cellspacing="0" width="100%" id="Static">
            	<tr><td colspan="6"><strong>1. 清理服务器缓存</strong></td></tr>
                <tr id="clearTr" style="display:none">
                	<td width="10">&nbsp;</td>
                	<td colspan="4"><div id="clear"></div></td>
                    <td width="10">&nbsp;</td>
                </tr>
                <tr>
                	<td width="10">&nbsp;</td>
                	<td colspan="4"><input type="button" value="开始清理缓存" onclick="CheckForm.Clear(this)" class="button" /></td>
                    <td width="10">&nbsp;</td>
                </tr>
                
                <tr><td colspan="6"><strong>2. 静态化设置</strong></td></tr>
                
                <tr id="StaticIndex" class="ceo">
                	<td width="10">&nbsp;</td>
                	<td width="180"><div style="line-height:18px;"><label for="S1"><input type="checkbox" onclick="if (this.checked){CheckForm.Static.index($('Static'), 'StaticIndex', 1, this)}else{try{$('StaticPre').parentNode.removeChild($('StaticPre'));}catch(e){}}" id="S1" style=" margin-right:10px;" />首页静态化</label></div></td>
                    <td colspan="4">&nbsp;</td>
                </tr>
                
                <tr id="StaticArticle" class="ceo">
                	<td width="10">&nbsp;</td>
                	<td width="180"><div style="line-height:18px;"><label for="S2"><input type="checkbox" onclick="if (this.checked){CheckForm.Static.Article($('Static'), 'StaticArticle', 1, this)}else{try{$('StaticPre').parentNode.removeChild($('StaticPre'));}catch(e){}}" id="S2" style=" margin-right:10px;" />内容页静态化</label></div></td>
                    <td colspan="4">&nbsp;</td>
                </tr>
                
                <tr id="StaticCategory" class="ceo">
                	<td width="10">&nbsp;</td>
                	<td width="180"><div style="line-height:18px;"><label for="S3"><input type="checkbox" onclick="if (this.checked){CheckForm.Static.Category($('Static'), 'StaticCategory', 1, this)}else{try{$('StaticPre').parentNode.removeChild($('StaticPre'));}catch(e){}}" id="S3" style=" margin-right:10px;" />分类列表静态化</label></div></td>
                    <td colspan="4">&nbsp;</td>
                </tr>
                
                <tr><td colspan="6"><strong>3. 初始化整站信息.</strong></td></tr>
                
            </table>
<%
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
                            <tr class="ceo">
                            	<td width="180"><div align="right"> BLOG 名称 </div></td>
                              	<td align="left"><input name="SiteName" type="text" size="30" class="text" value="<%=SiteName%>"/></td>
                                <td class="shuom">将显示在网页的标题处</td>
                            </tr>
                            <tr class="ceo">
		          				<td width="180"><div align="right"> BLOG 副标题 </div></td>
		          				<td align="left"><input name="blog_Title" type="text" size="50" class="text" value="<%=blog_Title%>"/></td>
                                <td class="shuom">将显示在页面的LOGO处</td>
		        			</tr>
							<tr class="ceo">
		          				<td width="180"><div align="right"> 站长昵称 </div></td>
		          				<td align="left"><input name="blog_master" type="text" size="10" class="text" value="<%=blog_master%>" maxlength="10"/></td>
                                <td class="shuom"></td>
		        			</tr>
                            <tr class="ceo">
		          				<td width="180"><div align="right"> BLOG 地址</div></td>
		          				<td align="left"><input name="SiteURL" type="text" size="50" class="text" value="<%=SiteURL%>" onblur="if (this.value.reg(REGEXP.REG_WEBURL)){$('checkweburl').innerHTML = '<font color=green>博客地址格式正确</font>';}else{$('checkweburl').innerHTML = '<font color=red>您所填写的博客地址格式不正确!</font>';}"/></td>
                                <td class="shuom">关系到<strong>RSS</strong>地址的可读性和静态页面<strong>CSS</strong>的正确加载<span id="checkweburl" class="Demo"></span></td>
		        			</tr>
                            <tr class="ceo">
                                <td width="180"><div align="right"> 站长邮件地址 </div></td>
                                <td align="left"><input name="blog_email" type="text" size="50" class="text" value="<%=blog_email%>" onblur="if (this.value.reg(REGEXP.REG_EMAIL)){$('checkemail').innerHTML = '<font color=green>邮箱格式正确</font>';}else{$('checkemail').innerHTML = '<font color=red>您所填写的邮箱格式不正确!</font>';}"/></td>
                                <td class="shuom">关系到发送邮件的默认账户和GRA头像的显示<span id="checkemail" class="Demo"></span></td>
                            </tr>
							<tr class="ceo">
                              	<td width="180"><div align="right"> 关键字设置 </div></td>
                              	<td align="left"><input name="blog_KeyWords" type="text" class="text" value="<%=blog_KeyWords%>" maxlength="123" size="50" /></td>
                                <td class="shuom">关系到<strong>SEO</strong>对网页首页的关键字的读取</td>
                            </tr>
                            <tr class="ceo">
                              	<td width="180"><div align="right"> 页面描述设置 </div></td>
                              	<td align="left"><input name="blog_Description" type="text" class="text" value="<%=blog_Description%>" maxlength="176" size="50" /></td>
                                <td class="shuom">关系到<strong>SEO</strong>对网页首页的描述的读取</td>
                            </tr>
                            <tr class="ceo">
                              	<td width="180"><div align="right"> 网站备案信息 </div></td>
                              	<td align="left"><input name="blog_about" type="text" size="50" class="text" value="<%=blogabout%>"/></td>
                                <td class="shuom">请将你的<strong>备案信息</strong>填写,如还未备案,请留空.</td>
                            </tr>  
                            <tr class="ceo">
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
                            <tr class="ceo">
                                <td width="180" align="right" valign="top" style="padding-top:8px">日志后缀列表</td>
                                <td><input type="text" name="blog_html" value="<%=blog_html%>" class="text" size="50" /></td>
                                <td class="shuom">自定义静态日志后缀列表<sup style="font-size:10px; color:#ff0000"><i>new</i></sup></td>
                            </tr>
                            <tr class="ceo">
                                <td width="180" align="right" valign="top">日志输出模式</td>
                                <td colspan="2">
                                	<label for="p1" ><input id="p1" name="blog_postFile" type="radio" value="0" disabled="disabled"/> 全动态模式 (不推荐)<br/> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:#999">文章数据从数据库里直接获取</span> <br/></label>
                                	<label for="p2" ><input id="p2" name="blog_postFile" type="radio" value="1" disabled="disabled"/> 伪静态模式 (适合喜欢个性化的用户)<br/> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:#999">采用自定义404页面伪静态,不需要IIS的REWRITE插件支持! 需要注意的是 : 您的空间必须支持自定义404页面.</span> <br/></label>
                               		<label for="p3" ><input id="p3" name="blog_postFile" type="radio" value="2" checked="checked"/> 全静态模式<span style="color:#f00;font-size:8px">new</span> (适合在乎seo和速度的用户)<br/> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:#999">把文章保存成全静态文件. 需要注意的是，这种模式下文章发布和编辑后需要重新生成文件. </span> <br/></label>
                                    <div style="border-top:1px solid #ccc;padding-top:5px;margin-top:5px"><img src="../images/notify.gif"/> <b>温馨提示:</b> 进行半静态和全静态切换后，需要到 <a href="ConContent.asp?Fmenu=General&Smenu=Misc">初始化数据</a> 里更新<b>日志列表缓存</b>。
</div><br/><%if not CheckObjInstalled("ADODB.Stream") then response.write "<b style='color:#f00'>需要 ADODB.Stream 组件支持</b>"%>
                                </td>
                            </tr>
                            <tr>
                            	<td width="180" align="right">日志预览分割类型</td>
                                <td>
                                	<select name="blog_SplitType">
                                    	<option value="0" <%if Not blog_SplitType then Response.Write ("selected=""selected""")%>>按照字符数量分割</option>
                                    	<option value="1" <%if blog_SplitType then Response.Write ("selected=""selected""")%>>按照行数分割</option>
                                    </select>
                                </td>
                                <td class="shuom">只对重新编辑日志或新建日志有效</td>
                            </tr>
			  				<tr>
                            	<td width="180" align="right">日志预览最大字符数</td>
                                <td><input name="blog_introChar" type="text" size="5" class="text" value="<%=blog_introChar%>"/> 个</td>
                                <td class="shuom">只对UBB编辑器有效</td>
                            </tr>
			  				<tr>
                            	<td width="180" align="right">日志预览切割行数</td>
                                <td><input name="blog_introLine" type="text" size="5" class="text" value="<%=blog_introLine%>"/> 行</td>
                                <td class="shuom">根据编辑器中的换行来进行分割</td>
                            </tr>
                        </table>
                    </div>
                </fieldset>
                <fieldset>
		   		 	<legend> 显示设置</legend>
		    		<div align="left">
		      		<table border="0" cellpadding="2" cellspacing="1">
			  			<tr>
                        	<td width="180" align="right">每页显示日志</td>
                            <td width="300"><input name="blogPerPage" type="text" size="5" class="text" value="<%=blogPerPage%>"/> 篇</td>
                            <td class="shuom">&nbsp;</td>
                        </tr>
			  			<tr>
                        	<td width="180" align="right">是否在首页显示图片友情链接</td>
                            <td width="300"><input name="blog_ImgLink" type="checkbox" value="1" <%if blog_ImgLink then response.write ("checked=""checked""")%>  /> </td>
                            <td class="shuom">在首页面上显示图片友情链接</td>
                        </tr>
		      		</table>
		    		</div>
				</fieldset>
                <fieldset>
                    <legend> 评论设置</legend>
                        <table border="0" cellpadding="2" cellspacing="1">
                        	
                            <tr>
                            	<td width="180" align="right">每页显示评论数</td>
                                <td width="300"><input name="blogcommpage" type="text" size="5" class="text" value="<%=blogcommpage%>"/> 篇</td>
                                <td>&nbsp;</td>
                            </tr>
                            
                        	<tr>
                            	<td width="180" align="right">发表评论时间间隔</td>
                                <td width="300"><input name="blog_commTimerout" type="text" size="5" class="text" value="<%=blog_commTimerout%>"/> 秒</td>
                                <td>&nbsp;</td>
                            </tr>
                            
                        	<tr>
                            	<td width="180" align="right">发表评论字数限制</td>
                                <td width="300"><input name="blog_commLength" type="text" size="5" class="text" value="<%=blog_commLength%>"/> 字</td>
                                <td>&nbsp;</td>
                            </tr>
                            
                        	<tr>
                            	<td width="180" align="right">发表评论必须输入验证码</td>
                                <td width="300"><input name="blog_validate" type="checkbox" value="1" <%if blog_validate then response.write ("checked=""checked""")%>  /> </td>
                                <td class="shuom">可以让会员不用输入验证码，只有全动态模式有效</td>
                            </tr>
                            
                        	<tr>
                            	<td width="180" align="right">禁用评论UBB代码</td>
                                <td width="300"><input name="blog_commUBB" type="checkbox" value="1" <%if blog_commUBB then response.write ("checked=""checked""")%>  /></td>
                                <td>&nbsp;</td>
                            </tr>
                            
                        	<tr>
                            	<td width="180" align="right">禁用评论贴图</td>
                                <td width="300"><input name="blog_commIMG" type="checkbox" value="1" <%if blog_commIMG then response.write ("checked=""checked""")%> /></td>
                                <td>&nbsp;</td>
                            </tr>
                            <tr>
                            	<td width="180" align="right">评论审核</td>
                                <td width="300"><input name="blog_commAduit" type="checkbox" value="1" <%if blog_commAduit then response.write ("checked=""checked""")%> /></td>
                                <td>&nbsp;</td>
                            </tr>
                        </table>
                    </fieldset>
                    <fieldset>
                        <legend> 用户注册与过滤</legend>
                        <div align="left">
                          <table border="0" cellpadding="2" cellspacing="1">
                            <tr>
                            	<td width="180" align="right" valign="middle">注册新用户</td>
                                <td align="left" colspan="2">
                                <div>
                                    <label for="p4" ><input id="p4" name="blog_Disregister" type="radio" value="0" <%if blog_Disregister = 0 then response.write ("checked=""checked""")%>/> 不允许注册 (不推荐)<br/> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:#999">用户将无法在您站上注册!</span> <br/></label>
                                	<label for="p5" ><input id="p5" name="blog_Disregister" type="radio" value="1" <%if blog_Disregister = 1 then response.write ("checked=""checked""")%>/> 普通注册 (推荐)<br/> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:#999">只需要填写用户名和密码来注册!</span> <br/></label>
                               		<label for="p6" ><input id="p6" name="blog_Disregister" type="radio" value="2" <%if blog_Disregister = 2 then response.write ("checked=""checked""")%>/> 验证注册<span style="color:#f00;font-size:8px">new</span> (特殊支持的用户使用)<br/> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:#999">通过发送邮件来完成注册! </span> <br/></label>
                                </div>
                                <div style="border-top:1px solid #ccc;padding-top:5px;margin-top:5px"><img src="../images/notify.gif"/> <b>温馨提示:</b> 选择 <strong>验证注册</strong> 必须是您的空间支持<strong>JMail组件</strong>!
</div>
                                </td>
                            </tr> 
                            <tr><td colspan="3" height="10"></td></tr>
                            <tr>
                            	<td align="right">访客记录最大值</td>
                                <td align="left"><input name="blog_CountNum" type="text" size="5" class="text" value="<%=blog_CountNum%>"/> </td>
                                <td class="shuom">设置为0则不进行任何访客记录</td>
                            </tr>
                         <tr>
                             <td valign="top"><div align="right">注册名字过滤</div></td>
                              <td align="left"><textarea name="blog_FilterName" cols="50" rows="5"><%=Register_UserNames%></textarea></td>
                              <td class="shuom">用"|"分割需要过滤的名字</td>
                            </tr>
                            <tr>
                              <td valign="top"><div align="right">IP过滤</div></td>
                              <td align="left"><textarea name="blog_FilterIP" cols="50" rows="5"><%=FilterIPs%></textarea></td>
                              <td class="shuom">以下Ip地址将无法访问Blog<br />使用"|"分割IP地址,IP地址可以包含通配符号"*"用来禁止某个IP段的无法访问Blog
                              </td>
                            </tr>
                         </table>
                        </div>
                    </fieldset>
		    <div align="left">
                <div class="SubButton">
		      		<input type="submit" name="Submit" value="保存配置" class="button"/>
		     	</div>
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