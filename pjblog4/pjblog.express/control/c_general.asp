<%
'=================================
' 站点基本设置
'=================================
Public Sub c_ceneral
%>
<script language="javascript">
$(function(){
	$("ul.slide").tabs("ul.toolbar", {
		current : "actived"
	});
/*
	 绑定基本设置
*/
	conMain.Base.BaseSetting();
/*
	 绑定日志评论
*/
	conMain.Base.ArtComm();
/*
	 绑定SEO优化
*/
	conMain.Base.seo.Post();
/*
	 绑定网站模式
*/
	conMain.Base.WebMode();
/*
	 绑定邮件上网
*/
	conMain.Base.WapMail();
/*
	 绑定注册模式
*/
	conMain.Base.RegFilter();
});
</script>
<style type="text/css">
.content h5{
	color:#1d6bb2;
	font-size:14px;
	margin:0;
}
h5.general{
	background:url(../../images/Control/Icon/zhengzhan.gif) no-repeat;
	line-height:28px;
	height:28px;
	padding-left:35px;
}

.toolbar li{
	padding:0px 20px 10px 20px;
	width:650px;
}

.select{
	width:750px;
}

.select .Zcontent{
	margin-top:12px;
}

.tagname{color:#424242}

.setting table tr td{
	line-height:30px;
}

.label{
	margin-bottom:10px;
}
</style>
<h5 class="general">站点基本设置</h5>
	<div class="select">
        <div class="tits">
            <ul class="slide">
                <li>基本设置</li>
                <li>网站模式</li>
                <li>WAP邮件</li>
                <li>日志评论</li>
                <li>开放过滤</li>
                <li>SEO设置</li>
                <li>其他设置</li>
            </ul>
        </div>
        
        <div class="Zcontent">
            <ul class="toolbar">
                <li><%BaseSetting%></li>
                <li><%WebMode%></li>
                <li><%WapMail%></li>
                <li><%ArtComm%></li>
                <li><%RegFilter%></li>
                <li><%SEO%></li>
                <li>Last One</li>
            </ul>
        </div>
   </div>
<%
end Sub

Public Sub BaseSetting
	Response.Write(yellowBox("整站优化及基本设置注意事项:<br />1. 自动发送邮件的设置,请参看<a href=""http://www.evio.name"" target=""_blank"">http://www.evio.name</a>上的说明.<br />2. 网站伪静态,请测试你空间是否支持IISReWrite组件,配置信息请放置到网站根目录的httpd.ini文件中.<br />3. 默认情况下, 网站未开启评论审核功能,请去 ＂<strong>日志评论</strong>＂ 模块开启评论审核功能.<br />4. 注册模式如果选择 ＂<strong>验证注册</strong>＂ 请确定是否网站服务器支持JMail组件,否则无法使用该功能.<br />5. 更多设置,请在 ＂<strong>其他设置</strong>＂ 中设置."))
%>
	<div class="basediv">
    	<div style="margin:12px 0"><input type="button" class="button" value="帮助" onclick="$('.yellowbox').slideToggle('fast')"></div>
        <div class="setting">
        	<form action="../pjblog.logic/control/log_general.asp?action=BaseSetting" method="post" id="BaseSetting">
        	<table cellpadding="3" cellspacing="0" width="100%">
            	<tr>
                	<td class="tagname">博客名称</td>
                    <td><input name="SiteName" type="text" size="30" class="text" value="<%=SiteName%>"/></td>
                    <td>将显示在网页的标题处</td>
                </tr>
                <tr>
                	<td class="tagname">博客副标题</td>
                    <td><input name="blog_Title" type="text" size="50" class="text" value="<%=blog_Title%>"/></td>
                    <td>将显示在页面的LOGO处</td>
                </tr>
                <tr>
                	<td class="tagname">站长昵称</td>
                    <td><input name="blog_master" type="text" size="10" class="text" value="<%=blog_master%>" maxlength="10"/></td>
                    <td></td>
                </tr>
                <tr>
                	<td class="tagname">博客地址</td>
                    <td><input name="SiteURL" type="text" size="50" class="text" value="<%=SiteURL%>"/></td>
                    <td>关系到<strong>RSS</strong>地址的可读性和静态页面<strong>CSS</strong>的正确加载<span id="checkweburl" class="Demo"></span></td>
                </tr>
                <tr>
                	<td class="tagname">网站备案信息</td>
                    <td><input name="blog_about" type="text" size="50" class="text" value="<%=blogabout%>"/></td>
                    <td>请将你的<strong>备案信息</strong>填写,如还未备案,请留空.</td>
                </tr>
            </table>
            <p><input type="submit" value="确定保存" class="button" /></p>
            </form>
        </div>
    </div>
<%
End Sub

Public Sub WebMode
%>
	<div class="basediv">
    	<form action="../pjblog.logic/control/log_general.asp?action=WebMode" method="post" id="WebMode">
    	<div style="margin:12px 0; line-height:30px; border-bottom:1px solid #7BBCF6"><strong>根据需要选择你喜欢的模式:</strong></div>
        <div>
        	<p><label for="p1" ><input id="p1" name="blog_postFile" type="radio" value="0" <%If blog_postFile = 0 Then%>checked="checked"<%End If%>/> 全动态模式 (不推荐)<br/> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:#999">文章数据从数据库里直接获取</span> <br/></label></p>
            <p><label for="p2" ><input id="p2" name="blog_postFile" type="radio" value="1" <%If blog_postFile = 1 Then%>checked="checked"<%End If%>/> 伪静态模式 (适合喜欢个性化的用户)<br/> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:#999">采用自定义404页面伪静态,不需要IIS的REWRITE插件支持! 需要注意的是 : 您的空间必须支持自定义404页面.</span> <br/></label></p>
            <p><label for="p3" ><input id="p3" name="blog_postFile" type="radio" value="2" <%If blog_postFile = 2 Then%>checked="checked"<%End If%>/> 全静态模式<span style="color:#f00;font-size:8px">new</span> (适合在乎seo和速度的用户)<br/> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:#999">把文章保存成全静态文件. 需要注意的是，这种模式下文章发布和编辑后需要重新生成文件. </span> <br/></label></p>
        </div>
        <div><input type="submit" value="确定" class="button" /></div>
        </form>
        
        <div style="margin:12px 0; line-height:30px; border-bottom:1px solid #7BBCF6"><strong>数据初始化:</strong></div>
        <div class="setting">
        	<table cellpadding="3" cellspacing="0" width="530">
                <tr id="t1">
                	<td class="tagname" width="150" align="right">重新统计网站数据</td>
                    <td align="center">
                    	<div class="processbar">
                        	<div class="process s1" id="n1"></div>
                        </div>
                    </td>
                    <td width="50">0%</td>
                    <td width="50"><input type="checkbox" id="s1" /></td>
                </tr>
                <tr id="t2">
                	<td class="tagname" width="150" align="right">清除访客记录</td>
                    <td align="center">
                    	<div class="processbar">
                        	<div class="process s2" id="n2"></div>
                        </div>
                   	</td>
                    <td width="50">0%</td>
                    <td width="50"><input type="checkbox" id="s2" /></td>
                </tr>
                <tr id="t3">
                	<td class="tagname" width="150" align="right">重建日志数据缓存</td>
                    <td align="center">
                    	<div class="processbar">
                        	<div class="process s3" id="n3"></div>
                        </div>
                   	</td>
                    <td width="50">14%</td>
                    <td width="50"><input type="checkbox" id="s3" /></td>
                </tr>
                <tr id="t4">
                	<td class="tagname" width="150" align="right">清理服务器缓存</td>
                    <td align="center">
                    	<div class="processbar">
                        	<div class="process s4" id="n4"></div>
                        </div>
                    </td>
                    <td width="50">0%</td>
                    <td width="50"><input type="checkbox" id="s4" /></td>
                </tr>
            </table>
            <p><input type="button" value="开始初始化" class="button" onclick="conMain.initialize.Get()" /></p>
        </div>
    </div>
<%
End Sub

Public Sub SEO
%>
	<div class="basediv">
    	<div style="margin:12px 0; line-height:30px; border-bottom:1px solid #7BBCF6"><strong>网站首页SEO设置:</strong></div>
        <div class="setting">
        	<form action="../pjblog.logic/control/log_general.asp?action=SEO" method="post" id="SEO">
        	<table cellpadding="3" cellspacing="0" width="100%">
                <tr>
                	<td class="tagname" width="85">关键字设置</td>
                    <td><input name="blog_KeyWords" type="text" class="text" value="<%=blog_KeyWords%>" maxlength="123" size="50" /></td>
                    <td>关系到<strong>SEO</strong>对网页首页的关键字的读取</td>
                </tr>
                <tr>
                	<td class="tagname" width="85">页面描述设置</td>
                    <td><textarea class="text" style="width:280px; height:50px;" name="blog_Description"><%=blog_Description%></textarea></td>
                    <td>关系到<strong>SEO</strong>对网页首页的描述的读取</td>
                </tr>
                <tr>
                	<td class="tagname" width="85">开启Ping服务</td>
                    <td><input type="checkbox" name="blog_IsPing" value="1" onclick="if (this.checked){$('#seo_ping').css('display', 'block')}else{$('#seo_ping').css('display', 'none')}" <%If blog_IsPing Then%>checked<%End If%>></td>
                    <td>搜索引擎会在最短时间内对相应博客进行抓取。</td>
                </tr>
                <tr>
                	<td class="tagname" width="85"></td>
                    <td nowrap="nowrap" colspan="2">
                    	<table id="seo_ping" style="display:<%If blog_IsPing Then%>block<%Else%>none<%End If%>;" cellpadding="3" cellspacing="0">
                    <%
					Dim conPing
					Set conPing = Server.CreateObject("Adodb.RecordSet")
						conPing.open "Select * From blog_Ping", Conn, 1, 1
						Do While Not conPing.Eof
					%>
                            	<tr id="ping_<%=Trim(conPing("Ping_ID"))%>">
                                	<td width="330">发送到: <span><%=Trim(conPing("Ping_Name"))%></span><input type="hidden" value="<%=Trim(conPing("Ping_url"))%>" id="ping_url_<%=Trim(conPing("Ping_ID"))%>" /></td>
                                	<td width="150"><a href="javascript:;" onclick="conMain.Base.seo.editPing(<%=Trim(conPing("Ping_ID"))%>, $('#ping_url_<%=Trim(conPing("Ping_ID"))%>').val())">编辑</a> | <a href="">删除</a> | <a href="javascript:;" onclick="conMain.Base.seo.see(this, $('#ping_url_<%=Trim(conPing("Ping_ID"))%>').val())">查看</a></td>
                                </tr>      
                     <%
						conPing.MoveNext
						Loop
						conPing.Close
					Set conPing = Nothing
					%>
                                <tr>
                                	<td class="add" onclick="conMain.Base.seo.ping(this)" colspan="2">添加新ping地址</td>
                                </tr>
                    	</table>
                    </td>
                </tr>
            </table>
            <p><input type="submit" value="确定保存" class="button" /></p>
            </form>
        </div>
    </div>
<%
End Sub

Public Sub WapMail
%>
	<div class="basediv">
    	<div style="margin:12px 0; line-height:30px; border-bottom:1px solid #7BBCF6"><strong>自动发信设置:</strong></div>
        <form action="../pjblog.logic/control/log_general.asp?action=WapMail" method="post" id="WapMail">
        <div class="setting">
        	<table cellpadding="3" cellspacing="0" width="100%">
                <tr>
                	<td class="tagname">SMTP服务端地址</td>
                    <td><input name="blog_smtp" type="text" class="text" value="<%=blog_smtp%>" maxlength="123" size="50" /></td>
                    <td>比如 smtp.pjblog4.com</td>
                </tr>
                <tr>
                	<td class="tagname">发信邮箱账号</td>
                    <td><input name="blog_email" type="text" class="text" value="<%=blog_email%>" maxlength="176" size="50" /></td>
                    <td>用户发信的账号,比如 admin@pjblog4.com</td>
                </tr>
                <tr>
                	<td class="tagname">发信邮箱密码</td>
                    <td><input name="blog_emailpass" type="password" class="text" value="<%=blog_emailpass%>" maxlength="176" size="50" /></td>
                    <td>用户发信的密码,必须正确可用.</td>
                </tr>
                <tr>
                	<td class="tagname">发信邮箱名称</td>
                    <td><input name="blog_sendmailName" type="text" class="text" value="<%=blog_sendmailName%>" maxlength="176" size="50" /></td>
                    <td>比如 master@pjblog4.com</td>
                </tr>
                <tr>
                	<td colspan="3"><input type="button" value="发信测试" onclick="" class="button" /></td>
                </tr>
            </table>
        </div>
        <div style="margin:12px 0; line-height:30px; border-bottom:1px solid #7BBCF6"><strong>手机上网WAP服务设置:</strong></div>
        <div class="setting">
        	<table cellpadding="3" cellspacing="0" width="100%">
                <tr>
                	<td class="tagname" width="210">允许使用wap方式浏览博客</td>
                    <td><input name="blog_wap" type="checkbox" value="1" <%if blog_wap then response.write ("checked=""checked""")%> /></td>
                    <td></td>
                </tr>
                <tr>
                	<td class="tagname">允许wap使用简单HTML</td>
                    <td><input name="blog_wapHTML" type="checkbox" value="1" <%if blog_wapHTML then response.write ("checked=""checked""")%>  /></td>
                    <td></td>
                </tr>
                <tr>
                	<td class="tagname">允许wap显示图片</td>
                    <td><input name="blog_wapImg" type="checkbox" value="1" <%if blog_wapImg then response.write ("checked=""checked""")%>   /></td>
                    <td></td>
                </tr>
                <tr>
                	<td class="tagname">允许wap保留文章超链接</td>
                    <td><input name="blog_wapURL" type="checkbox" value="1" <%if blog_wapURL then response.write ("checked=""checked""")%>   /></td>
                    <td></td>
                </tr>
                <tr>
                	<td class="tagname">允许通过wap登录</td>
                    <td><input name="blog_wapLogin" type="checkbox" value="1" <%if blog_wapLogin then response.write ("checked=""checked""")%>  /></td>
                    <td></td>
                </tr>
                <tr>
                	<td class="tagname">允许通过wap发评论</td>
                    <td><input name="blog_wapComment" type="checkbox" value="1" <%if blog_wapComment then response.write ("checked=""checked""")%>   /></td>
                    <td></td>
                </tr>
                <tr>
                	<td class="tagname">Wap日志显示数量</td>
                    <td><input name="blog_wapNum" type="text" size="5" class="text" value="<%=blog_wapNum%>"/> 篇</td>
                    <td></td>
                </tr>
            </table>
        </div>
        <p><input type="submit" value="确定保存" class="button" /></p>
        </form>
    </div>
<%
End Sub

Public Sub ArtComm
%>
	<div class="basediv">
    	<div style="margin:12px 0; line-height:30px; border-bottom:1px solid #7BBCF6"><strong>日志基本设置:</strong></div>
        <form action="../pjblog.logic/control/log_general.asp?action=ArtComm" method="post" id="ArtComm">
        <div class="setting">
        	<table cellpadding="3" cellspacing="0" width="100%">
                <tr>
                	<td class="tagname">日志后缀列表</td>
                    <td><input type="text" name="blog_html" value="<%=blog_html%>" class="text" size="50" /></td>
                    <td>自定义静态日志后缀列表<sup style="font-size:10px; color:#ff0000"><i>new</i></sup></td>
                </tr>
                <tr>
                	<td class="tagname">日志预览分割类型</td>
                    <td>
                    	<select name="blog_SplitType">
                            <option value="0" <%if Not blog_SplitType then Response.Write ("selected=""selected""")%>>按照字符数量分割</option>
                            <option value="1" <%if blog_SplitType then Response.Write ("selected=""selected""")%>>按照行数分割</option>
                        </select>
                    </td>
                    <td>只对重新编辑日志或新建日志有效</td>
                </tr>
                <tr>
                	<td class="tagname">日志预览最大字符数</td>
                    <td><input name="blog_introChar" type="text" size="5" class="text" value="<%=blog_introChar%>"/> 个</td>
                    <td>只对UBB编辑器有效</td>
                </tr>
                <tr>
                	<td class="tagname">日志预览切割行数</td>
                    <td><input name="blog_introLine" type="text" size="5" class="text" value="<%=blog_introLine%>"/> 行</td>
                    <td>根据编辑器中的换行来进行分割</td>
                </tr>
                <tr>
                	<td class="tagname">每页显示日志</td>
                    <td><input name="blogPerPage" type="text" size="5" class="text" value="<%=blogPerPage%>"/> 篇</td>
                    <td></td>
                </tr>
            </table>
        </div>
        <div style="margin:12px 0; line-height:30px; border-bottom:1px solid #7BBCF6"><strong>评论设置:</strong></div>
        <div class="setting">
        	<table cellpadding="3" cellspacing="0" width="100%">
                <tr>
                	<td class="tagname" width="210">每页显示评论数</td>
                    <td><input name="blogcommpage" type="text" size="5" class="text" value="<%=blogcommpage%>"/> 篇</td>
                    <td></td>
                </tr>
                <tr>
                	<td class="tagname">发表评论时间间隔</td>
                    <td><input name="blog_commTimerout" type="text" size="5" class="text" value="<%=blog_commTimerout%>"/> 秒</td>
                    <td></td>
                </tr>
                <tr>
                	<td class="tagname">发表评论字数限制</td>
                    <td><input name="blog_commLength" type="text" size="5" class="text" value="<%=blog_commLength%>"/> 字</td>
                    <td></td>
                </tr>
                <tr>
                	<td class="tagname">发表评论必须输入验证码</td>
                    <td><input name="blog_validate" type="checkbox" value="1" <%if blog_validate then response.write ("checked=""checked""")%>  /></td>
                    <td>可以让会员不用输入验证码，只有全动态模式有效</td>
                </tr>
                <tr>
                	<td class="tagname">禁用评论UBB代码</td>
                    <td><input name="blog_commUBB" type="checkbox" value="1" <%if blog_commUBB then response.write ("checked=""checked""")%>  /></td>
                    <td></td>
                </tr>
                <tr>
                	<td class="tagname">禁用评论贴图</td>
                    <td><input name="blog_commIMG" type="checkbox" value="1" <%if blog_commIMG then response.write ("checked=""checked""")%> /></td>
                    <td></td>
                </tr>
                <tr>
                	<td class="tagname">评论审核</td>
                    <td><input name="blog_commAduit" type="checkbox" value="1" <%if blog_commAduit then response.write ("checked=""checked""")%> /></td>
                    <td>会员或者游客发表的评论都需要通过管理员审核</td>
                </tr>
            </table>
        </div>
        <p><input type="submit" value="确定保存" class="button" /></p>
        </form>
    </div>
<%
End Sub

Public Sub RegFilter
%>
	<div class="basediv">
    	<form action="../pjblog.logic/control/log_general.asp?action=RegFilter" method="post" id="RegFilter">
    	<div style="margin:12px 0; line-height:30px; border-bottom:1px solid #7BBCF6"><strong>注册模式选择(请谨慎选择):</strong></div>
        <div>
        	<p><label for="p4" ><input id="p4" name="blog_Disregister" type="radio" value="0" <%if blog_Disregister = 0 then response.write ("checked=""checked""")%>/> 不允许注册 (不推荐)<br/> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:#999"> 选择后,用户将无法在您站上注册!</span> <br/></label></p>
            <p><label for="p5" ><input id="p5" name="blog_Disregister" type="radio" value="1" <%if blog_Disregister = 1 then response.write ("checked=""checked""")%>/> 普通注册 (推荐)<br/> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:#999">只需要填写用户名和密码来注册!和PJBlog3注册方式相同.</span> <br/></label></p>
            <p><label for="p6" ><input id="p6" name="blog_Disregister" type="radio" value="2" <%if blog_Disregister = 2 then response.write ("checked=""checked""")%>/> 验证注册<span style="color:#f00;font-size:8px">new</span> (需要服务端JMail组件的支持)<br/> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:#999">需要告诉用户,注册后请返回注册邮箱激活注册,否则,该用户处于待激活状态,无法登入本站. </span> <br/></label></p>
        </div>
        <div><input type="submit" value="确定" class="button" /></div>
        </form>
    </div>
<%
End Sub
%>