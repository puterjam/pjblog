        <%ST(A)%>
			<div id="Content_ContentList" class="content-width"><a name="body" accesskey="B" href="#body"></a>
				<div class="pageContent">
					<img src="<$Cate_icon$>" style="margin:0px 2px -4px 0px" alt=""/> <strong><a href="default.asp?cateID=<$log_CateID$>" title="查看所有【<$Cate_Title$>】的日志"><$Cate_Title$></a></strong> <a href="feed.asp?cateID=<$log_CateID$>" target="_blank" title="订阅所有【<$Cate_Title$>】的日志" accesskey="O"><img border="0" src="images/rss.png" alt="订阅所有【<$Cate_Title$>】的日志" style="margin-bottom:-1px"/></a>
				</div>
				<div class="Content">
					<div class="Content-top"><div class="ContentLeft"></div><div class="ContentRight"></div>
					<h1 class="ContentTitle"><strong><$log_Title$></strong><$log_hiddenIcon$></h1>
					<h2 class="ContentAuthor">作者:<$log_Author$> 日期:<$log_PostTime$></h2>
				</div>
			    <div class="Content-Info">
					<div class="InfoOther">字体大小: <a href="javascript:SetFont('12px')" accesskey="1">小</a> <a href="javascript:SetFont('14px')" accesskey="2">中</a> <a href="javascript:SetFont('16px')" accesskey="3">大</a></div>
					<div class="InfoAuthor"><img src="images/weather/hn2_<$log_weather$>.gif" style="margin:0px 2px -6px 0px" alt=""/><img src="images/weather/hn2_t_<$log_weather$>.gif" alt=""/> <img src="images/<$log_level$>.gif" style="margin:0px 2px -1px 0px" alt=""/><$EditAndDel$></div>
				</div>
				<div id="logPanel" class="Content-body">
					<$ArticleContent$>
					<br/><br/>
				</div>
				<div class="Content-body">
					<$log_Modify$>
					<$log_Navigation$>
					<img src="images/From.gif" style="margin:0px 2px -4px 0px" alt=""/><strong>文章来自:</strong> <a href="<$log_FromUrl$>" target="_blank"><$log_From$></a><br/>
					<img src="images/icon_trackback.gif" style="margin:4px 2px -4px 0px" alt=""/><strong>引用通告:</strong> <a href="trackback.asp?tbID=<$LogID$>&amp;action=view" target="_blank">查看所有引用</a> | <a href="javascript:;" title="获得引用文章的链接" onclick="getTrackbackURL(<$LogID$>)">我要引用此文章</a><br/>
					<img src="images/tag.gif" style="margin:4px 2px -4px 0px" alt=""/><strong>Tags:</strong> <$log_tag$><br/>
					<img src="images/notify.gif" style="margin:4px 2px -4px 0px" alt=""/><strong>相关日志:</strong>
                    <div class="Content-body" id="related_tag" style="margin-left:25px"></div>
                    <script language="javascript" type="text/javascript">Related(<$LogID$>, 1, 1, related_tag)</script>
				</div>
				<div class="Content-bottom"><div class="ContentBLeft"></div><div class="ContentBRight"></div>评论: <$log_CommNums$> | <a href="trackback.asp?tbID=<$LogID$>&amp;action=view" target="_blank">引用: <$log_QuoteNums$></a> | 查看次数: <$log_ViewNums$></div>
			</div>
		</div>
