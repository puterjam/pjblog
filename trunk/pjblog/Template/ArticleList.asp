<%ST(A)%><$log_Author$><%ST(A)%><$log_viewCount$><%ST(A)%>
		<div class="Content">
			<div class="Content-top"><div class="ContentLeft"></div><div class="ContentRight"></div>
				 <$ShowButton$>
				 <h1 class="ContentTitle"><img src="<$Cate_icon$>" style="margin:0px 2px -4px 0px;" alt="" class="CateIcon"/><a class="titleA" href="<$pLink$>"><$log_Title$></a><$log_hiddenIcon$></h1>
				 <h2 class="ContentAuthor">作者:<$log_Author$> 日期:<$log_PostTime$></h2>
			</div>
			<div id="log_<$LogID$>"<$ShowStyle$>>
				<div class="Content-body">
					 <$log_Intro$>
					 <$log_readMore$>
					 <$log_tag$>
				</div>
				<div class="Content-bottom"><div class="ContentBLeft"></div><div class="ContentBRight"></div>
						分类:<a href="default.asp?cateID=<$log_CateID$>" title=""><$Cate_Title$></a> | <a href="?id=<$LogID$>">固定链接</a> | <a href="<$pLink$>#comm_top">评论: <$log_CommNums$></a> | <a href="trackback.asp?tbID=<$LogID$>&amp;action=view" target="_blank">引用: <$log_QuoteNums$></a> | 查看次数: <$log_viewC$>
							 <$editRight$>
			    </div>
			</div>
		</div>
	<%ST(A)%>
		<tr>
			<td valign="top">
			    <a href="default.asp?cateID=<$log_CateID$>" ><img border="0" alt="查看 <$Cate_Title$> 的日志" src="<$Cate_icon$>" style="margin:0px 2px -3px 0px"/></a>
				<a href="<$pLink$>" title="作者:<$log_Author$> 日期:<$log_PostTime$>"><$log_Title$></a><$log_hiddenIcon$>
			</td>
			<td valign="top" width="60"><nobr><a href="<$pLink$>#comm_top" title="评论"><$log_CommNums$></a> | <a href="trackback.asp?tbID=<$LogID$>&amp;action=view" target="_blank" title="引用通告"><$log_QuoteNums$></a> | <span title="查看次数"><$log_viewC$></span></nobr></td>
		</tr>
   <%ST(A)%>
		<div class="Content">
			<div class="Content-top"><div class="ContentLeft"></div><div class="ContentRight"></div>
				 <$ShowButton$>
				 <h1 class="ContentTitle"><img src="<$Cate_icon$>" style="margin:0px 2px -4px 0px;" alt="" class="CateIcon"/><a class="titleA" href="<$pLink$>"><$Show_Title$></a><$log_hiddenIcon$></h1>
				 <h2 class="ContentAuthor">作者:<$log_Author$> 日期:<$log_PostTime$></h2>
			</div>
			<div id="log_<$LogID$>"<$ShowStyle$>>
				<div class="Content-body">
					 <$log_Secret$>
				</div>
				<div class="Content-bottom"><div class="ContentBLeft"></div><div class="ContentBRight"></div>
						分类:<a href="default.asp?cateID=<$log_CateID$>" title=""><$Cate_Title$></a> | <a href="?id=<$LogID$>">固定链接</a> | <a href="<$pLink$>#comm_top">评论: <$log_CommNums$></a> | <a href="trackback.asp?tbID=<$LogID$>&amp;action=view" target="_blank">引用: <$log_QuoteNums$></a> | 查看次数: <$log_viewC$>
							 <$editRight$>
			    </div>
			</div>
		</div>
   <%ST(A)%>
		<tr>
			<td valign="top">
			    <a href="default.asp?cateID=<$log_CateID$>" ><img border="0" alt="查看 <$Cate_Title$> 的日志" src="<$Cate_icon$>" style="margin:0px 2px -3px 0px"/></a>
				<a href="<$pLink$>" title="作者:<$log_Author$> 日期:<$log_PostTime$>"><$Show_Title$></a><$log_hiddenIcon$>
			</td>
			<td valign="top" width="60"><nobr><a href="<$pLink$>#comm_top" title="评论"><$log_CommNums$></a> | <span title="引用通告"><$log_QuoteNums$></span> | <span title="查看次数"><$log_viewC$></span></nobr></td>
		</tr>