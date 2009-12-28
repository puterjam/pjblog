<%
'=================================
' 后台欢迎页面
'=================================
Public Sub c_welcome
	%>   
<script language="javascript" type="text/javascript">
	$(function(){
	  	$("#tip_word").hover(
			function(){
				$("#tip_word .edit").css("display", "block");
			},
			function(){
				$("#tip_word .edit").css("display", "none");
			}
		);
		
    	$("ul.slide > li:first-child").addClass("actived");
		$("div.Zcontent > ul > li").hide();
		$("div.Zcontent > ul > li:first").show();
    	$("div.Zcontent ul li").attr("id", function(){return idNumber("No") + $("div.Zcontent ul li").index(this)});
     	$("ul.slide li").click(function(){
         	var c = $("ul.slide li");
         	var index = c.index(this);
         	var p = idNumber("No");
         	show(c, index, p);
    	});
     
     	function show(controlMenu, num, prefix){
        	var content = prefix + num;
         	$('#'+content).siblings().hide();
         	$('#'+content).show();
        	controlMenu.eq(num).addClass("actived").siblings().removeClass("actived");
     	};

     	function idNumber(prefix){
         	var idNum = prefix;
         	return idNum;
     	};
 });
</script>

<style type="text/css">

.welcome .left{
	float:left;
	width:480px;
}
.welcome .right{
	float:right;
	width:259px;
}
.welcome .right .side{
	margin:0 0 30px 0;
	padding:0;
}
.welcome .right .side .title{
}
.welcome .right .side .title .top{
	background: url(../../images/Control/side_top.jpg) no-repeat;
	width:259px;
	height:3px;
}
.welcome .right .side .title .mid{
	background:url(../../images/Control/side_mid.jpg) repeat-y;
	height:26px;
	line-height:26px;
}
.welcome .right .side .title .mid .tit{
	padding:0px 10px 0px 10px;
	font-weight:700;
}
.welcome .right .side .title .bom{
	background:url(../../images/Control/side_bom.jpg) no-repeat;
	width:259px;
	height:3px;
}

.contList ol{
	list-style:none;
	margin:10px;
	padding:0px;
}
.contList ol li{
	border-bottom:1px solid #eee;
	margin-bottom:10px;
}
.article{
	height:40px;
}
.article div{
	line-height:20px;
}
.article .tis{
	color:#333333;
	height:20px;
	overflow:hidden;
	background:url(../../images/Control/side_point.gif) no-repeat 0px 10px;
	padding-left:10px;
}
.article .info{
	font-size:10px;
	color:#ccc;
}
.article .info a:link, .article .info a:visited{
	color:#ccc
}
.article .info a:hover{
	color:#ff6000;
}
.article .info .time{
	float:left;
	font-style:oblique;
}
.article .info .read{
	float:right;
	font-style:oblique;
}

.side .search{
	background:url(../../images/Control/side_search_bj.jpg) no-repeat;
	width:259px;
	height:109px;
}

.user{
	padding:12px;
}

.user .photo{
}
.user .photo .le{
	float:left;
}
.user .photo .ri{
	float:right;
	text-align: left;
	width:300px;
	line-height:30px;
}
.user .photo .ri .aut{
	color:#64b0f5;
	border-bottom:1px solid #c9e3fc
}
.user .photo .ri .autinfo{
	font-style:oblique;
	font-size:11px;
}
.tip{
	position:relative;
	left:-12px;
}

.tip_top{
	background:url(../../images/Control/tip_top.jpg) no-repeat;
	width:287px;
	height:9px;
}
.tip_bom{
	background:url(../../images/Control/tip_bom.jpg) no-repeat;
	width:287px;
	height:51px;
}
.tip_bom .te{
	margin:0px 20px 0px 30px;
}

.tip_bom .te .word{
	line-height:20px;
	height:20px;
	overflow:hidden;
}

.tip_bom .te .edit{
	text-align:right;
	display:none;
}

#setUser{
	margin-top:10px;
	margin-left:25px;
	color:#4678a5;
	background:url(../../images/Control/list.gif) no-repeat right;
	padding-right:20px;
}
#setUser:hover{
	cursor:pointer;
}
.trim{
	background:url(../../images/Control/post-bg.jpg) no-repeat;
	width:485px;
	height:134px;
	padding:0;
	margin:0;
}

.trim .title{
	line-height:50px;
	height:50px;
	padding:0px 15px 0px 15px;
}
.trim .title .o{border-bottom:1px solid #cce2f7;}

.trim .Mcontent{
	padding:0px 15px 0px 15px;
	height:76px;
	overflow:hidden;
}
.trim .Mcontent .l{
	width:75px;
	height:65px;
	float:left
}

.trim .Mcontent .l img{
	margin-top:6px;
}

.trim .Mcontent .r{
	float:left;
	padding:20px 10px 0px 20px;
	line-height:25px;
	font-size:12px;
	color:#1d6bb2;
	font-weight:700;
}

.trim .Mcontent .r span{
	font-size:11px;
	font-weight: normal;
	color:#919191;
}

.trim .Mcontent .r a:link, .trim .Mcontent .r a:visited{
	color:#919191;
}

.trim .Mcontent .r a:hover{
	color:#ff6000;
}

.edit a:link, .edit a:visited{
	color:#ccc;
}
.edit a:hover{
	color:#ff6000;
}

.tipwordbutton{
	background:none;
	border:0px;
	color:#3498F1;
	font-family:Arial, Helvetica, sans-serif;
	cursor:pointer;
}
.tipwordbutton:hover{
	color:#3498F1;
}

.person_top{
	background:#fff url(../../images/Control/peisontip_top.gif) no-repeat;
	width:176px;
	height:14px;
}

.person_mid{
	background:#fff url(../../images/Control/peisontip_mid.gif) repeat-y;
	width:176px;
}

.person_mid .cContent{
	padding:4px 15px 25px 15px;
}

.person_mid .cContent ul{
	list-style:none;
	margin:0;
	padding:0;
}

.person_mid .cContent ul li{
	line-height:26px;
	border-bottom:1px solid #efefef;
	padding:0px 10px 0px 10px;
}

.person_mid .cContent ul li a:link, .person_mid .cContent ul li a:visited{
	color:#6B6B6B;
	text-decoration:none;
}

.person_mid .cContent ul li a:hover{
	color:#ff6000;
}

.person_bom{
	background:#fff url(../../images/Control/peisontip_bom.gif) no-repeat;
	width:176px;
	height:6px;
}

.ad{
	background:url(../../images/Control/advanser.jpg) no-repeat;
	width:485px;
	height:95px;
	margin-top:15px
}

.select{
	width:485px;
	overflow:hidden;
	margin-top:15px;
}

.select .tits{
	border-bottom:1px solid #7bbcf6;
	height:30px;
}

.select .tits ul.slide{
	list-style:none;
	margin:0;
	padding:0;
	padding-left:15px;
}

.select .tits ul.slide li{
	height:30px;
	line-height:30px;
	width:92px;
	margin-right:3px;
	background:url(../../images/Control/selectdiv.gif) no-repeat;
	text-align:center;
	float:left;
	cursor:pointer;
}

.select .Zcontent{
	width:485px;
	overflow:hidden;
	margin-top:15px;
}

.select .Zcontent ul{
	margin:0;
	padding:0;
	list-style:none;
}

.select .Zcontent ul li{
	width:485px;
	color:#6f6e6e;
}

.actived{
	border-bottom:1px solid #fff;
}

.normal{
	border-bottom:1px solid #7bbcf6;
}
.slideTop a:link, .slideTop a:visited{
	color:#1d6bb2;
}
.slideTop a:hover{
	color:#ff6000;
}

.slideContent{
	width:485px;
	overflow:hidden;
}

.slideContent .slideItem{
	border-bottom:1px solid #ddd;
	padding:10px;
}

.slideContent .slideItem .slideItem_left{
	width:40px;
	float:left;
}

.slideContent .slideItem .slideItem_right{
	float:right;
	width:400px;
}

.slideContent .slideItem .slideItem_right .h6{
	line-height:20px;
	font-size:11px;
	font-style:normal;
	font-weight:700
}

.slideContent .slideItem .slideItem_right .h5{
	margin:4px 0px 4px 0px;
	text-indent:2em;
	line-height:22px;
}

.slideContent .slideItem .slideItem_right .h4{
	height:20px!important;
	line-height:20px;
}
</style>
		<div class="welcome">
        	<div class="left">
                <div class="user">
                	<div class="photo">
                    	<div class="le photo-l"><img src="http://www.gravatar.com/avatar/6bab7c91a03d47fe1aa5b5b6b6f8cc55?s=80&r=r&d=http%3a%2f%2fwww.gravatar.com%2favatar%2fad516503a11cd5ca435acc9bb6523536%3fs%3d35"></div>
                        <div class="ri">
							<div class="aut"><%=memName%></div>
                            <div class="autinfo">最后登入 1986-10-31 12:00:01,IP 127.0.0.1</div>
                            <div class="tip">
                            	<div class="tip_top"></div>
                                <div class="tip_bom" id="tip_word"><div class="te"><div class="word">还有什么好说的,我草你吗...</div><div class="edit"><a href="javascript:;" onclick="conMain.tip.edit()">修改</a></div></div></div>
                            </div>
                        </div>
                        <div class="clear"></div>
                    </div>
                    <span id="setUser" onclick="conMain.PersonSet(this)">个人设置</span>
                </div>
                
                <div class="trim">
                    <div class="title">
                    	<div class="o">您好, <strong><%=memName%></strong> , 您可以点击以下图表快速发表日志.</div>
                    </div>
                    <div class="Mcontent">
                        <div class="l"><a href=""><img src="../../images/Control/trimbg.png" width="75" height="65" border="0" /></a></div>
                        <div class="r">发表自己的日志哦!<br /><span><a href="">快速发表</a></span></div>
                        <div class="clear"></div>
                    </div>
                </div>
                
               <div class="ad">
               
               </div> 
               
               <div class="select">
               		<div class="tits">
                    	<ul class="slide">
                        	<li>最新评论</li>
                            <li>最新留言</li>
                            <li>待审链接</li>
                        </ul>
                    </div>
                    
                    <div class="Zcontent">
                    	<ul class="toolbar">
                        	<li>
                            	<div class="slideTop">
                                	<span style="float:left">最近新评论</span>
                                    <span style="float:right"><a href="">查看更多评论</a></span>
                                    <div class="clear"></div>
                                </div>
                                <div class="slideContent"><%SlideComment%></div>
                            </li>
                            <li>2</li>
                            <li>3</li>
                        </ul>
                    </div>
               </div>
               
            </div>
            
            <div class="right">
            	<div class="side">
                	<div class="title">
                    	<div class="top"></div>
                        <div class="mid"><div class="tit">PJBlog 官方信息</div></div>
                        <div class="bom"></div>
                    </div>
                    <div class="contList">
                    	<ol>
                        	<li>
                            	<div class="article">
                                	<div class="tis">有点点感冒,自己公司的事都没处理完,估计先处理公司的事.</div>
                                    <div class="info">
                                    	<div class="time">1986-10-31</div>
                                        <div class="read"><a href="">Read</a></div>
                                        <div class="clear"></div>
                                    </div>
                                </div>
                            </li>
                            
                            <li>
                            	<div class="article">
                                	<div class="tis">有点点感冒,自己公司的事都没处理完,估计先处理公司的事.</div>
                                    <div class="info">
                                    	<div class="time">1986-10-31</div>
                                        <div class="read"><a href="">Read</a></div>
                                        <div class="clear"></div>
                                    </div>
                                </div>
                            </li>
                            
                            <li>
                            	<div class="article">
                                	<div class="tis">有点点感冒,自己公司的事都没处理完,估计先处理公司的事.</div>
                                    <div class="info">
                                    	<div class="time">1986-10-31</div>
                                        <div class="read"><a href="">Read</a></div>
                                        <div class="clear"></div>
                                    </div>
                                </div>
                            </li>
                            
                            <li>
                            	<div class="article">
                                	<div class="tis">有点点感冒,自己公司的事都没处理完,估计先处理公司的事.</div>
                                    <div class="info">
                                    	<div class="time">1986-10-31</div>
                                        <div class="read"><a href="">Read</a></div>
                                        <div class="clear"></div>
                                    </div>
                                </div>
                            </li>
                            
                            <li>
                            	<div class="article">
                                	<div class="tis">有点点感冒,自己公司的事都没处理完,估计先处理公司的事.</div>
                                    <div class="info">
                                    	<div class="time">1986-10-31</div>
                                        <div class="read"><a href="">Read</a></div>
                                        <div class="clear"></div>
                                    </div>
                                </div>
                            </li>
                        </ol>
                    </div>
                </div>
                
                <div class="side">
                	<div class="search"></div>
                </div>
                
            </div>
            <div class="clear"></div>
        </div>
	<%
end Sub

Public Sub SlideComment
	Dim Rs, GRA
	Set Rs = Conn.Execute("Select top 5 * From blog_Comment Where comm_IsAudit=False Order By comm_PostTime Desc")
	If Rs.Bof Or Rs.Eof Then
		Response.Write("<p><strong>暂时没有新未审核评论</strong></p>")
	Else
		Set GRA = New Gravatar
%>
	<script language="javascript" type="text/javascript">
		$(function(){
			$(".slideItem_right > .slidePak > .h4").hide();
			$(".slideItem").hover(
				function(){
					$(this).find(".slideItem_right > .slidePak > .h4").show();
				},
				function(){
					$(this).find(".slideItem_right > .slidePak > .h4").hide();
				}
			);
		});
	</script>
<%
		Do While Not Rs.Eof
%>
		<div class="slideItem">
        	<div class="slideItem_left">
            	<%
					GRA.Gravatar_r = "r"
					GRA.Gravatar_EmailMd5 = MD5(Lcase(Trim(Rs("comm_Email").value)))
				%>
                <div class="photo-s"><img src="<%=GRA.outPut%>" /></div>
            </div>
            <div class="slideItem_right">
            	<div class="h6"><span style="float:left"><%=Rs("comm_Author").value%> :</span><span style="float:right; font-size:10px; font-style: oblique; font-weight:normal;"><%=Asp.DateToStr(Rs("comm_PostTime").value, "Y-m-d H:I:S")%></span><div class="clear"></div></div>
                <div class="h5"><%=Rs("comm_Content").value%></div>
                <div style="height:20px;" class="slidePak"><div class="h4">回复 | &nbsp;删除</div></div>
            </div>
            <div class="clear"></div>
        </div>
<%
		Rs.MoveNext
		Loop
		Set GRA = Nothing
	End If
End Sub
%>