<%
'*************** PJblog4 基本设置 *******************
' PJblog4 Copyright 2009
' Update: 2009-11-03
' Author: evio
' Owner : PuterJam
' Mode  : Class
' MoName: Enter Init
'***************************************************
Dim Init : Set Init = New log_Init

Class log_Init
	
	Private GlobalCache
	Public UserID
	
	' ***********************************************
	'	整站初始化类初始化
	' ***********************************************
	Private Sub Class_Initialize
		Call GlobalByte
		Call checkCookies
		Call GlobalRight
    End Sub 
     
	' ***********************************************
	'	整站初始化类终结化
	' ***********************************************
    Private Sub Class_Terminate
    End Sub
	
	' ***********************************************
	'	加载整站变量
	' ***********************************************
	Private Sub GlobalByte
		GlobalCache = Cache.GlobalCache(1)
		SiteName = GlobalCache(0, 0)						'站点名字
        SiteURL = GlobalCache(1, 0)							'站点地址
        blogPerPage = Int(GlobalCache(2, 0))				'每页日志数
        blog_LogNums = Int(GlobalCache(3, 0))				'日志总数
        blog_CommNums = Int(GlobalCache(4, 0))				'评论总数
        blog_MemNums = Int(GlobalCache(5, 0))				'会员总数
        blog_VisitNums = Int(GlobalCache(6, 0))				'访问量
        blogBookPage = Int(GlobalCache(7, 0))				'每页留言数(备用)
        blog_MessageNums = Int(GlobalCache(8, 0))			'留言总数(备用)
        blogcommpage = Int(GlobalCache(9, 0))				'每页评论数
        blogaffiche = GlobalCache(10, 0)					'公告
        blogabout = GlobalCache(11, 0)						'备案信息
        blogcolsize = Int(GlobalCache(12, 0))				'每页书签数(备用)
        blog_colNums = Int(GlobalCache(13, 0))				'书签总数(备用)
        blog_TbCount = Int(GlobalCache(14, 0))				'引用通告总数
        blog_showtotal = CBool(GlobalCache(15, 0))			'是否显示统计(备用)
        Register_UserNames = GlobalCache(16, 0)				'注册名字过滤
        Register_UserName = Split(Register_UserNames, "|")
        FilterIPs = GlobalCache(17, 0)						'IP地址过滤
        FilterIP = Split(FilterIPs, "|")
        blog_commTimerout = Int(GlobalCache(18, 0))			'发表评论时间间隔
        blog_commUBB = Int(GlobalCache(19, 0))				'是否禁用评论UBB代码
        blog_commIMG = Int(GlobalCache(20, 0))				'是否禁用评论贴图
        blog_postFile = GlobalCache(21, 0) 					'动态输出日志文件
        blog_postCalendar = CBool(GlobalCache(22, 0)) 		'动态输出日志日历文件
        blog_DefaultSkin = GlobalCache(23, 0)				'默认界面
        blog_SkinName = GlobalCache(24, 0)					'界面名称
        blog_SplitType = CBool(GlobalCache(25, 0))			'日志分割类型
        blog_introChar = GlobalCache(26, 0)					'日志预览最大字符数
        blog_introLine = GlobalCache(27, 0)					'日志预览切割行数
        blog_validate = CBool(GlobalCache(28, 0))			'发表评论是否都需要验证
        blog_Title = GlobalCache(29, 0)						'Blog副标题
        blog_ImgLink = CBool(GlobalCache(30, 0))			'是否在首页显示图片友情链接
        blog_commLength = Int(GlobalCache(31, 0))			'评论长度
        blog_downLocal = CBool(GlobalCache(32, 0))			'是否使用防盗链下载
        blog_DisMod = CBool(GlobalCache(33, 0))				'默认显示内容
        blog_Disregister = Int(GlobalCache(34, 0))				'是否允许注册
        blog_master = GlobalCache(35, 0)					'blog管理员姓名
        blog_email = GlobalCache(36, 0)						'blog管理员邮件地址
        blog_CountNum = GlobalCache(37, 0)					'访客统计最大次数
        blog_wapNum = Int(GlobalCache(38, 0))				'Wap 文章列表数量
        blog_wapImg = CBool(GlobalCache(39, 0))				'Wap 文章显示图片
        blog_wapHTML = CBool(GlobalCache(40, 0))			'Wap 文章使用简单HTML
        blog_wapLogin = CBool(GlobalCache(41, 0))			'Wap 允许登录
        blog_wapComment = CBool(GlobalCache(42, 0))			'Wap 允许评论
        blog_wap = CBool(GlobalCache(43, 0))				'使用 wap
        blog_wapURL = CBool(GlobalCache(44, 0))				'使用 wap 转换文章超链接
        blog_KeyWords = GlobalCache(45, 0)					'站点首页KeyWords
        blog_Description = GlobalCache(46, 0)				'站点首页Description
        blog_SaveTime = GlobalCache(47, 0)					'Ajax草稿自动保存时间间隔
		blog_IsPing = CBool(GlobalCache(48, 0)) 			'是否开启Ping功能
		blog_html = Trim(GlobalCache(49, 0))				'日志文件后缀列表
		blog_commAduit = GlobalCache(50, 0)					'评论是否需要审核
	End Sub
	
	' ***********************************************
	'	加载整站权限
	' ***********************************************
	Private Sub GlobalRight
		Dim blog_Status_Len, i, blog_Status
		blog_Status = Cache.UserRight(1)
        blog_Status_Len = UBound(blog_Status, 2)
        For i = 0 To blog_Status_Len
            If blog_Status(0, i) = memStatus Then
                stat_title = blog_Status(1, i)
                FillRight blog_Status(2, i)
                UP_FileSize = blog_Status(3, i)
                UP_FileTypes = blog_Status(4, i)
                UP_FileType = Split(UP_FileTypes, "|")
            End If
        Next
	End Sub
	
	' ***********************************************
	'	写入权限变量
	' ***********************************************
	Private Sub FillRight(StatusCode)
		stat_AddAll = CBool(Mid(StatusCode, 1, 1))
		stat_Add = CBool(Mid(StatusCode, 2, 1))
		stat_EditAll = CBool(Mid(StatusCode, 3, 1))
		stat_Edit = CBool(Mid(StatusCode, 4, 1))
		stat_DelAll = CBool(Mid(StatusCode, 5, 1))
		stat_Del = CBool(Mid(StatusCode, 6, 1))
		stat_CommentAdd = CBool(Mid(StatusCode, 7, 1))
		stat_CommentEdit = CBool(Mid(StatusCode, 8, 1))
		stat_CommentDel = CBool(Mid(StatusCode, 9, 1))
		stat_FileUpLoad = CBool(Mid(StatusCode, 10, 1))
		stat_Admin = CBool(Mid(StatusCode, 11, 1))
		stat_ShowHiddenCate = CBool(Mid(StatusCode, 12, 1))
		
		Response.Cookies(Sys.CookieName)("memRight") = StatusCode
		If DateDiff("d",Date(),Request.Cookies(Sys.CookieName)("exp"))>0 Then
			Response.Cookies(Sys.CookieName).Expires = Date + DateDiff("d",Date(),Request.Cookies(Sys.CookieName)("exp"))
		End If
	End Sub
	
	' ***********************************************
	'	验证Cookie记录
	' ***********************************************
	Private Sub checkCookies
		Dim Guest_IP, Guest_Browser, Guest_Refer, SQL
		Guest_IP = Asp.getIP
		Guest_Browser = Asp.getBrowser(Request.ServerVariables("HTTP_USER_AGENT"))
	
		If Session("GuestIP") <> Guest_IP Then
			Conn.Execute("UPDATE blog_Info SET blog_VisitNums=blog_VisitNums+1")
			Cache.GlobalCache(2)
			Session("GuestIP") = Guest_IP
			If blog_CountNum > 0 And Guest_Browser(1) <> "Unkown" Then
				Dim tmpC
				tmpC = Sys.doGet("select count(coun_ID) as cnt from [blog_Counter]")(0)
				Guest_Refer = Trim(Request.ServerVariables("HTTP_REFERER"))
				If tmpC >= blog_CountNum Then
					Dim tmpLC
					tmpLC = Sys.doGet("select top 1 coun_ID from [blog_Counter] order by coun_Time ASC")(0)
					Sys.doGet("update [blog_Counter] set coun_Time=#"&Now()&"#,coun_IP='"&Guest_IP&"',coun_OS='"&Guest_Browser(1)&"',coun_Browser='"&Guest_Browser(0)&"',coun_Referer='"&HTMLEncode(Asp.CheckStr(Guest_Refer))&"' where coun_ID="&tmpLC)
				Else
					Sys.doGet("INSERT INTO blog_Counter(coun_IP,coun_OS,coun_Browser,coun_Referer) VALUES ('"&Guest_IP&"','"&Guest_Browser(1)&"','"&Guest_Browser(0)&"','"&HTMLEncode(Asp.CheckStr(Guest_Refer))&"')")
				End If
			End If
		End If
	
		Dim tempName, tempHashKey
		tempName = Asp.CheckStr(Request.Cookies(Sys.CookieName)("memName"))
		tempHashKey = Asp.CheckStr(Request.Cookies(Sys.CookieName)("memHashKey"))
		If tempHashKey = "" Then
			logout(False)
		Else
			Dim CheckCookie
			Set CheckCookie = Server.CreateObject("ADODB.RecordSet")
			SQL = "SELECT Top 1 mem_ID,mem_Name,mem_Password,mem_salt,mem_Status,mem_LastIP,mem_lastVisit,mem_hashKey FROM blog_member WHERE mem_Name='"&tempName&"' AND mem_hashKey='"&tempHashKey&"' AND mem_hashKey<>''"
			CheckCookie.Open SQL, Conn, 1, 1
			If CheckCookie.EOF And CheckCookie.BOF Then
				logout(False)
			Else
				UserID = CheckCookie("mem_ID")
	'            If CheckCookie("mem_LastIP")<>Guest_IP Or IsNull(CheckCookie("mem_LastIP")) Then
	'                logout(True)
	'            Else
					memName = Asp.CheckStr(Request.Cookies(Sys.CookieName)("memName"))
					memStatus = CheckCookie("mem_Status")
	'            End If
			End If
			CheckCookie.Close
			Set CheckCookie = Nothing
		End If
	
	End Sub
	
	' ***********************************************
	'	退出登入
	' ***********************************************
	Public Sub logout(clearHashKey)
		On Error Resume Next
		Response.Cookies(Sys.CookieName)("DisValidate") = "False"
		If clearHashKey Then Sys.doGet("UPDATE blog_member set mem_hashKey='' where mem_ID=" & UserID)
		If Err Then Err.Clear
		Response.Cookies(Sys.CookieName)("memName") = ""
		Response.Cookies(Sys.CookieName)("memHashKey") = ""
		memStatus = "Guest"
	End Sub
	
	' ***********************************************
	'	退出后台
	' ***********************************************
	Public Sub c_Logout
		session(Sys.CookieName&"_System") = ""
		session(Sys.CookieName&"_disLink") = ""
		session(Sys.CookieName&"_disCount") = ""
		Response.Write ("<script>try{top.location=""default.asp""}catch(e){location=""default.asp""}</script>")
	End Sub
	
	' ***********************************************
	'	组成日志也连接(后台)
	' ***********************************************
	Public Function ArticleUrl(Deep, cCate, cname, ctype)
		If Len(cCate) = 0 Then
			ArticleUrl = Deep & "html/" & cname & "." & ctype
		Else
			ArticleUrl = Deep & "html/" & cCate & "/" & cname & "." & ctype
		End If
	End Function
	
	' ***********************************************
	'	组成日志页连接(前台)
	' ***********************************************
	Public Function doArticleUrl(ByVal c_id)
		If Not Asp.IsInteger(c_id) Then Exit Function
		Dim Rs, cname, ctype, cfolder
		Set Rs = Conn.Execute("Select Top 1 T.log_cname, T.log_ctype, B.cate_Folder From blog_Content As T, blog_Category As B Where  T.log_CateID=B.cate_ID And T.log_ID=" & c_id)
		If Rs.Bof Or Rs.Eof Then
			Exit Function
		Else
			cname = Rs(0).value
			ctype = Rs(1).value
			cfolder = Rs(2).value
			If Len(cfolder) = 0 Then
				doArticleUrl = cname & "." & ctype
			Else
				doArticleUrl = cfolder & "/" & cname & "." & ctype
			End If
		End If
		Set Rs = Nothing
	End Function
	
	' ***********************************************
	'	组成分类页连接(前台)
	' ***********************************************
	Public Function doCategoryUrl(ByVal c_id)
		If Not Asp.IsInteger(c_id) Then Exit Function
		Dim Rs, cname, ctype, cfolder
		Set Rs = Conn.Execute("Select Top 1 cate_Folder, cate_OutLink, cate_URL From blog_Category Where cate_ID=" & c_id)
		If Rs.Bof Or Rs.Eof Then
			Exit Function
		Else
			If Rs(1).value Then
				doCategoryUrl = Rs(2).value
			Else
				If Len(Rs(0).value) > 0 Then
					doCategoryUrl = "cate_" & Rs(0).value & "_1.html"
				Else
					doCategoryUrl = "cate_" & c_id & "_1.html"
				End If
			End If
		End If
		Set Rs = Nothing
	End Function
	
	'**********************************************
	'	获取在线人数
	'**********************************************
	
	Public Function getOnline
		getOnline = 1
		If Len(Application(Sys.CookieName&"_onlineCount"))>0 Then
			If DateDiff("s", Application(Sys.CookieName&"_userOnlineCountTime"), Now())>60 Then
				Application.Lock()
				Application(Sys.CookieName&"_online") = Application(Sys.CookieName&"_onlineCount")
				Application(Sys.CookieName&"_onlineCount") = 1
				Application(Sys.CookieName&"_onlineCountKey") = Asp.randomStr(2)
				Application(Sys.CookieName&"_userOnlineCountTime") = Now()
				Application.Unlock()
			Else
				If Session(Sys.CookieName&"userOnlineKey")<>Application(Sys.CookieName&"_onlineCountKey") Then
					Application.Lock()
					Application(Sys.CookieName&"_onlineCount") = Application(Sys.CookieName&"_onlineCount") + 1
					Application.Unlock()
					Session(Sys.CookieName&"userOnlineKey") = Application(Sys.CookieName&"_onlineCountKey")
				End If
			End If
		Else
			Application.Lock
			Application(Sys.CookieName&"_online") = 1
			Application(Sys.CookieName&"_onlineCount") = 1
			Application(Sys.CookieName&"_onlineCountKey") = Asp.randomStr(2)
			Application(Sys.CookieName&"_userOnlineCountTime") = Now()
			Application.Unlock
		End If
		getOnline = Application(Sys.CookieName&"_online")
	End Function
	
	Public Sub HeadNav
		Dim NavArray, i
		NavArray = Data.NavList(1)
		If UBound(NavArray, 1) = 0 Then
			Exit Sub
		Else
			Response.Write("<ul><li class=""menuL""></li>")
			For i = 0 To UBound(NavArray, 2)
				If NavArray(3, i) Then
					Response.Write("<li><a class=""menuA menuB"" href=""../html/" & NavArray(4, i) & """ title=""" & NavArray(2, i) & """>" & NavArray(1, i) & "</a></li><li class=""menuDiv""></li>")
				Else
					If Len(NavArray(5, i)) > 0 Then
						Response.Write("<li><a class=""menuA menuB"" href=""../html/cate_" & NavArray(5, i) & "_1.html"" title=""" & NavArray(2, i) & """>" & NavArray(1, i) & "</a></li><li class=""menuDiv""></li>")
					Else
						Response.Write("<li><a class=""menuA menuB"" href=""../html/cate_" & NavArray(0, i) & "_1.html"" title=""" & NavArray(2, i) & """>" & NavArray(1, i) & "</a></li><li class=""menuDiv""></li>")
					End If
				End If
			Next
			Response.Write("<li class=""menuR""></li></ul>")
		End If
	End Sub
	
End Class
%>