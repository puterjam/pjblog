<%
'==========================后台信息处理===============================

Sub doAction
    Dim weblog, WebFso
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
			weblog("blog_KeyWords") = checkURL(CheckStr(Request.Form("blog_KeyWords")))
			weblog("blog_Description") = checkURL(CheckStr(Request.Form("blog_Description")))
    '评论留言邮件通知
            weblog("blog_smtp")=trim(Request.form("blog_smtp"))
            weblog("blog_smtpuser")=trim(Request.form("blog_smtpuser"))
            weblog("blog_smtppassword")=trim(Request.form("blog_smtppassword"))
            weblog("blog_jmail")=trim(Request.form("blog_jmail"))
            weblog("blog_smtpmail")=trim(Request.form("blog_smtpmail"))
			weblog("blog_html") = Trim(Request.Form("blog_html"))
            If Int(Request.Form("blog_Isjmail")) = 1 Then 
                weblog("blog_Isjmail")=1
            Else
                weblog("blog_Isjmail")=0
            End If
            If Int(Request.Form("blog_reply_Isjmail")) = 1 Then 
                weblog("blog_reply_Isjmail")=1
            Else
                weblog("blog_reply_Isjmail")=0
            End If
    '评论留言邮件通知结束

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
			If Len(Request.Form("blog_SaveTime")) > 0 Then
	            weblog("blog_SaveTime") = CheckStr(Request.Form("blog_SaveTime"))
			Else
	            weblog("blog_SaveTime") = 3600
			End If
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
			If CheckStr(Request.Form("blog_PasswordProtection")) = "1" Then weblog("blog_PasswordProtection") = True Else weblog("blog_PasswordProtection") = False
			If CheckStr(Request.Form("AuditOpen")) = "1" Then weblog("blog_AuditOpen") = True Else weblog("blog_AuditOpen") = False
			If CheckStr(Request.Form("GravatarOpen")) = "1" Then weblog("blog_GravatarOpen") = True Else weblog("blog_GravatarOpen") = False
			If CheckStr(Request.Form("PingOpen")) = "1" Then weblog("blog_IsPing") = True Else weblog("blog_IsPing") = False

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
			' ping 提交
        ElseIf Request.Form("whatdo") = "pingupdate" Then
			Dim ping_name, ping_url, Ping_Update_Class, ping_up_i, ping_id
			Dim new_ping_name, new_ping_url
			ping_name = Split(CheckStr(Request.Form("ping_name")), ",")
			ping_url = Split(CheckStr(Request.Form("ping_url")), ",")
			ping_id = Split(CheckStr(Request.Form("ping_id")), ",")
			new_ping_name = CheckStr(Request.Form("new_ping_name"))
			new_ping_url = CheckStr(Request.Form("new_ping_url"))
			If UBound(ping_name) >= 0 Then
				Set Ping_Update_Class = New ping
					For ping_up_i = 0 To UBound(ping_name)
						Ping_Update_Class.Ping_Name = ping_name(ping_up_i)
						Ping_Update_Class.Ping_Url = ping_url(ping_up_i)
						Ping_Update_Class.updatePingBase(ping_id(ping_up_i))
					Next
				Set Ping_Update_Class = Nothing
			End If
			If Len(new_ping_name) > 0 Then
				Set Ping_Update_Class = New ping
					Ping_Update_Class.Ping_Name = new_ping_name
					Ping_Update_Class.Ping_Url = new_ping_url
					Ping_Update_Class.insPingBase
				Set Ping_Update_Class = Nothing
			End If
			session(CookieName&"_ShowMsg") = True
            session(CookieName&"_MsgText") = "保存Ping服务信息成功!"
            RedirectUrl("ConContent.asp?Fmenu=General&Smenu=Ping")
			' ping 删除
        ElseIf Request.Form("whatdo") = "pingdelete" Then
			Dim  ping_SelectId, Ping_Del_Class, Ping_Del_i
			ping_SelectId = Split(CheckStr(Request.Form("pingSelectID")), ",")
			If UBound(ping_SelectId) >= 0 Then
				Set Ping_Del_Class = New ping
					For Ping_Del_i = 0 To UBound(ping_SelectId)
						Ping_Del_Class.DelPingBase(ping_SelectId(Ping_Del_i))
					Next
				Set Ping_Del_Class = Nothing
				session(CookieName&"_ShowMsg") = True
            	session(CookieName&"_MsgText") = "删除" & (UBound(ping_SelectId) + 1) & "条Ping服务信息成功!"
			End If
            RedirectUrl("ConContent.asp?Fmenu=General&Smenu=Ping")
		ElseIf Request.Form("whatdo") = "Misc" Then
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

            If Request.Form("ReBulidIndex") = 1 Then
                Dim lArticle
                Set lArticle = New ArticleCache
                lArticle.SaveCache
                Set lArticle = Nothing
                session(CookieName&"_ShowMsg") = True
                session(CookieName&"_MsgText") = Session(CookieName&"_MsgText")&"重新输出日志索引! "
            End If

			'分段静态 - evio
			If Request.Form("ReBulidPartStatus") = 1 Then
			   Dim ReBulidPartStatusSart, ReBulidPartStatusEnd, ReArtLists, ReArea, TI
			   ReBulidPartStatusSart = trim(Request.Form("ReBulidPartStatusSart"))
			   ReBulidPartStatusEnd = trim(Request.Form("ReBulidPartStatusEnd"))
			   If (len(ReBulidPartStatusSart) > 0) and (len(ReBulidPartStatusEnd) > 0) and (Int(ReBulidPartStatusEnd) >= Int(ReBulidPartStatusSart)) then
			      ReArtLists = PartStatus(ReBulidPartStatusSart, ReBulidPartStatusEnd)
				  ReArea = Split(ReArtLists, "|")
				  for TI = 0 to (UBound(ReArea)-1)
				      If Request.Form("ReBulidArticle") = 2 Then
					     PostArticle ReArea(TI), True
					  Else
				         PostArticle ReArea(TI), False
					  End if
				  next
				  session(CookieName&"_ShowMsg") = True
                  session(CookieName&"_MsgText") = Session(CookieName&"_MsgText")&"分段静态 "&UBound(ReArea)&" 篇日志成功!"
			   Else
			      session(CookieName&"_ShowMsg") = True
                  session(CookieName&"_MsgText") = Session(CookieName&"_MsgText")&"<font color=red>请填写分段ID的开始和结束,并且结束ID比开始ID大!</font>"
			   End If
			End If
			'分段静态 - evio

            If Request.Form("ReBulidArticle") = 1 Or Request.Form("ReBulidArticle") = 2 Then
                Dim LoadArticle, LoadArticleCount, LogLen, silent
                Set LoadArticleCount = conn.Execute("SELECT count(log_ID) FROM blog_Content")
                LogLen = LoadArticleCount(0).Value
                Set LoadArticleCount = Nothing
                silent = Request.Form("silentMode")
                If silent <> 1 Then


%>
			<h5 style="margin-bottom:4px;">正在静态化相关的日志缓存... 共<%=LogLen%>篇日志需要处理. <span style="color:#f00">静态化过程，请不要关闭您的浏览器</span></h5>
			<div style="background:#fff;padding:1px;border:1px solid #333">
				<div id="per" style="width:0%;background:#0000a0 url(images/per.png);overflow:hidden;color:#fff;font-weight:bold;padding:2px;font-family:verdana;font-size:12px;white-space:nowrap">0%</div>
			</div>
			<script>
				var _t = <%=LogLen%>;
				var _p = document.getElementById("per");
				var _cc = document.getElementById("cc");
				var d = new Date();
				function _s(c) {
					var p = parseInt((c/_t)*100);
					var d1 = new Date();
					_p.style.width = p + "%";
					_p.innerHTML = c + " of " + _t + " : " + p + "%" + " : " + ((d1 -d)/1000) + "s";
					//_cc.innerHTML = c;
				}
			</script>
		<%
End If
Application.Lock
Application(CookieName & "_SiteEnable") = 0
Application(CookieName & "_SiteDisbleWhy") = "抱歉!网站在初始化数据，请稍后在访问. :P"
Application.UnLock



Set LoadArticle = conn.Execute("SELECT log_ID FROM blog_Content")
Dim iLen
Do Until LoadArticle.EOF
    If Request.Form("ReBulidArticle") = 2 Then
        PostArticle LoadArticle("log_ID"), True
    Else
        PostArticle LoadArticle("log_ID"), False
    End If

    If silent <> 1 Then
        iLen = iLen + 1
        Response.Write ("<script>_s(" & iLen & ");</script>")
        Response.Flush()
    End If

    LoadArticle.movenext
Loop

Application.Lock
Application(CookieName & "_SiteEnable") = 1
Application(CookieName & "_SiteDisbleWhy") = ""
Application.UnLock

session(CookieName&"_ShowMsg") = True
session(CookieName&"_MsgText") = Session(CookieName&"_MsgText")&"共处理了 "&LogLen&" 篇日志文件! "

If silent<>1 Then
    Response.Write ("<script>setTimeout(""location = 'ConContent.asp?Fmenu=General&Smenu=Misc'"",2000);</script>")
    Response.Write ("<div><b style='color:#040;font-size:12px;margin:4px;'>静态化完成... </b></div>")
Else
    RedirectUrl("ConContent.asp?Fmenu=General&Smenu=Misc")
End If
Else
    RedirectUrl("ConContent.asp?Fmenu=General&Smenu=Misc")
End If



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
        Cate_source_name = conn.Execute("select cate_Part from blog_Category where cate_ID="&Cate_source)(0)
        Cate_source_count = conn.Execute("select cate_count from blog_Category where cate_ID="&Cate_source)(0)
        Cate_target_name = conn.Execute("select cate_Part from blog_Category where cate_ID="&Cate_target)(0)
		' ***********************************
		'	批量移动开始
		' ***********************************
		If blog_postFile = 2 Then
			Set WebFso = New cls_FSO
				Dim ArticlesID, Articlei, ArticleSplit
				ArticlesID = WebFso.FileItem("article/" & Cate_source_name & "/")
				ArticleSplit = Split(ArticlesID, "|")
				If Int(ArticleSplit(0)) > 0 Then
					For Articlei = 1 To UBound(ArticleSplit)
						WebFso.MoveFile "article/" & Cate_source_name & "/" & ArticleSplit(Articlei), "article/" & Cate_target_name & "/" & ArticleSplit(Articlei)
					Next
				End If
			Set WebFso = Nothing
		end if
		' ***********************************
		'	批量移动结束
		' ***********************************
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
        Dim LCate_ID, LCate_icons, LCate_Name, LCate_Intro, Lcate_URL, Lcate_Order, Lcate_count, LCate_local, Lcate_Secret, LCate_Part, LOldCate_Name
        Dim NCate_ID, NCate_icons, NCate_Name, NCate_Intro, Ncate_URL, Ncate_Order, NCate_local, Ncate_OutLink, Ncate_Secret, Ncate_Part
        LCate_ID = Split(Request.Form("Cate_ID"), ", ")
        LCate_Name = Split(Request.Form("Cate_Name"), ", ")
		' New Add
		If blog_postFile = 2 Then
			LCate_Part = Split(Request.Form("Cate_Part"), ", ")
			LOldCate_Name = Split(Request.Form("OldCate_Name"), ", ")
		end if
		' New Add
        LCate_icons = Split(Request.Form("Cate_icons"), ", ")
        LCate_Intro = Split(Request.Form("Cate_Intro"), ", ")
        Lcate_URL = Split(Request.Form("cate_URL"), ", ")
        Lcate_Order = Split(Request.Form("cate_Order"), ", ")
        Lcate_count = Split(Request.Form("cate_count"), ", ")
        LCate_local = Split(Request.Form("Cate_local"), ", ")
        Lcate_Secret = Split(Request.Form("cate_Secret"), ", ")
        For i = 0 To UBound(LCate_Name)
		' ***********************************
		'	判断是否改变了文件夹名 开始
		' ***********************************
			If blog_postFile = 2 Then
				Set WebFso = New cls_FSO
				If LCate_Part(i) <> LOldCate_Name(i) Then
					If LOldCate_Name(i) = "" Then
						WebFso.CreateFolder("article/" & LCate_Part(i) & "/")
						Dim ChangDb, dbDress, NewDress
						Set ChangDb = Conn.Execute("Select * From blog_Content where log_CateID=" & LCate_ID(i))
						If Not (ChangDb.Bof And ChangDb.Eof) Then
							Do While Not ChangDb.Eof
								dbDress = Alias(ChangDb("log_ID"))
								NewDress = "article/" & LCate_Part(i) & "/" & Replace(dbDress, "article/", "")
								WebFso.MoveFile dbDress, NewDress
							ChangDb.MoveNext
							Loop
						End If
						Set ChangDb = Nothing
					Else
						WebFso.FolderRename "article/" & LOldCate_Name(i), LCate_Part(i)
					End If
				End If
				Set WebFso = Nothing
				Session(CookieName&"_MsgText") = "请到 <a href=""ConContent.asp?Fmenu=General&Smenu=Misc"" style=""color:#00f"">站点基本设置-初始化数据</a> ,重新生成所有日志到文件。 "
			End if
		    ' ***********************************
			'	判断是否改变了文件夹名 结束
			' ***********************************
            SQL = "SELECT * FROM blog_Category where cate_ID="&Int(CheckStr(LCate_ID(i)))
            weblog.Open SQL, Conn, 1, 3
            weblog("cate_Name") = CheckStr(LCate_Name(i))
			If blog_postFile = 2 Then
				weblog("Cate_Part") = left(FilterHtmlTags(CheckStr(LCate_Part(i))),50)
			End if
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
		If blog_postFile = 2 Then
			Ncate_Part = left(FilterHtmlTags(Trim(CheckStr(Request.Form("New_Cate_Part")))),50)
		end if
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
			If blog_postFile = 2 Then
            AddCateArray = Array(Array("cate_Name", NCate_Name), Array("cate_icon", NCate_icons), Array("Cate_Intro", NCate_Intro), Array("cate_URL", Ncate_URL), Array("cate_OutLink", Ncate_OutLink), Array("cate_Order", Int(Ncate_Order)), Array("Cate_local", NCate_local), Array("Cate_Secret", NCate_Secret), Array("Cate_Part", Ncate_Part))
			else
			AddCateArray = Array(Array("cate_Name", NCate_Name), Array("cate_icon", NCate_icons), Array("Cate_Intro", NCate_Intro), Array("cate_URL", Ncate_URL), Array("cate_OutLink", Ncate_OutLink), Array("cate_Order", Int(Ncate_Order)), Array("Cate_local", NCate_local), Array("Cate_Secret", NCate_Secret))
			end if
            If DBQuest("blog_Category", AddCateArray, "insert") = 0 Then session(CookieName&"_MsgText") = session(CookieName&"_MsgText")&"新日志分类添加成功，"
			if Ncate_URL="" or Ncate_URL=empty or len(Ncate_URL)=0 then
			If blog_postFile = 2 Then
				Set WebFso = New cls_FSO
					If Len(Ncate_Part) > 0 Then WebFso.CreateFolder("article/" & Ncate_Part & "/")
				Set WebFso = Nothing
			End if
			End if
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
		Set WebFso = New cls_FSO
        Set DelLog = Conn.Execute("select log_ID FROM blog_Content WHERE log_CateID=" & DelCate)
        Do While Not DelLog.EOF
            DelID = DelLog("log_ID")
            If DeleteLog(DelID) = 1 Then
                P1 = P1 + 1
            Else
                P2 = P2 + 1
            End If
            DelLog.movenext
        Loop
		If blog_postFile = 2 Then
			Dim us : us = conn.execute("select cate_URL from blog_Category where cate_ID="&DelCate)(0)
			if (len(us) = 0) or (us = "") or (us = empty) or (us = null) then
				Dim uPa : uPa = Conn.execute("select cate_Part from blog_Category where cate_ID="&DelCate)(0)
				If Len(uPa) > 0 Then WebFso.DeleteFolder "article/" & Conn.execute("select cate_Part from blog_Category where cate_ID="&DelCate)(0), True
			end if
		End if
        Conn.Execute("DELETE * FROM blog_Category WHERE cate_ID="&DelCate)
		Set WebFso = Nothing
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

    Dim selCommID, doCommID, doTitle, doRedirect, t1, t2, bookDB, commDB, doCommentAudit, doAuditTrue
    selCommID = Split(Request.Form("selectCommentID"), ", ")
    doCommID = Split(Request.Form("CommentID"), ", ")
	doCommentAudit = Split(Request.Form("CommentAudit"), ", ")

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
                Set bookDB = Conn.Execute("select book_Messager FROM blog_book WHERE book_ID="&selCommID(i))
                conn.Execute("update blog_Info set blog_MessageNums=blog_MessageNums-1")
                conn.Execute("update blog_Member set mem_PostMessageNums=mem_PostMessageNums-1 where mem_Name='"&bookDB("book_Messager")&"'")
                conn.Execute("DELETE * from blog_book where book_ID="&selCommID(i))
                doTitle = "留言"
                doRedirect = "msg"
            Else
                t1 = Int(Split(selCommID(i), "|")(0))
                t2 = Int(Split(selCommID(i), "|")(1))
                Set commDB = Conn.Execute("select comm_Author FROM blog_Comment WHERE comm_ID="&t1)
                conn.Execute("update blog_Content set log_CommNums=log_CommNums-1 where log_ID="&t2)
                conn.Execute("update blog_Info set blog_CommNums=blog_CommNums-1")
                conn.Execute("update blog_Member set mem_PostComms=mem_PostComms-1 where mem_Name='"&commDB("comm_Author")&"'")
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
                    conn.Execute("UPDATE blog_book SET book_Content='"&checkStr(Request.Form("message_"&doCommID(i)))&"',book_reply='"&checkStr(Request.Form("reply_"&doCommID(i)))&"' WHERE book_ID="&doCommID(i))
                End If
                doTitle = "留言"
                doRedirect = "msg"
            ElseIf Request.Form("whatdo") = "comment" Then
				If Int(doCommentAudit(i)) = 0 Then
					doAuditTrue = True
				ElseIf Int(doCommentAudit(i)) = 1 Then
					doAuditTrue = False
				Else
					session(CookieName&"_ShowMsg") = True
        			session(CookieName&"_MsgText") = "参数错误"
        			RedirectUrl("ConContent.asp?Fmenu=Comment&Smenu="&doRedirect)
				End If
                conn.Execute("UPDATE blog_Comment SET comm_Content='"&checkStr(Request.Form("message_"&doCommID(i)))&"',comm_IsAudit=" & doAuditTrue & " WHERE comm_ID="&doCommID(i))
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
		'--------------------------批量删除用户----------------------------
    ElseIf Request.Form("whatdo") = "DelUserAll" Then
	    Dim DelSelectUserID, UseI
		DelSelectUserID = Split(request.form("mem_CheckBox"), ",")
		if int(ubound(DelSelectUserID)) >= 0 then
		   For UseI = 0 to UBound(DelSelectUserID)
		       conn.Execute("delete * from blog_Member where mem_ID="&DelSelectUserID(UseI))
		   next
		   Conn.Execute("UPDATE blog_Info SET blog_MemNums=blog_MemNums-"&UBound(DelSelectUserID)+1&"")
		   getInfo(2)
		   session(CookieName&"_ShowMsg") = True
           session(CookieName&"_MsgText") = "<span style=""color:#900"">"&UBound(DelSelectUserID)+1&" 个用户资料</span> 删除成功!"
           RedirectUrl("ConContent.asp?Fmenu=Members&Smenu=Users")
		Else
		   session(CookieName&"_ShowMsg") = True
           session(CookieName&"_MsgText") = "您没有选择要删除的用户!"
           RedirectUrl("ConContent.asp?Fmenu=Members&Smenu=Users")
		End If
    Else
        session(CookieName&"_ShowMsg") = True
        session(CookieName&"_MsgText") = "非法提交内容!"
        RedirectUrl("ConContent.asp?Fmenu=Members&Smenu=")
    End If
    '==========================友情链接管理===============================
ElseIf Request.Form("action") = "Links" Then
    Dim LinkID, LinkName, LinkURL, LinkLogo, LinkOrder, LinkMain, LinkClass
	Dim LinkClassID, LinkClassName, LinkClassOrder, LinkClassTitle
    '--------------------------友情链接过滤----------------------------
    If Request.Form("whatdo") = "Filter" Then
        session(CookieName&"_disLink") = CheckStr(Request.Form("disLink"))
        session(CookieName&"_disCount") = CheckStr(Request.Form("disCount"))
		session(CookieName&"_disLinkClassID") = CheckStr(Request.Form("disLinkClass"))
        RedirectUrl("ConContent.asp?Fmenu=Link&Smenu=")
        '--------------------------保存友情链接----------------------------
    ElseIf Request.Form("whatdo") = "SaveLink" Then
        Dim TLinkName, TLinkURL, TLinkLogo, TLinkOrder, TLinkClass
        LinkID = Split(Request.Form("LinkID"), ", ")
        LinkName = Split(Request.Form("LinkName"), ", ")
        LinkURL = Split(Request.Form("LinkURL"), ", ")
        LinkLogo = Split(Request.Form("LinkLogo"), ", ")
        LinkOrder = Split(Request.Form("LinkOrder"), ", ")
		LinkClass = Split(Request.Form("LinkClass"), ", ")
        For i = 0 To UBound(LinkID)
            If UBound(LinkName)<0 Then TLinkName = "未知" Else TLinkName = LinkName(i)
            If UBound(LinkURL)<0 Then TLinkURL = "http://" Else TLinkURL = LinkURL(i)
            If UBound(LinkLogo)<0 Then TLinkLogo = "" Else TLinkLogo = LinkLogo(i)
            If UBound(LinkOrder)<0 Then TLinkOrder = "0" Else TLinkOrder = LinkOrder(i)
            conn.Execute("update blog_Links set link_Name='"&CheckStr(TLinkName)&"',link_URL='"&CheckStr(TLinkURL)&"',link_Image='"&CheckStr(TLinkLogo)&"',link_Order='"&CheckStr(TLinkOrder)&"',Link_ClassID="&LinkClass(i)&" where link_ID="&LinkID(i))
        Next
        LinkID = Request.Form("new_LinkID")
        LinkName = Request.Form("new_LinkName")
        LinkURL = Request.Form("new_LinkURL")
        LinkLogo = Request.Form("new_LinkLogo")
        LinkOrder = Request.Form("new_LinkOrder")
		LinkClass = Request.Form("new_LinkClass")
        If Len(LinkOrder)<1 Then LinkOrder = conn.Execute("select count(*) from blog_Links")(0)
        If Len(Trim(CheckStr(LinkName)))>0 Then
            conn.Execute("insert into blog_Links (link_Name,link_URL,link_Image,link_Order,link_IsShow,Link_ClassID) values ('"&CheckStr(LinkName)&"','"&CheckStr(LinkURL)&"','"&CheckStr(LinkLogo)&"','"&CheckStr(LinkOrder)&"',true,"&CheckStr(LinkClass)&")")
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
	    '--------------------------批量删除友情链接----------------------------
	ElseIf Request.Form("whatdo") = "DelLinks" Then
        Dim LinksID, PartLinks, LinksNum
        LinksID = Request.Form("checklinkID")
        PartLinks = split(LinksID, ", ")
        If int(ubound(PartLinks)) >= 0 then
           For LinksNum = 0 to ubound(PartLinks)
               conn.Execute("DELETE * from blog_Links where link_ID="&PartLinks(LinksNum))
           Next
        Session(CookieName&"_ShowMsg") = True
        Session(CookieName&"_MsgText") = session(CookieName&"_MsgText")&"共"&ubound(PartLinks)+1&"条链接删除成功!"
        Bloglinks(2)
        PostLink
        RedirectUrl("ConContent.asp?Fmenu=Link&Smenu=&page="&Request.Form("page"))
        Else
        Session(CookieName&"_ShowMsg") = True
        Session(CookieName&"_MsgText") = session(CookieName&"_MsgText")&"您没有选择要删除的链接!"
        RedirectUrl("ConContent.asp?Fmenu=Link&Smenu=&page="&Request.Form("page"))
        End If
		'--------------------------批量取消通过友情链接----------------------------
	ElseIf Request.Form("whatdo") = "ReSetLiks" Then
		Dim ReSetSelectID, SplitReSetSelectID, SplitReSetCout
		ReSetSelectID = Request.Form("checklinkID")
		SplitReSetSelectID = Split(ReSetSelectID, ",")
		If UBound(SplitReSetSelectID) >= 0 Then
			For SplitReSetCout = 0 To UBound(SplitReSetSelectID)
				Conn.Execute("Update blog_Links Set link_IsShow=False Where link_ID=" & SplitReSetSelectID(SplitReSetCout))
			Next
			Bloglinks(2)
        	PostLink
		End If
		Session(CookieName&"_ShowMsg") = True
        Session(CookieName&"_MsgText") = session(CookieName&"_MsgText")&"批量取消友情链接成功!"
		RedirectUrl("ConContent.asp?Fmenu=Link&Smenu=&page="&Request.Form("page"))
    '/////////////////////////////////////////////////////////////////////////////
	'LinkClassID, LinkClassName, LinkClassOrder, LinkClassTitle
		'--------------------------保存友情链接分类----------------------------
    ElseIf Request.Form("whatdo") = "LinkClassUpdate" Then

		Dim TLinkClassID, TLinkClassName, TLinkClassTitle, TLinkClassOrder
		LinkClassID = Split(Request.Form("LinkClass_ID"), ",")
		LinkClassName = Split(Request.Form("LinkClass_Name"), ",")
		LinkClassTitle = Split(Request.Form("LinkClass_Title"), ",")
		LinkClassOrder = Split(Request.Form("LinkClass_Order"), ",")
		
		For i = 0 To UBound(LinkClassID)
            If UBound(LinkClassName) < 0 Then TLinkClassName = "未知" Else TLinkClassName = LinkClassName(i)
            If UBound(LinkClassTitle) < 0 Then TLinkClassTitle = "没有链接分类说明" Else TLinkClassTitle = LinkClassTitle(i)
            If UBound(LinkClassOrder) < 0 Then TLinkClassOrder = "0" Else TLinkClassOrder = LinkClassOrder(i)
            conn.Execute("update blog_LinkClass set LinkClass_Name='"&CheckStr(TLinkClassName)&"',LinkClass_Title='"&CheckStr(TLinkClassTitle)&"',LinkClass_Order="&CheckStr(TLinkClassOrder)&" where LinkClass_ID="&LinkClassID(i))
        Next
		
        LinkClassName = Request.Form("New_LinkClass_Name")
        LinkClassTitle = Request.Form("New_LinkClass_Title")
        LinkClassOrder = Request.Form("New_LinkClass_Order")
        If Len(LinkClassOrder)<1 Then LinkClassOrder = Int(Int(conn.Execute("select count(*) from blog_LinkClass")(0)) + 1)
        If Len(Trim(CheckStr(LinkClassName)))>0 Then
            conn.Execute("insert into blog_LinkClass (LinkClass_Name,LinkClass_Title,LinkClass_Order) values ('" & LinkClassName & "', '" & LinkClassTitle & "', " & LinkClassOrder & ")")
            session(CookieName&"_MsgText") = "新友情链接分类添加成功! "
        End If
		session(CookieName&"_ShowMsg") = True
        session(CookieName&"_MsgText") = session(CookieName&"_MsgText")&"保存链接分类成功!"
        RedirectUrl("ConContent.asp?Fmenu=Link&Smenu=LinkClass")
		'--------------------------保存友情链接分类----------------------------
    ElseIf Request.Form("whatdo") = "DelLinkClasses" Then
		Dim LinkClassSelectID, LinkClassNum
		LinkClassSelectID = Split(Request.Form("SelectID"), ",")
		If Int(UBound(LinkClassSelectID)) >= 0 then
           For LinkClassNum = 0 to UBound(LinkClassSelectID)
               conn.Execute("DELETE * from blog_LinkClass where LinkClass_ID="&LinkClassSelectID(LinkClassNum))
           Next
        	Session(CookieName&"_ShowMsg") = True
        	Session(CookieName&"_MsgText") = session(CookieName&"_MsgText")&"共"&ubound(LinkClassSelectID) + 1&"条链接分类删除成功!"
        	RedirectUrl("ConContent.asp?Fmenu=Link&Smenu=LinkClass")
		End If
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
            Dim key_Image
            KeyWordID = Split(Request.Form("KeyWordID"), ", ")
            KeyWord = Split(Request.Form("KeyWord"), ", ")
            KeyWordURL = Split(Request.Form("KeyWordURL"), ", ")
            key_Image = Split(Request.Form("key_Image"), ", ")
            For i = 0 To UBound(KeyWordID)
                If Int(KeyWordID(i))<> -1 Then
                    conn.Execute("update blog_Keywords set key_Text='"&CheckStr(KeyWord(i))&"',key_URL='"&CheckStr(KeyWordURL(i))&"',key_Image='"&CheckStr(key_Image(i))&"' where key_ID="&KeyWordID(i))
                Else
                    If Len(Trim(CheckStr(KeyWord(i))))>0 Then
                        conn.Execute("insert into blog_Keywords (key_Text,key_URL,key_Image) values ('"&CheckStr(KeyWord(i))&"','"&CheckStr(KeyWordURL(i))&"','"&CheckStr(key_Image(i))&"')")
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
		EmptyEtag
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
		EmptyEtag
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
            session(CookieName&"_MsgText") = "<span style=""color:#900"">“"&conn.Execute("select title from blog_module where id="&moduleID)(0)&"”</span> 是内置模块无法删除!"
            RedirectUrl("ConContent.asp?Fmenu=Skins&Smenu=module")
        Else
            moduleName = conn.Execute("select title from blog_module where id="&moduleID)(0)
            delModule moduleID
            session(CookieName&"_ShowMsg") = True
            session(CookieName&"_MsgText") = "<span style=""color:#900"">“"&moduleName&"”</span> 删除成功!"
            log_module(2)
			EmptyEtag
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
		EmptyEtag
        session(CookieName&"_ShowMsg") = True
        session(CookieName&"_MsgText") = "<span style=""color:#900"">“"&GetPlugName&"”</span> 设置保存成功!"
        RedirectUrl("ConContent.asp?Fmenu=Skins&Smenu=PluginsOptions&Plugins="&GetPlugName)
    End If
    '==========================附件管理===============================
ElseIf Request.Form("action") = "Attachments" Then
    '-------------------------------删除附件------------------------------
    If Request.Form("whatdo") = "DelFiles" Then
        Dim getFolders, getFiles, GetFolder, GetFile, getFolderCount, getFileCount
		Set WebFso = New cls_FSO
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
            If WebFso.DeleteFolder(GetFolder) Then
                getFolderCount = getFolderCount + 1
            End If
        Next
        For Each GetFile in getFiles
            If WebFso.DeleteFile(GetFile) Then
                getFileCount = getFileCount + 1
            End If
        Next
		Set WebFso = Nothing
        session(CookieName&"_ShowMsg") = True
        session(CookieName&"_MsgText") = "有 <span style=""color:#900"">"&getFileCount&" 文件, "&getFolderCount&" 个文件夹</span> 被删除!"
        RedirectUrl("ConContent.asp?Fmenu=SQLFile&Smenu=Attachments")
    End If
ElseIf Request.Form("action") = "Attachment2" Then
       Dim FilesID, FilesPath, FilesCounts, i
       If Request.form("S_Action")="DelSelect" then
            FilesID = split(Request.form("SelectFilesID"),", ")
            For i = 0 to ubound(FilesID)
                Conn.execute("Delete * from blog_Files where id="&FilesID(i))
            next
            session(CookieName&"_ShowMsg") = True
            session(CookieName&"_MsgText") = session(CookieName&"_MsgText")&(ubound(FilesID)+1)&" 个附件被删除!"
            Response.Redirect("ConContent.asp?Fmenu=SQLFile&Smenu=Attachment")
           Else
             FilesID = split(Request.form("FilesID"),", ")
             FilesPath = split(Request.form("url"),", ")
             FilesCounts = split(Request.form("count"),", ")
             For i = 0 to ubound(FilesID)
             If Int(FilesID(i)) <> -1 then
			    If FilesCounts(i) = "" or FilesCounts(i) < 0 then FilesCounts(i) = 0
                Conn.execute("update blog_Files set FilesPath='"&CheckStr(FilesPath(i))&"',FilesCounts='"&FilesCounts(i)&"' where id="&FilesID(i))
		     Else
				If len(trim(CheckStr(FilesPath(i))))>0 then
				  If FilesCounts(i) = "" or FilesCounts(i) < 0 then FilesCounts(i) = 0
		          Conn.execute("insert into blog_Files (FilesPath,FilesCounts) values ('"&CheckStr(FilesPath(i))&"',"&FilesCounts(i)&")")
		          session(CookieName&"_MsgText")="新附件添加成功! "
		        End If
			 End If
            Next
            session(CookieName&"_ShowMsg") = True
            session(CookieName&"_MsgText") = session(CookieName&"_MsgText")&"附件保存成功!"
            Response.Redirect("ConContent.asp?Fmenu=SQLFile&Smenu=Attachment")
        End If
ElseIf Request.Form("action") = "UpLoadSet" Then
       Dim Antimdown, FileNameMid, FileNameType, FileNameLeft, FileNameRight, Place, Calculate
       Dim SYX, SYY, SYPenColor, SYPenWidth, SYPaddingH, SYPaddingV, SYAlpha, PicPath, PicWidth, PicHeight
       Dim Character, FontColor, FontSize, FontFamily, FontBold, FontItalic, FontShadowColor, FontShadowXOffset, FontShadowYOffset
	   Dim UpLoadSet
	   Antimdown = Request.Form("Antimdown")
       FileNameMid = Request.Form("FileNameMid")
       FileNameType = Request.Form("FileNameType")
       FileNameLeft = Request.Form("FileNameLeft")
       FileNameRight = Request.Form("FileNameRight")
       Place = Request.Form("Place")
       Calculate = Request.Form("Calculate")
       SYX = Request.Form("SYX")
       SYY = Request.Form("SYY")
       SYPenColor = Request.Form("SYPenColor")
       SYPenWidth = Request.Form("SYPenWidth")
       SYPaddingH = Request.Form("SYPaddingH")
       SYPaddingV = Request.Form("SYPaddingV")
       SYAlpha = Request.Form("SYAlpha")
       PicPath = Request.Form("PicPath")
       PicWidth = Request.Form("PicWidth")
       PicHeight = Request.Form("PicHeight")
       Character = Request.Form("Character")
       FontColor = Request.Form("FontColor")
       FontSize = Request.Form("FontSize")
       FontFamily = Request.Form("FontFamily")
       FontBold = Request.Form("FontBold")
       FontItalic = Request.Form("FontItalic")
       FontShadowColor = Request.Form("FontShadowColor")
       FontShadowXOffset = Request.Form("FontShadowXOffset")
       FontShadowYOffset= Request.Form("FontShadowYOffset")
	   If Antimdown <> 1 then Antimdown = 0
	   If Calculate <> 1 then Calculate = 0
	   If FontBold <> 1 then FontBold = 0
	   If FontItalic <> 1 then FontItalic = 0
	   UpLoadSet = Antimdown & "|" & FileNameMid & "|" & FileNameType & "|" & FileNameLeft & "|" & FileNameRight & "|" & Place &_
	   "|" & Calculate & "|" & SYX & "|" & SYY & "|" & SYPenColor & "|" & SYPenWidth & "|" & SYPaddingH &_
	   "|" & SYPaddingV & "|" & SYAlpha & "|" & PicPath & "|" & PicWidth & "|" & PicHeight & "|" & Character &_
	   "|" & FontColor & "|" & FontSize & "|" & FontFamily & "|" & FontBold & "|" & FontItalic & "|" & FontShadowColor &_
	   "|" & FontShadowXOffset & "|" & FontShadowYOffset
       conn.Execute("update blog_Info set blog_UpLoadSet='"&UpLoadSet&"'")
	   getInfo(2)
	   session(CookieName&"_ShowMsg") = True
	   session(CookieName&"_MsgText") = session(CookieName&"_MsgText")&"水印设置保存成功!"
	   RedirectUrl("ConContent.asp?Fmenu=General&Smenu=UpLoadSet")
ElseIf Request.Form("action") = "codeEditor" Then '在线编辑器
    If Request.Form("whatdo") = "save" Then
    	dim referer,fPath,cCode,saveCode
    	referer = Request.Form("referer")
    	fPath = Request.Form("filePath")
    	cCode = Request.Form("code")

		saveCode = SaveToFile(cCode,fPath)

        session(CookieName&"_ShowMsg") = True
        session(CookieName&"_MsgText") = saveCode(1)
        RedirectUrl("ConContent.asp?Fmenu=CodeEditor&referer="&referer&"&file=" & fPath)
    end if
Else'登录欢迎

End If
End Sub


%>
