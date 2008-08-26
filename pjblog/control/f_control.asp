<%
'==================================
'  后台通用函数文件
'    更新时间: 2008-7-13
'==================================

'----------- 显示操作信息 ----------------------------
Sub getMsg
    If session(CookieName&"_ShowMsg") = True Then
        response.Write ("<div id=""msgInfo"" align=""center""><img src=""images/Control/aL.gif"" style=""margin-bottom:-11px;""/><span class=""alertTxt"">" & session(CookieName&"_MsgText") & "</span><img src=""images/Control/aR.gif"" style=""margin-bottom:-11px;""/></div>")
        response.Write ("<script>setTimeout('hiddenMsg()',3000);function hiddenMsg(){document.getElementById('msgInfo').style.display='none';}</script>")
        session(CookieName&"_ShowMsg") = False
        session(CookieName&"_MsgText") = ""
    End If
End Sub

'----------- 退出后台程序 ----------------------------
Sub c_Logout
    session(CookieName&"_System") = ""
    session(CookieName&"_disLink") = ""
    session(CookieName&"_disCount") = ""
    Response.Write ("<script>try{top.location=""default.asp""}catch(e){location=""default.asp""}</script>")
End Sub

'----------- 获得路径的文件信息 ----------------------------

Function getPathList(pathName)
    Dim FSO, ServerFolder, getInfo, getInfos, tempS
    getInfo = ""
    Set FSO = Server.CreateObject("Scripting.FileSystemObject")

    Set ServerFolder = FSO.GetFolder(Server.MapPath(pathName))
    Dim ServerFolderList, ServerFolderEvery
    Set ServerFolderList = ServerFolder.SubFolders
    tempS = ""
    For Each ServerFolderEvery IN ServerFolderList
        getInfo = getInfo&tempS&ServerFolderEvery.Name
        tempS = "*"
    Next
    getInfo = getInfo&"|"
    Dim ServerFileList, ServerFileEvery
    Set ServerFileList = ServerFolder.Files
    tempS = ""
    For Each ServerFileEvery IN ServerFileList
        getInfo = getInfo&tempS&ServerFileEvery.Name
        tempS = "*"
    Next
    Set FSO = Nothing
    getInfos = Split(getInfo, "|")
    getPathList = getInfos
End Function

'----------- 获取文件信息 ----------------------------

Function getFileInfo(FileName)
    Dim FSO, File, FileInfo(8)
    Set FSO = Server.CreateObject("Scripting.FileSystemObject")
    If FSO.FileExists(Server.MapPath(FileName)) Then
        Set File = FSO.GetFile(Server.MapPath(FileName))
        FileInfo(0) = File.Size
        
        If FileInfo(0) / 1000>1 Then
            FileInfo(0) = Int(FileInfo(0) / 1000)&" KB"
        Else
            FileInfo(0) = FileInfo(0)&" Bytes"
        End If
        
        FileInfo(1) = LCase(Right(FileName, 4))
        FileInfo(2) = File.DateCreated
        FileInfo(3) = File.Type
        FileInfo(4) = File.DateLastModified
        FileInfo(5) = File.Path
        FileInfo(6) = "" 'File.ShortPath 部分服务器不支持
        FileInfo(7) = File.Name
        FileInfo(8) = "" 'File.ShortName 部分服务器不支持
    End If
    getFileInfo = FileInfo
    Set FSO = Nothing
End Function

'----------- 获取文件图标 ----------------------------

Function getFileIcons(Str)
    Dim FileIcon, Target
    Select Case Str
        Case ".jpg"
            FileIcon = "jpg.gif"
        Case ".gif"
            FileIcon = "gif.gif"
        Case ".bmp"
            FileIcon = "bmp.gif"
        Case ".png"
            FileIcon = "png.gif"
        Case ".zip"
            FileIcon = "zip.gif"
        Case ".rar"
            FileIcon = "rar.gif"
        Case ".swf"
            FileIcon = "swf.gif"
        Case ".mdb"
            FileIcon = "mdb.gif"
        Case ".doc"
            FileIcon = "doc.gif"
        Case ".xls"
            FileIcon = "xls.gif"
        Case ".pdf"
            FileIcon = "pdf.gif"
        Case ".mbk"
            FileIcon = "mbk.gif"
        Case ".mp3"
            FileIcon = "mp3.gif"
        Case ".wmv"
            FileIcon = "wma.gif"
        Case ".wma"
            FileIcon = "wma.gif"
        Case ".js"
            FileIcon = "js.gif"
        Case ".html"
            FileIcon = "html.gif"
        Case ".htm"
            FileIcon = "htm.gif"
        Case ".css"
            FileIcon = "css.gif"
        Case ".asp"
            FileIcon = "asp.gif"
        Case ".xml"
            FileIcon = "xml.gif"
        Case Else
            FileIcon = "unknow.gif"
    End Select
    getFileIcons = "<img border=""0"" src=""images/file/"&FileIcon&""" style=""margin:4px 3px -3px 0px""/>"
End Function


'----------- 获得目标大小 ----------------------------

Function GetTotalSize(GetLocal, GetType)
    On Error Resume Next
    Dim FSO
    Set FSO = Server.CreateObject("Scripting.FileSystemObject")
    If Err<>0 Then
        Err.Clear
        GetTotalSize = "Fail"
    Else
        Dim SiteFolder
        If GetType = "Folder" Then
            Set SiteFolder = FSO.GetFolder(GetLocal)
        Else
            Set SiteFolder = FSO.GetFile(GetLocal)
        End If
        GetTotalSize = SiteFolder.Size
        If GetTotalSize>1024 * 1024 Then
            GetTotalSize = GetTotalSize / 1024 / 1024
            If InStr(GetTotalSize, ".") Then GetTotalSize = Left(GetTotalSize, InStr(GetTotalSize, ".") + 2)
            GetTotalSize = GetTotalSize&" MB"
        Else
            GetTotalSize = Fix(GetTotalSize / 1024)&" KB"
        End If

        Set SiteFolder = Nothing
    End If
    Set FSO = Nothing
End Function

Function CopyFiles(TempSource, TempEnd) '复制文件
    Dim CopyFSO
    Set CopyFSO = Server.CreateObject("Scripting.FileSystemObject")
    If CopyFSO.FileExists(TempEnd) Then
        CopyFiles = "目标备份文件 <b>" & TempEnd & "</b> 已存在，请先删除!"
        Set CopyFSO = Nothing
        Exit Function
    End If
    If CopyFSO.FileExists(TempSource) Then
    Else
        CopyFiles = "要复制的源数据库文件 <b>"&TempSource&"</b> 不存在!"
        Set CopyFSO = Nothing
        Exit Function
    End If
    CopyFSO.CopyFile TempSource, TempEnd
    CopyFiles = "已经成功复制文件 <b>"&TempSource&"</b> 到 <b>"&TempEnd&"</b>"
    Set CopyFSO = Nothing
End Function

Function DeleteFiles(FilePath) '删除文件
    On Error Resume Next
    Dim FSO
    Set FSO = Server.CreateObject("Scripting.FileSystemObject")
    If FSO.FileExists(FilePath) Then
        FSO.DeleteFile FilePath, True
        If Err Then
            Set FSO = Nothing
            DeleteFiles = False
            Err.Clear
            Exit Function
        End If
        DeleteFiles = True
    Else
        DeleteFiles = False
    End If
    Set FSO = Nothing
End Function

Function FileExist(FilePath) '判断是否存在文件
    FileExist = False
    Dim FSO
    Set FSO = Server.CreateObject("Scripting.FileSystemObject")
    FilePath = Server.MapPath(FilePath)
    If FSO.FileExists(FilePath) Then FileExist = True
End Function

Function DisI(b)
    If b Then
        response.Write "<span style=""color:#00cc00""><b>支持</b></span>"
    Else
        response.Write "<span style=""color:#FF0000""><b>不支持</b></span>"
    End If
End Function

Function getModName '获得模块名字
    Dim Bmodules, Bmodule, tempB
    tempB = ""
    Set blog_module = Conn.Execute("SELECT * FROM blog_module")
    Do Until blog_module.EOF
        Bmodules = Bmodules&tempB&blog_module("name")
        tempB = "|&|"
        blog_module.movenext
    Loop
    Bmodule = Split(Bmodules, "|&|")
    Set blog_module = Nothing
    getModName = Bmodule
End Function


'----------- 释放网站缓存 ----------------------------

Function FreeApplicationMemory
    On Error Resume Next
    Response.Write "释放网站缓存数据列表：<div style='padding:5px 5px 5px 10px;'>"
    Dim Thing,i
    i=0
    For Each Thing IN Application.Contents
    	
        If Left(Thing, Len(CookieName)) = CookieName Then
       		i=i+1
        	if i<30 then
            	Response.Write "<span style='color:#666'>" & thing & "</span><br/>"
            elseif i<31 then
            	Response.Write "<span style='color:#666'>...</span><br/>"
            end if
            
            If IsObject(Application.Contents(Thing)) Then
                Application.Contents(Thing).Close
                Set Application.Contents(Thing) = Nothing
                Application.Contents.Remove(Thing)
            ElseIf IsArray(Application.Contents(Thing)) Then
                Set Application.Contents(Thing) = Nothing
                Application.Contents.Remove(Thing)
            Else
                Application.Contents.Remove(Thing)
            End If
        End If
        
    Next
    response.Write "<br/><span style='color:#666'>共清理了 " & i & " 个缓存数据</span><br/>"
    response.Write "</div>"
End Function

'----------- 更新缓存信息 ----------------------------

Function FreeMemory
    Call reloadcache
    If blog_postFile > 0 Then
        Dim lArticle
        Set lArticle = New ArticleCache
        lArticle.SaveCache
        Set lArticle = Nothing
    End If
End Function

Function InstallPlugingSetting(Path, ModName, action) '安装插件配置文件
    InstallPlugingSetting = -1
    Dim SettingXML, KeyLen, tempI, ModSetting, KeyName, KeyValue
    If action = "delete" Then
        conn.Execute "delete * from blog_ModSetting where set_ModName='"&ModName&"'"
        InstallPlugingSetting = 0
        Exit Function
    End If

    If action = "insert" Then
        Dim modSetC, tmpC
        tmpC = conn.Execute("select count(set_id) as sC from [blog_ModSetting] where set_ModName='"&ModName&"'")(0)
        If tmpC>0 Then
            InstallPlugingSetting = 0
            Exit Function
        End If
        Set SettingXML = New PXML
        SettingXML.XmlPath = Path
        SettingXML.Open
        If SettingXML.getError = 0 Then
            KeyLen = Int(SettingXML.GetXmlNodeLength("PluginOptions/Key"))
            For tempI = 0 To KeyLen -1
                KeyName = SettingXML.GetAttributes("PluginOptions/Key", "name", tempI)
                KeyValue = SettingXML.GetAttributes("PluginOptions/Key", "value", tempI)
                ModSetting = Array(Array("set_ModName", ModName), Array("set_KeyName", KeyName), Array("set_KeyValue", KeyValue))
                DBQuest "blog_ModSetting", ModSetting, "insert"
            Next
        End If
    End If
    InstallPlugingSetting = 0
End Function

Sub delPlugins(PluginsName) '删除插件
    conn.Execute("delete * from blog_module where name='"&PluginsName&"'")
End Sub

Sub delModule(ModID) '删除模块
    conn.Execute("DELETE * from blog_module where id="&ModID)
End Sub

Sub SavehtmlCode(HCode, modID) '写HTML代码到数据库
    conn.Execute("update blog_module set HtmlCode='"& HCode & "' where id=" & modID)
End Sub

Sub PostLink()
    Dim LoadTemplate, Temp, SaveArticle
    LoadTemplate = LoadFromFile("Template/Link.asp")
    If LoadTemplate(0) = 0 Then '读取成功后写入信息
        Temp = LoadTemplate(1)
        Dim blog_Links, ImgLink, TextLink
        Set blog_Links = conn.Execute("select * from blog_Links where link_IsShow=true order by link_Order asc")
        Do Until blog_Links.EOF
            If Len(blog_Links("link_Image"))>0 Then
                ImgLink = ImgLink&"<a href="""&blog_Links("link_URL")&""" target=""_blank""><img src="""&blog_Links("link_Image")&""" alt="""&blog_Links("link_Name")&""" border=""0"" style=""margin:3px;width:88px;height:31px""/></a>"
            Else
                TextLink = TextLink&"<div class=""link"" style=""width:108px;float:left;overflow:hidden;margin-right:8px;height:24px;line-height:180%""><a href="""&blog_Links("link_URL")&""" target=""_blank"" title="""&blog_Links("link_Name")&""">"&blog_Links("link_Name")&"</a></div>"
            End If
            blog_Links.movenext
        Loop
        Temp = Replace(Temp, "<$ImgLink$>", ImgLink)
        Temp = Replace(Temp, "<$TextLink$>", TextLink)
        SaveArticle = SaveToFile(Temp, "post/link.html")
    End If
End Sub


'=========================================
'安装插件
'=========================================

Sub InstallPlugins
    On Error Resume Next
    Dim PlugName, PlugTitle, PlugType, PlugSortID, PlugHtmlCode, PlugPluginPath, PlugSettingXML, MenuID, ConfigPath
    Dim MainItem, SubItem, CreateTableSQL, UpdateTableSQL
    Dim tmpC
    MenuID = -1
    ConfigPath = ""
    Bmodule = getModName
    PluginsFolder = checkstr(Request.QueryString("Plugins"))
    Set PluginsXML = New PXML
    PluginsXML.XmlPath = "Plugins/"&PluginsFolder&"/install.xml"
    PluginsXML.Open
    If PluginsXML.getError = 0 Then
        If Len(PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginName"))>0 And IsvalidPlugins(PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginName")) And (Not IsvalidValue(Bmodule, PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginName"))) And IsvalidValue(TypeArray, PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginType")) Then
            PlugName = PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginName")
            PlugTitle = PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginTitle")
            PlugType = PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginType")
            PlugSortID = Int(conn.Execute("select count(*) from blog_module")(0)) * -1
            PlugHtmlCode = PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginHtmlCode")
            PlugSettingXML = PluginsXML.SelectXmlNodeText("PluginInstall/main/SettingFile")
            ConfigPath = PluginsXML.SelectXmlNodeText("PluginInstall/main/ConfigPath")
            If Err Then
                ConfigPath = ""
                Err.Clear
            End If

            PlugPluginPath = ""
            If PlugType = "function" Then '判断模块插件
                tempI = conn.Execute("select count(*) from blog_Category")(0)
                PlugPluginPath = PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginPath")
                If Len(PlugPluginPath)>0 Then '插入插件菜单
                    MainItem = Array(Array("cate_Name", PlugTitle), Array("cate_Intro", PlugTitle), Array("cate_OutLink", True), Array("cate_URL", "LoadMod.asp?plugins="&PlugName), Array("cate_icon", "images/icons/1.gif"), Array("cate_Order", tempI))
                    DBQuest "blog_Category", MainItem, "insert"
                    CategoryList(2)
                    MenuID = conn.Execute("select top 1 cate_ID from blog_Category order by cate_ID desc")(0)
                End If
            End If

            '写入模块所需要的SQL语句
            CreateTableSQL = PluginsXML.SelectXmlNodeText("PluginInstall/main/CreateTableSQL")
            UpdateTableSQL = PluginsXML.SelectXmlNodeText("PluginInstall/main/UpdateTableSQL")

            If Len(CreateTableSQL)>0 Then conn.Execute(CreateTableSQL)
            If Len(UpdateTableSQL)>0 Then conn.Execute(UpdateTableSQL)
            If Len(PlugSettingXML)>0 And InstallPlugingSetting("Plugins/"&PluginsFolder&"/"&PlugSettingXML, PlugName, "insert") = 0 Then '检测并安装插件配置文件
                MainItem = Array(Array("name", PlugName), Array("title", PlugTitle), Array("type", PlugType), Array("SortID", PlugSortID), Array("HtmlCode", PlugHtmlCode), Array("PluginPath", PlugPluginPath), Array("IsInstall", -1), Array("InstallFolder", PluginsFolder), Array("ConfigPath", ConfigPath), Array("SettingXML", PlugSettingXML), Array("CateID", MenuID))
                Dim ModSetTemp
                Set ModSetTemp = New ModSet
                ModSetTemp.Open PlugName
                ModSetTemp.ReLoad()
            Else
                MainItem = Array(Array("name", PlugName), Array("title", PlugTitle), Array("type", PlugType), Array("SortID", PlugSortID), Array("HtmlCode", PlugHtmlCode), Array("PluginPath", PlugPluginPath), Array("IsInstall", -1), Array("InstallFolder", PluginsFolder), Array("ConfigPath", ConfigPath), Array("CateID", MenuID))
            End If

            DBQuest "blog_module", MainItem, "insert" '插入插件信息
            SubItemLen = Int(PluginsXML.GetXmlNodeLength("PluginInstall/SubItem/item"))

            For tempI = 0 To SubItemLen -1 '安装插件辅助信息
                PlugType = PluginsXML.SelectXmlNodeItemText("PluginInstall/SubItem/item/PluginType", tempI)
                If Not PlugType = "function" Then
                    PlugTitle = PluginsXML.SelectXmlNodeItemText("PluginInstall/SubItem/item/PluginTitle", tempI)
                    PlugSortID = Int(conn.Execute("select count(*) from blog_module")(0)) * -1
                    PlugHtmlCode = PluginsXML.SelectXmlNodeItemText("PluginInstall/SubItem/item/PluginHtmlCode", tempI)
                    SubItem = Array(Array("name", PlugName&"SubItem"&(tempI + 1)), Array("title", PlugTitle), Array("type", PlugType), Array("SortID", PlugSortID), Array("HtmlCode", PlugHtmlCode))
                    DBQuest "blog_module", SubItem, "insert"
                End If
            Next
        Else
            session(CookieName&"_ShowMsg") = True
            Session(CookieName&"_MsgText") = "<font color=""#ff0000"">安装失败</font> 安装插件和其他插件造成冲突，或插件已经安装!"
            RedirectUrl("ConContent.asp?Fmenu=Skins&Smenu=PluginsInstall")
        End If
        PluginsXML.CloseXml()
        log_module(2)
        FixPlugins(0)
        session(CookieName&"_ShowMsg") = True
        Session(CookieName&"_MsgText") = "<font color=""#ff0000"">"&PlugName&"</font> 插件安装完成并且移动到 <a href=""ConContent.asp?Fmenu=Skins&Smenu=Plugins"" style=""color:#0000ff"">已装插件管理</a>"
        RedirectUrl("ConContent.asp?Fmenu=Skins&Smenu=PluginsInstall")
    Else
        session(CookieName&"_ShowMsg") = True
        session(CookieName&"_MsgText") = "安装插件时发生错误,错误代码: <font color=""#ff0000"">"&PluginsXML.getError&"</font>，提示：请检查插件名称显示为红色的插件目录是否存在"
        PluginsXML.CloseXml()
        RedirectUrl("ConContent.asp?Fmenu=Skins&Smenu=PluginsInstall")
    End If
End Sub

'====================写入插件ASP代码======================

Sub FixPlugins(fixType)
    Dim PlugName, PlugTitle, PlugType, PlugSortID, PlugHtmlCode, PlugPluginPath, PlugASPCode, PlugSettingXML
    Dim MainItem, SubItem, CreateTableSQL, UpdateTableSQL, Temp, ASPCode
    ASPCode = ""
    Set Bmodule = conn.Execute("select InstallFolder from blog_module where IsInstall=true")
    If Bmodule.EOF Or Bmodule.bof Then
        If CheckObjInstalled("ADODB.Stream") Then SaveToFile "", "plugins.asp"
    Else
        Do Until Bmodule.EOF
            PluginsFolder = Bmodule("InstallFolder")
            Set PluginsXML = New PXML
            PluginsXML.XmlPath = "Plugins/"&PluginsFolder&"/install.xml"
            PluginsXML.Open
            If PluginsXML.getError = 0 Then
                PlugName = PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginName")
                PlugASPCode = PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginASPCode")
                Temp = "&lt;%'---- ASPCode For "&PlugName&" ----%&gt;"
                Temp = Replace(Temp, "&lt;", "<")
                Temp = Replace(Temp, "&gt;", ">")
                ASPCode = ASPCode & Temp & Chr(10)
                If Len(PlugASPCode)>0 Then
                    ASPCode = ASPCode & PlugASPCode & Chr(10)
                End If
                SubItemLen = Int(PluginsXML.GetXmlNodeLength("PluginInstall/SubItem/item"))
                For tempI = 0 To SubItemLen -1
                    PlugType = PluginsXML.SelectXmlNodeItemText("PluginInstall/SubItem/item/PluginType", tempI)
                    If Not PlugType = "function" Then
                        PlugASPCode = PluginsXML.SelectXmlNodeItemText("PluginInstall/SubItem/item/PluginASPCode", tempI)
                        Temp = "&lt;%'---- ASPCode For "&PlugName&"SubItem"&(tempI + 1)&" ----%&gt;"
                        Temp = Replace(Temp, "&lt;", "<")
                        Temp = Replace(Temp, "&gt;", ">")
                        ASPCode = ASPCode & Temp & Chr(10)
                        If Len(PlugASPCode)>0 Then
                            ASPCode = ASPCode & PlugASPCode & Chr(10)
                        End If
                    End If
                Next
                PluginsXML.CloseXml()

                log_module(2)
                If CheckObjInstalled("ADODB.Stream") Then SaveToFile ASPCode, "plugins.asp"
            Else
                If CBool(fixType) Then
                    session(CookieName&"_ShowMsg") = True
					session(CookieName&"_MsgText") = "安装插件时发生错误,错误代码: <font color=""#ff0000"">"&PluginsXML.getError&"</font>，提示：请检查插件名称显示为红色的插件目录是否存在"
                    PluginsXML.CloseXml()
                    RedirectUrl("ConContent.asp?Fmenu=Skins&Smenu=Plugins")
                End If
            End If
            Bmodule.movenext
        Loop
    End If
    If CBool(fixType) Then
        session(CookieName&"_ShowMsg") = True
        Session(CookieName&"_MsgText") = "插件已经修复完成!"
        RedirectUrl("ConContent.asp?Fmenu=Skins&Smenu=Plugins")
    End If
End Sub

'====================删除日志方法======================

Function DeleteLog(LogID)
    Dim logDB, comCount, post_CateID
    Set logDB = conn.Execute("select log_IsDraft,log_CateID from blog_content where log_ID="&LogID)
    If logDB.EOF And logDB.bof Then
        DeleteLog = 0
    Else
        comCount = Conn.Execute("select count(blog_ID) FROM blog_Comment WHERE blog_ID="&LogID)(0)
        response.Write comCount&"<br/>"
        Conn.Execute("update blog_Info set blog_CommNums=blog_CommNums-"&comCount)
        If logDB("log_IsDraft") = False Then
            Conn.Execute("UPDATE blog_Info SET blog_LogNums=blog_LogNums-1")
            Conn.Execute("UPDATE blog_Category SET cate_count=cate_count-1 where cate_ID="&logDB("log_CateID"))
        End If
        Conn.Execute("DELETE * FROM blog_Comment WHERE blog_ID="&LogID)
        Conn.Execute("DELETE * FROM blog_Content WHERE log_ID="&LogID)
        DeleteFiles Server.MapPath("post/"&LogID&".asp")
        DeleteFiles Server.MapPath("article/"&LogID&".htm")
        DeleteLog = 1
    End If
End Function


Function bc(t, s)
    Dim tl, sl, i
    bc = False
    sl = Len(s)
    tl = Len(t)
    If tl< sl Then
        bc = True
        Exit Function
    End If
    For i = 1 To sl
        If Mid(t, i, 1)<>Mid(s, i, 1) Then
            bc = True
            Exit Function
        End If
    Next
End Function

Sub saveFilterKey(keyword)
    Dim tempStr, saveXML, keywords, i
    tempStr = "<?xml version=""1.0"" ?><spam>"
    keywords = Split(keyword, ", ")
    For i = 0 To UBound(keywords)
        If Len(Trim(keywords(i)))>0 Then tempStr = tempStr & "<key><![CDATA[" & Trim(keywords(i)) & "]]></key>"
    Next
    tempStr = tempStr & "</spam>"
    saveXML = SaveToFile (tempStr, "spam.xml")
    session(CookieName&"_ShowMsg") = True
    If saveXML(0)<>0 Then session(CookieName&"_MsgText") = saveXML(1) Else session(CookieName&"_MsgText") = "过滤列表保存成功."
    RedirectUrl("ConContent.asp?Fmenu=Comment&Smenu=spam")
End Sub

Sub saveReFilterKey()
    Dim tempStr, saveXML, keywords, i
    tempStr = "<?xml version=""1.0"" ?><reg>"
    tempStr = tempStr & "<text><![CDATA[" & Request.Form("testArea") & "]]></text>"

    For i = 1 To Request.Form("des").Count
        tempStr = tempStr & "<key des=""" & htmlE(Request.Form("des").Item(i)) & """ re=""" & htmlE(Request.Form("re").Item(i)) & """ times=""" & Request.Form("times").Item(i) & """/>"
    Next

    tempStr = tempStr & "</reg>"

    saveXML = SaveToFile (tempStr, "reg.xml")

    session(CookieName&"_ShowMsg") = True
    If saveXML(0)<>0 Then session(CookieName&"_MsgText") = saveXML(1) Else session(CookieName&"_MsgText") = "过滤列表保存成功."
    RedirectUrl("ConContent.asp?Fmenu=Comment&Smenu=reg")
End Sub

Function htmlE(reString)
    Dim Str
    Str = reString
    If Not IsNull(Str) Then
        Str = Replace(Str, "&", "&amp;")
        Str = Replace(Str, ">", "&gt;")
        Str = Replace(Str, "<", "&lt;")
        Str = Replace(Str, """", "&quot;")
        Str = Replace(Str, "\", "\\")
        htmlE = Str
    End If
End Function

Function delskin(FolderName)
    Dim upl
    Set upl=Server.CreateObject("Scripting.FileSystemObject")
    If upl.FolderExists(Server.MapPath("skins/"&folderName)) Then
        upl.deleteFolder Server.MapPath("skins/"&folderName)
    End If
    Set upl = Nothing
    session(CookieName&"_ShowMsg") = True
    Session(CookieName&"_MsgText") = "已经删除了"&FolderName
    RedirectUrl("ConContent.asp?Fmenu=Skins&Smenu=")
end function

'-----------获取分类标题-------------------

Function categoryTitle()
    Dim Fmenu, Smenu, cTitle
    Fmenu = Request.QueryString("Fmenu")
    Smenu = Request.QueryString("Smenu")
    Set cTitle = Server.CreateObject("Scripting.Dictionary")
    cTitle.Add "General.Misc" , "站点基本设置 - 初始化数据"
    cTitle.Add "General.clear" , "站点基本设置 - 清理服务器缓存"
    cTitle.Add "General.visitors" , "站点基本设置 - 查看访客记录"
    cTitle.Add "General." , "站点基本设置 - 设置基本信息"

    cTitle.Add "Categories.move" , "日志分类管理 - 批量移动日志"
    cTitle.Add "Categories.tag" , "日志分类管理 - Tag管理"
    cTitle.Add "Categories.del" , "日志分类管理 - 批量删除日志"
    cTitle.Add "Categories." , "日志分类管理 - 设置日志分类"

    cTitle.Add "Comment.spam" , "评论留言管理 - 初级过滤设置 (垃圾关键字过滤黑名单)"
    cTitle.Add "Comment.reg" , "评论留言管理 - 高级过滤设置 (利用正则表达式过滤)"
    cTitle.Add "Comment.trackback" , "评论留言管理 - 引用管理"
    cTitle.Add "Comment.msg" , "评论留言管理 - 留言管理"
    cTitle.Add "Comment." , "评论留言管理 - 评论管理"

    cTitle.Add "Skins.module" , "界面设置 - 设置模块"
    cTitle.Add "Skins.Plugins" , "界面设置 - 已装插件管理"
    cTitle.Add "Skins.PluginsInstall" , "界面设置 - 安装插件"
    cTitle.Add "Skins.editModule" , "界面设置 - 可视化编辑HTML代码"
    cTitle.Add "Skins.editModuleNormal" , "界面设置 - 编辑HTML源代码"
    cTitle.Add "Skins.PluginsOptions" , "界面设置 - 插件配置"
    cTitle.Add "Skins." , "界面设置 - 设置外观"

    cTitle.Add "SQLFile.Attachments" , "数据库与附件 - 附件管理"
    cTitle.Add "SQLFile." , "数据库与附件 - 数据库管理"

    cTitle.Add "Members.Users" , "帐户与权限管理 - 帐户管理"
    cTitle.Add "Members.EditRight" , "帐户与权限管理 - 编辑权限细节"
    cTitle.Add "Members." , "帐户与权限管理 - 权限管理"

    cTitle.Add "Link." , "友情链接管理"

    cTitle.Add "smilies.KeyWord" , "表情与关键字 - 关键字管理"
    cTitle.Add "smilies." , "表情与关键字 - 表情管理"

    cTitle.Add "Status." , "服务器配置信息 - PJBlog3"

    cTitle.Add "welcome." , "欢迎使用PJBlog3"
    
    cTitle.Add "CodeEditor." , "在线代码编辑器 beta"
    
    categoryTitle = cTitle(Fmenu & "." & Smenu)
    Set cTitle = Nothing
End Function
%>
