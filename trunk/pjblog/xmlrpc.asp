<!--#include file="const.asp" -->
<!--#include file="conn.asp" -->
<!--#include file="common/cache.asp" -->
<!--#include file="common/function.asp" -->
<!--#include file="common/library.asp" -->
<!--#include file="common/ubbcode.asp" -->
<!--#include file="common/checkUser.asp" -->
<!--#include file="common/ModSet.asp" -->
<!--#include file="class/cls_logAction.asp" -->
<!--#include file="class/cls_article.asp" -->
<!--#include file="common/xml.asp" -->
<%
'======================================
'  XML-RPC for PJBlog
'======================================

'读取Blog设置信息
getInfo(1)

'写入关键字列表
Keywords(1)

'写入表情符号
Smilies(1)

'写入标签
Tags(1)

Response.Charset = "UTF-8"
Response.ContentType = "text/xml"
Response.Expires = 60
Response.Buffer = True

Dim DebugOn
DebugOn = False

XmlBin = Request.BinaryRead(Request.TotalBytes)
SaveToFile bin2str(XmlBin),"debug\out_"&randomStr(3)&"_"&Year(now)&Month(now)&Day(now)&Hour(now)&Minute(now)&Second(now)&".txt"
Dim xmlPrc, XmlBin
Set xmlPrc = New PXML

If DebugOn Then
    xmlPrc.xmlPath = "out.xml"
    xmlPrc.Open()
Else
    xmlPrc.OpenXML(XmlBin)
End If

If xmlPrc.getError = 0 Then
    Dim strAction
    Dim userName, passWord, intPosts, bolPublish, logTitle, logDescription, logPostTime, logCategories, tagWords, logCategoryId, logID, fileBits, fileName,logIntro
    strAction = xmlPrc.SelectXmlNodeText("methodName")
    Select Case strAction
Case "blogger.getUsersBlogs":
        userName = xmlPrc.SelectXmlNodeText("params/param[1]/value")
        passWord = xmlPrc.SelectXmlNodeText("params/param[2]/value")
        login2 userName, passWord
        If stat_Admin Then Call getUsersBlogs() Else Call returnError(0, "permitted to get Users Blogs.")
Case "metaWeblog.getCategories":
        userName = xmlPrc.SelectXmlNodeText("params/param[1]/value")
        passWord = xmlPrc.SelectXmlNodeText("params/param[2]/value")
        login2 userName, passWord
        If stat_Admin Then Call getCategories() Else Call returnError(0, "permitted to get Categories.")
Case "mt.getCategoryList":
        userName = xmlPrc.SelectXmlNodeText("params/param[1]/value")
        passWord = xmlPrc.SelectXmlNodeText("params/param[2]/value")
        login2 userName, passWord
        If stat_Admin Then Call getCategories() Else Call returnError(0, "permitted to get Categories.")
Case "mt.getPostCategories": 
	    if xmlPrc.GetXmlNodeLength("params/param[0]/value/string")>0 then 
	      logID=xmlPrc.SelectXmlNodeText("params/param[0]/value/string") 
	    else 
	      logID=xmlPrc.SelectXmlNodeText("params/param[0]/value/struct/member[name=""PostID""]/value") 
	    end if 
	    if len(toInt(logID))<1 then Call returnError(0,"parameter error.") 
	    logID=int(toInt(logID)) 
	    userName=xmlPrc.SelectXmlNodeText("params/param[1]/value") 
	    passWord=xmlPrc.SelectXmlNodeText("params/param[2]/value") 
	     login2 userName,passWord 
	    if stat_Admin then Call getPostCategories(logID) else Call returnError(0,"permitted to get post categories.")     
Case "metaWeblog.getRecentPosts":
        userName = xmlPrc.SelectXmlNodeText("params/param[1]/value")
        passWord = xmlPrc.SelectXmlNodeText("params/param[2]/value")
        intPosts = xmlPrc.SelectXmlNodeText("params/param[3]/value/int")

        If intPosts = 0 Then
            intPosts = xmlPrc.SelectXmlNodeText("params/param[3]/value/i4")
        End If

        login2 userName, passWord
        If stat_Admin Then Call getRecentPosts(intPosts) Else Call returnError(0, "permitted to get Recent Posts.")
Case "metaWeblog.newPost":
        userName = xmlPrc.SelectXmlNodeText("params/param[1]/value")
        passWord = xmlPrc.SelectXmlNodeText("params/param[2]/value")
        logTitle = xmlPrc.SelectXmlNodeText("params/param[3]/value/struct/member[name=""title""]/value")
        logDescription = xmlPrc.SelectXmlNodeText("params/param[3]/value/struct/member[name=""description""]/value")
        logPostTime = xmlPrc.SelectXmlNodeText("params/param[3]/value/struct/member[name=""dateCreated""]/value/dateTime.iso8601")
		'tagWords=xmlPrc.GetXmlNodeLength("params/param[3]/value/struct/member[name=""tagwords""]/value/array/data/value")
        '支持windows live writer的关键字和Zoundry的标签
		tagWords=xmlPrc.SelectXmlNodeText("params/param[3]/value/struct/member[name=""mt_keywords""]/value")
		logIntro=xmlPrc.SelectXmlNodeText("params/param[3]/value/struct/member[name=""mt_excerpt""]/value")
		bolPublish = xmlPrc.SelectXmlNodeText("params/param[4]/value/boolean")
        login2 userName, passWord
        If stat_Admin Then Call newPost(logTitle, logDescription,logIntro,logPostTime, tagWords, bolPublish ) Else Call returnError(0, "permitted to post a new log.")
Case "metaWeblog.editPost":
        If xmlPrc.GetXmlNodeLength("params/param[0]/value/string")>0 Then
            logID = xmlPrc.SelectXmlNodeText("params/param[0]/value/string")
        Else
            logID = xmlPrc.SelectXmlNodeText("params/param[0]/value/struct/member[name=""PostID""]/value")
        End If
        If Len(toInt(logID))<1 Then Call returnError(0, "parameter error.")
        logID = Int(toInt(logID))
        userName = xmlPrc.SelectXmlNodeText("params/param[1]/value")
        passWord = xmlPrc.SelectXmlNodeText("params/param[2]/value")
        logTitle = xmlPrc.SelectXmlNodeText("params/param[3]/value/struct/member[name=""title""]/value")
        logDescription = xmlPrc.SelectXmlNodeText("params/param[3]/value/struct/member[name=""description""]/value")
        logPostTime = xmlPrc.SelectXmlNodeText("params/param[3]/value/struct/member[name=""dateCreated""]/value/dateTime.iso8601")
		tagWords=xmlPrc.SelectXmlNodeText("params/param[3]/value/struct/member[name=""mt_keywords""]/value")
		logIntro=xmlPrc.SelectXmlNodeText("params/param[3]/value/struct/member[name=""mt_excerpt""]/value")
		'tagWords=xmlPrc.GetXmlNodeLength("params/param[3]/value/struct/member[name=""tagwords""]/value/array/data/value")
		bolPublish = xmlPrc.SelectXmlNodeText("params/param[4]/value/boolean")
        login2 userName, passWord
        If stat_Admin Then Call editPost(logID, logTitle, logDescription,logIntro, logPostTime,tagWords, bolPublish) Else Call returnError(0, "permitted to post a new log.")
Case "mt.setPostCategories":
        If xmlPrc.GetXmlNodeLength("params/param[0]/value/string")>0 Then
            logID = xmlPrc.SelectXmlNodeText("params/param[0]/value/string")
        Else
            logID = xmlPrc.SelectXmlNodeText("params/param[0]/value/struct/member[name=""PostID""]/value")
        End If
        If Len(toInt(logID))<1 Then Call returnError(0, "parameter error.")
        logID = Int(toInt(logID))
        userName = xmlPrc.SelectXmlNodeText("params/param[1]/value")
        passWord = xmlPrc.SelectXmlNodeText("params/param[2]/value")
        logCategoryId = xmlPrc.SelectXmlNodeText("params/param[3]/value/array/data/value[0]/struct/member[name=""categoryId""]/value")
        login2 userName, passWord
        If stat_Admin Then Call setPostCategories(logID, logCategoryId) Else Call returnError(0, "permitted to set post categories.")
Case "metaWeblog.getPost":
        If xmlPrc.GetXmlNodeLength("params/param[0]/value/string")>0 Then
            logID = xmlPrc.SelectXmlNodeText("params/param[0]/value/string")
        Else
            logID = xmlPrc.SelectXmlNodeText("params/param[0]/value/struct/member[name=""PostID""]/value")
        End If
        If Len(toInt(logID))<1 Then Call returnError(0, "parameter error.")
        logID = Int(toInt(logID))
        userName = xmlPrc.SelectXmlNodeText("params/param[1]/value")
        passWord = xmlPrc.SelectXmlNodeText("params/param[2]/value")
        login2 userName, passWord
        If stat_Admin Then Call getPost(logID) Else Call returnError(0, "permitted to set post categories.")
Case "metaWeblog.newMediaObject":
        userName = xmlPrc.SelectXmlNodeText("params/param[1]/value")
        passWord = xmlPrc.SelectXmlNodeText("params/param[2]/value")
        fileBits = xmlPrc.SelectXmlNodeText("params/param[3]/value/struct/member[name=""bits""]/value")
        fileName = xmlPrc.SelectXmlNodeText("params/param[3]/value/struct/member[name=""name""]/value")
        login2 userName, passWord
        If stat_Admin Then Call newMediaObject(fileName, fileBits) Else Call returnError(0, "permitted to upload file.")
Case "blogger.deletePost":
        If xmlPrc.GetXmlNodeLength("params/param[1]/value/string")>0 Then
            logID = xmlPrc.SelectXmlNodeText("params/param[1]/value/string")
        Else
            logID = xmlPrc.SelectXmlNodeText("params/param[1]/value/struct/member[name=""PostID""]/value")
        End If
        If Len(toInt(logID))<1 Then Call returnError(0, "parameter error.")
        logID = Int(toInt(logID))
        userName = xmlPrc.SelectXmlNodeText("params/param[2]/value")
        passWord = xmlPrc.SelectXmlNodeText("params/param[3]/value")
        login2 userName, passWord
        If stat_Admin Then Call deletePost(logID) Else Call returnError(0, "permitted to delete log.")
        Case Else
            xmlPrc.CloseXml()
            Call returnError(0, strAction)
    End Select
Else
    xmlPrc.CloseXml()
    Call returnError(0, "action Error2.")
End If

'=================Function In XML-PRC=========================

'----------------------return response Error---------------------------

Function returnError(faultCode, faultString)
    Response.Clear
    Response.Write "<?xml version=""1.0"" encoding=""UTF-8""?><methodResponse>"
    Response.Write "<fault><value><struct>"
    Response.Write "<member><name>faultCode</name><value><int>"&faultCode&"</int></value></member>"
    Response.Write "<member><name>faultString</name><value><string>"&faultString&"</string></value></member>"
    Response.Write "</struct></value></fault>"
    Response.Write "</methodResponse>"
    Response.End
End Function

'----------------------blogger.getUsersBlogs---------------------------

Function getUsersBlogs()
    Response.Clear
    Response.Write "<?xml version=""1.0"" encoding=""UTF-8""?>"
    Response.Write "<methodResponse><params><param><value><array><data><value><struct>"
    Response.Write "<member><name>url</name><value><string>"&CheckStr(SiteURL)&"</string></value></member>"
    Response.Write "<member><name>blogid</name><value><string>"&CheckStr(CookieName)&"</string></value></member>"
    Response.Write "<member><name>blogName</name><value><string>"&CheckStr(UnCheckStr(SiteName))&"</string></value></member>"
    Response.Write "</struct></value></data></array></value></param></params></methodResponse>"
    Response.End
End Function


'----------------------metaWeblog.getRecentPosts---------------------------

Function getRecentPosts(intNum)
    Dim RecentPosts

    Dim RS, dbRow, i

    SQL = "SELECT TOP "&intNum&" L.log_ID,L.log_Title,L.log_Author,L.log_Content,L.log_PostTime,L.log_edittype,C.cate_Name,L.log_IsDraft FROM blog_Content AS L,blog_Category AS C WHERE C.cate_ID=L.log_cateID ORDER BY L.log_PostTime DESC"
    Set RS = Conn.Execute(SQL)
    If RS.EOF Or RS.BOF Then
        ReDim dbRow(0, 0)
    Else
        dbRow = RS.getrows()
    End If
    RS.Close
    Set RS = Nothing
    Call CloseDB

    RecentPosts = "<?xml version=""1.0"" encoding=""UTF-8""?><methodResponse><params><param><value><array><data>"

    If UBound(dbRow, 1)<>0 Then
        For i = 0 To UBound(dbRow, 2)
            RecentPosts = RecentPosts & "<value><struct>"
            RecentPosts = RecentPosts & "<member><name>title</name><value><string>"&dbRow(1, i)&"</string></value></member>"
            If dbRow(5, i) = 1 Then
                t = AddSiteURL(UBBCode(HTMLEncode(dbRow(3, i)), 0, 0, 0, 1, 1))
                RecentPosts = RecentPosts & "<member><name>description</name><value><string>"&CheckStr(t)&"</string></value></member>"
            Else
                t = AddSiteURL(UnCheckStr(dbRow(3, i)))
                RecentPosts = RecentPosts & "<member><name>description</name><value><string>"&CheckStr(t)&"</string></value></member>"
            End If
            RecentPosts = RecentPosts & "<member><name>dateCreated</name><value><dateTime.iso8601>"&DateToStr(dbRow(4, i), "y-m-dTH:I:S")&"</dateTime.iso8601></value></member>"
            RecentPosts = RecentPosts & "<member><name>categories</name><value><array><data><value><string>"&dbRow(6, i)&"</string></value></data></array></value></member>"
            RecentPosts = RecentPosts & "<member><name>tagwords</name><value><string><![CDATA[df34,err]]></string></value></member>"
            RecentPosts = RecentPosts & "<member><name>postid</name><value><string>"&CheckStr(dbRow(0, i))&"</string></value></member>"
            RecentPosts = RecentPosts & "<member><name>userid</name><value><string>"&CheckStr(dbRow(2, i))&"</string></value></member>"
            RecentPosts = RecentPosts & "<member><name>link</name><value><string>"&CheckStr(SiteURL&"/default.asp?id="&dbRow(0, i))&"</string></value></member>"
            RecentPosts = RecentPosts & "<member><name>permaLink</name><value><string>"&CheckStr(SiteURL&"/default.asp?id="&dbRow(0, i))&"</string></value></member>"
            RecentPosts = RecentPosts & "</struct></value>"
        Next
    End If
    RecentPosts = RecentPosts & "</data></array></value></param></params></methodResponse>"

    'SaveToFile RecentPosts,"debug\recent2_"&randomStr(3)&"_"&Year(now)&Month(now)&Day(now)&Hour(now)&Minute(now)&Second(now)&".txt"

    Response.Clear
    Response.Write RecentPosts
End Function

'----------------------metaWeblog.getPost---------------------------

Function getPost(lID)
    Dim RecentPosts,getLog,lArticle
    Set lArticle = New logArticle
    getLog = lArticle.getLog(lID)

    If getLog(0)<0 Then
       Call returnError(0, "Can't find log.")
  	   Set RecentPosts = nothing
  	   Call CloseDB
  	   exit Function
    End If
'    SQL = "SELECT TOP 1 L.log_ID,L.log_Title,L.log_Author,L.log_Content,L.log_PostTime,L.log_edittype,C.cate_Name,L.log_IsDraft FROM blog_Content AS L,blog_Category AS C WHERE C.cate_ID=L.log_cateID And L.log_ID="&lID&" ORDER BY L.log_PostTime DESC"

    RecentPosts = "<?xml version=""1.0"" encoding=""UTF-8""?><methodResponse><params><param>"
    RecentPosts = RecentPosts & "<value><struct>"
    RecentPosts = RecentPosts & "<member><name>title</name><value><string>"&CheckStr(lArticle.logTitle)&"</string></value></member>"
    If lArticle.logEditType = 1 Then
        t = AddSiteURL(UBBCode(HTMLEncode(lArticle.logMessage), 0, 0, 0, 1, 1))
        RecentPosts = RecentPosts & "<member><name>description</name><value><string>"&CheckStr(t)&"</string></value></member>"
    Else
        t = AddSiteURL(UnCheckStr(lArticle.logMessage))
        RecentPosts = RecentPosts & "<member><name>description</name><value><string>"&CheckStr(t)&"</string></value></member>"
    End If
    
    RecentPosts = RecentPosts & "<member><name>dateCreated</name><value><dateTime.iso8601>"&DateToStr(lArticle.logPubTime, "y-m-dTH:I:S")&"</dateTime.iso8601></value></member>"
    RecentPosts = RecentPosts & "<member><name>categories</name><value><array><data><value><string>"&lArticle.categoryID&"</string></value></data></array></value></member>"
    RecentPosts = RecentPosts & "<member><name>mt_keywords</name><value><string>"&CheckStr(lArticle.logTags)&"</string></value></member>"
    if not CBool(lArticle.logIntroCustom) then
    	RecentPosts = RecentPosts & "<member><name>mt_excerpt</name><value><string>"&CheckStr(lArticle.logIntro)&"</string></value></member>"
    else
    	RecentPosts = RecentPosts & "<member><name>mt_excerpt</name><value><string><![CDATA[]]></string></value></member>"
    end if
    RecentPosts = RecentPosts & "<member><name>postid</name><value><string>"&CheckStr(lID)&"</string></value></member>"
    RecentPosts = RecentPosts & "<member><name>userid</name><value><string>"&CheckStr(lArticle.logAuthor)&"</string></value></member>"
    RecentPosts = RecentPosts & "<member><name>link</name><value><string>"&CheckStr(SiteURL&"/default.asp?id="&lID)&"</string></value></member>"
    RecentPosts = RecentPosts & "<member><name>permaLink</name><value><string>"&CheckStr(SiteURL&"/default.asp?id="&lID)&"</string></value></member>"
    RecentPosts = RecentPosts & "</struct></value>"
    RecentPosts = RecentPosts & "</param></params></methodResponse>"
    response.Write RecentPosts
   
    Call CloseDB
End Function

'----------------------mt.getCategoryList/metaWeblog.getCategories---------------------------

Function getCategories()
    Dim Categories
    Categories = "<?xml version=""1.0"" encoding=""UTF-8""?><methodResponse><params><param><value><array><data>"
    Dim Arr_Category, Category_Len, i
    CategoryList(3)
    Arr_Category = Application(CookieName&"_blog_Category")
    If UBound(Arr_Category, 1) = 0 Then Call returnError(0, "no Categories")
    Category_Len = UBound(Arr_Category, 2)

    For i = 0 To Category_Len
        If Not Arr_Category(4, i) Then
            Categories = Categories & "<value><struct>"
            Categories = Categories & "<member><name>description</name><value><string>"&CheckStr(Arr_Category(1, i))&"</string></value></member>"
            Categories = Categories & "<member><name>httpUrl</name><value><string>"&CheckStr(SiteURL&"/default.asp?cateID="&Arr_Category(0, i))&"</string></value></member>"
            Categories = Categories & "<member><name>rssUrl</name><value><string>"&CheckStr(SiteURL&"/feed.asp?cateID="&Arr_Category(0, i))&"</string></value></member>"
            Categories = Categories & "<member><name>title</name><value><string>"&CheckStr(Arr_Category(1, i))&"</string></value></member>"
            Categories = Categories & "<member><name>categoryId</name><value><string>"&CheckStr(Arr_Category(0, i))&"</string></value></member>"
            Categories = Categories & "<member><name>categoryName</name><value><string>"&CheckStr(Arr_Category(1, i))&"</string></value></member>"
            Categories = Categories & "</struct></value>"
        End If
    Next
    Categories = Categories & "</data></array></value></param></params></methodResponse>"
    Response.Write Categories
End Function

'----------------------metaWeblog.newPost---------------------------

Function newPost(lTitle, lDescription,lIntro, lPostTime, lTagwords, lPublish)
    '====get Last Category
    Dim Arr_Category, Category_Len, i, lastCID
    CategoryList(1)
    Arr_Category = Application(CookieName&"_blog_Category")
    If UBound(Arr_Category, 1) = 0 Then Call returnError(0, "no Categories")
    Category_Len = UBound(Arr_Category, 2)
    For i = 0 To Category_Len
        If Not Arr_Category(4, i) Then
            lastCID = Arr_Category(0, i)
            Exit For
        End If
    Next
    Dim newPostStr
    Dim lArticle, postLog
    Set lArticle = New logArticle
    lArticle.categoryID = lastCID
    lArticle.logTitle = lTitle
    lArticle.logEditType = 0
    lArticle.logIsDraft = Not CBool(lPublish)
    lArticle.logMessage = lDescription
    lArticle.logPubTime = Now()
	lArticle.logTags    = lTagwords
    if len(Trim(lIntro))>0 then
    	lArticle.logIntroCustom = 1
    	lArticle.logIntro = lIntro
    else
    	lArticle.logIntroCustom = 0
    end if
    lArticle.logAuthor = memName
    postLog = lArticle.postLog
    Set lArticle = Nothing
    If postLog(2)<0 Then
        Call returnError(0, postLog(1))
    Else
        newPostStr = "<?xml version=""1.0"" encoding=""UTF-8""?><methodResponse><params><param>"
        newPostStr = newPostStr & "<value><string>"&postLog(2)&"</string></value>"
        newPostStr = newPostStr & "</param></params></methodResponse>"
        response.Write newPostStr
    End If
End Function

'----------------------metaWeblog.editPost---------------------------

Function editPost(lID, lTitle, lDescription,lIntro, lPostTime, lTagwords,lPublish)
    '====get Last Category
    Dim editPostStr
    Dim lArticle, editLog
    Set lArticle = New logArticle
    editLog = lArticle.getLog(lID)
    If editLog(0)<0 Then Call returnError(0, "Can't find log.")

    lArticle.logTitle = lTitle
    lArticle.logEditType = 0
    lArticle.logIsDraft = Not CBool(lPublish)
    lArticle.logMessage = lDescription
    lArticle.logPubTime = Now()
    if len(Trim(lIntro))>0 then
    	lArticle.logIntroCustom = 1
    	lArticle.logIntro = lIntro
    else
    	lArticle.logIntroCustom = 0
    end if
	lArticle.logTags    = lTagwords
    editLog = lArticle.editLog(lID)
    Set lArticle = Nothing
    If editLog(2)<0 Then
        Call returnError(0, editLog(1))
    Else
        editPostStr = "<?xml version=""1.0"" encoding=""UTF-8""?><methodResponse><params><param>"
        editPostStr = editPostStr & "<value><struct>"
        editPostStr = editPostStr & "<member><name>PostID</name><value><string>"&editLog(2)&"</string></value></member>"
        editPostStr = editPostStr & "</struct></value>"
        editPostStr = editPostStr & "</param></params></methodResponse>"
        response.Write editPostStr
    End If
End Function



'----------------------mt.setPostCategories---------------------------

Function setPostCategories(lID, lCID)
    Dim returnStr
    Dim lArticle, editLog
    Set lArticle = New logArticle
    editLog = lArticle.getLog(lID)
    Set lArticle = Nothing
    If editLog(0)<0 Then
        Call returnError(0, "Can't find log.")
    Else
        If IsInteger(lCID) Then
            Dim lastCID
            lastCID = conn.Execute("select top 1 log_cateID from blog_Content where log_ID="&lID)(0)
            Conn.Execute("UPDATE blog_Category SET cate_count=cate_count-1 where cate_ID="&lastCID)
            Conn.Execute("UPDATE blog_Category SET cate_count=cate_count+1 where cate_ID="&lCID)
            Conn.Execute("UPDATE blog_Content SET log_cateID="&lCID&" where log_ID="&lID)
        End If

        If blog_postFile>0 Then
            Set lArticle = New ArticleCache
            lArticle.SaveCache
            Set lArticle = Nothing
            PostArticle lID, False
        End If

        returnStr = "<?xml version=""1.0"" encoding=""UTF-8""?><methodResponse><params><param><value><boolean>1</boolean></value></param></params></methodResponse>"
        response.Write returnStr
    End If
End Function

Function deletePost(lID)
    Dim lArticle, DeleteLog, returnStr
    Set lArticle = New logArticle
    DeleteLog = lArticle.deleteLog(lID)
    Set lArticle = Nothing
    If DeleteLog(0) = 0 Then
        returnStr = "<?xml version=""1.0"" encoding=""UTF-8""?><methodResponse><params><param><value><boolean>1</boolean></value></param></params></methodResponse>"
        response.Write returnStr
    Else
        Call returnError(0, "Can't delete log.")
    End If
End Function

'----------------------metaWeblog.newMediaObject"--------------------------

Function newMediaObject(fName, fBits)
    'On Error Resume Next
    If stat_FileUpLoad = False Then Call returnError(0, "permitted to upload file.")
    Dim upl, FSOIsOK
    FSOIsOK = 1
    Set upl = Server.CreateObject("Scripting.FileSystemObject")
    If Err<>0 Then
        Err.Clear
        Set upl = Nothing
        Call returnError(0, "Can't create folder.")
    End If
    Dim D_Name
    D_Name = "month_"&DateToStr(Now(), "ym")
    If upl.FolderExists(Server.MapPath("attachments/"&D_Name)) = False Then
        upl.CreateFolder Server.MapPath("attachments/"&D_Name)
    End If

    Dim FileExt, i
    FileExt = ""
    For i = Len(fName) To 1 step -1
        If Mid(fName, i, 1) = "." Then Exit For
        FileExt = Mid(fName, i, 1) & FileExt
    Next
    Dim tStream, base64M
    Set base64M = New base64
    Set tStream = Server.CreateObject("adodb.stream")
    tStream.Type = 1
    tStream.Mode = 3
    tStream.Open
    tStream.Position = 0
    tStream.Write base64M.decode(fBits)

    If tStream.Size > Int(UP_FileSize) Then
        tStream.Close
        Set tStream = Nothing
        Set base64M = Nothing
        Call returnError(0, "permitted to upload file.")
    End If
    If IsvalidFile(UCase(FixName(FileExt))) = False Then
        tStream.Close
        Set tStream = Nothing
        Set base64M = Nothing
        Call returnError(0, "permitted to upload file.")
    End If

    Dim fullPath
    fName = randomStr(1)&Year(Now)&Month(Now)&Day(Now)&Hour(Now)&Minute(Now)&Second(Now)&"."&FileExt
    fullPath = "attachments/"&D_Name&"/"&fName
    tStream.SaveToFile Server.MapPath(fullPath)
    If Err<>0 Then
        tStream.Close
        Set tStream = Nothing
        Set base64M = Nothing
        Call returnError(0, Err.Description)
        Err.Clear
    End If
    tStream.Close
    Set tStream = Nothing
    Set base64M = Nothing
    response.Write "<?xml version=""1.0"" encoding=""UTF-8""?><methodResponse><params><param><value><struct><member><name>url</name><value><string>"&fullPath&"</string></value></member></struct></value></param></params></methodResponse>"

End Function

'----------------------mt.getPostCategories--------------------------- 
Function getPostCategories(lID) 
    dim RecentPosts 
    Dim RS,dbRow,i 
    SQL="Select TOP 1 L.log_ID,L.log_Title,L.log_Author,L.log_Content,L.log_PostTime,L.log_edittype,C.cate_Name,L.log_IsDraft,C.cate_ID FROM blog_Content AS L,blog_Category AS C Where C.cate_ID=L.log_cateID And L.log_ID="&lID&" orDER BY L.log_PostTime DESC" 
    Set RS=Conn.ExeCute(SQL) 
    if RS.EOF or RS.BOF then 
        ReDim dbRow(0,0) 
        Call returnError(0,"Can't find log.") 
     else 
        dbRow=RS.getrows() 
    end if 
    RS.close 
    set RS=nothing 
    call CloseDB 
    RecentPosts="<?xml version=""1.0"" encoding=""UTF-8""?><methodResponse><params><param><value><array><data>" 
    i=0 
            RecentPosts = RecentPosts & "<value><struct>" 
            RecentPosts = RecentPosts & "<member><name>categoryId</name><value><string>"&CheckStr(dbRow(8,i))&"</string></value></member>" 
            RecentPosts = RecentPosts & "<member><name>categoryName</name><value><string>"&CheckStr(dbRow(6,i))&"</string></value></member>" 
            RecentPosts = RecentPosts & "</struct></value>" 
    RecentPosts = RecentPosts & "</data></array></value></param></params></methodResponse>"  
    response.write RecentPosts 
End Function


'------------------------------------------------------------------

Function bin2str(binstr)
    Dim varlen, clow, ccc, skipflag, i
	'中文字符Skip标志
    skipflag = 0
    ccc = ""
    If Not IsNull(binstr) Then
        varlen = LenB(binstr)
        For i = 1 To varlen
            If skipflag = 0 Then
                clow = MidB(binstr, i, 1)
				'判断是否中文的字符
                If AscB(clow)>127 Then
					'AscW会把二进制的中文双字节字符高位和低位反转，所以要先把中文的高低位反转
					ccc =ccc & Chr(AscW(MidB(binstr,i+1,1) & clow))
                    skipflag = 1
                Else
                    ccc = ccc & Chr(AscB(clow))
                End If
            Else
                skipflag = 0
            End If
        Next
    End If
    bin2str = ccc
End Function

Function AddSiteURL(ByVal Str)
    If IsNull(Str) Then
        AddSiteURL = ""
        Exit Function
    End If

    Dim re
    Set re = New RegExp
    With re
        .IgnoreCase = True
        .Global = True

        .Pattern = "<img (.*?)src=""(?!(http|https)://)(.*?)"""
        Str = .Replace(Str, "<img $1src=""" & SiteURL & "$3""")

        .Pattern = "<a (.*?)href=""(?!(http|https|ftp|mms|rstp)://)(.*?)"""
        Str = .Replace(Str, "<a $1href=""" & SiteURL & "$3""")
    End With
    Set re = Nothing

    AddSiteURL = Str
End Function

Function toInt(Str)
    Dim tmS, i
    tmS = ""
    toInt = ""
    For i = 1 To Len(Str)
        If Asc(Mid(Str, i, 1))>47 And Asc(Mid(Str, i, 1))<58 Then
            tmS = tmS&Mid(Str, i, 1)
        End If
    Next
    toInt = tmS
End Function

'=====================base64 encode/decode==============

Class base64
    Private objXmlDom
    Private objXmlNode

    Private Sub Class_Initialize()
        Set objXmlDom = Server.CreateObject(getXMLDOM())
    End Sub

    Private Sub Class_Terminate()
        Set objXmlDom = Nothing
    End Sub

    Public Function encode(AnsiCode)
        encode = ""
        Set objXmlNode = objXmlDom.createElement("file")
        objXmlNode.datatype = "bin.base64"
        objXmlNode.nodeTypedvalue = AnsiCode
        encode = objXmlNode.text
        Set objXmlNode = Nothing
    End Function

    Public Function decode(base64Code)
        decode = ""
        Set objXmlNode = objXmlDom.createElement("file")
        objXmlNode.datatype = "bin.base64"
        objXmlNode.text = base64Code
        decode = objXmlNode.nodeTypedvalue
        Set objXmlNode = Nothing
    End Function

End Class
%>
