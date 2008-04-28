<!--#include file="const.asp" -->
<!--#include file="conn.asp" -->
<!--#include file="common/cache.asp" -->
<!--#include file="common/function.asp" -->
<!--#include file="common/library.asp" -->
<!--#include file="common/ubbcode.asp" -->
<!--#include file="common/checkUser.asp" -->
<!--#include file="common/ModSet.asp" -->
<!--#include file="class/cls_logAction.asp" -->
<!--#include file="common/xml.asp" -->
<%
'======================================
'  XML-RPC for PJBlog2
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
Response.ContentType="text/xml"
Response.Expires=60
Response.Buffer=True

dim DebugOn
DebugOn=false

XmlBin=Request.BinaryRead(Request.TotalBytes)
'SaveToFile bin2str(XmlBin),"debug\out_"&randomStr(3)&"_"&Year(now)&Month(now)&Day(now)&Hour(now)&Minute(now)&Second(now)&".txt"
Dim xmlPrc,XmlBin
set xmlPrc=new PXML

if DebugOn then 
	xmlPrc.xmlPath = "out.xml"
	xmlPrc.Open()
 else
	xmlPrc.OpenXML(XmlBin)
end if

if xmlPrc.getError=0 then
        dim strAction
        dim userName,passWord,intPosts,bolPublish,logTitle,logDescription,logPostTime,logCategories,tagWords,logCategoryId,logID,fileBits,fileName
        strAction=xmlPrc.SelectXmlNodeText("methodName")
		Select Case strAction
			Case "blogger.getUsersBlogs":
			    userName=xmlPrc.SelectXmlNodeText("params/param[1]/value")
			    passWord=xmlPrc.SelectXmlNodeText("params/param[2]/value")
			    login2 userName,passWord
			    if stat_Admin then call getUsersBlogs() else Call returnError(0,"permitted to get Users Blogs.")
			Case "metaWeblog.getCategories":
			    userName=xmlPrc.SelectXmlNodeText("params/param[1]/value")
			    passWord=xmlPrc.SelectXmlNodeText("params/param[2]/value")
			    login2 userName,passWord
			    if stat_Admin then call getCategories() else Call returnError(0,"permitted to get Categories.")
			Case "mt.getCategoryList":
			    userName=xmlPrc.SelectXmlNodeText("params/param[1]/value")
			    passWord=xmlPrc.SelectXmlNodeText("params/param[2]/value")
			    login2 userName,passWord
			    if stat_Admin then call getCategories() else Call returnError(0,"permitted to get Categories.")
			Case "metaWeblog.getRecentPosts":
			    userName=xmlPrc.SelectXmlNodeText("params/param[1]/value")
			    passWord=xmlPrc.SelectXmlNodeText("params/param[2]/value")
				intPosts=xmlPrc.SelectXmlNodeText("params/param[3]/value/int")
				
				if intPosts=0 then
				    intPosts=xmlPrc.SelectXmlNodeText("params/param[3]/value/i4")
				end if
				
			    login2 userName,passWord
			    if stat_Admin then call getRecentPosts(intPosts) else Call returnError(0,"permitted to get Recent Posts.")
			Case "metaWeblog.newPost":
			    userName=xmlPrc.SelectXmlNodeText("params/param[1]/value")
			    passWord=xmlPrc.SelectXmlNodeText("params/param[2]/value")
				logTitle=xmlPrc.SelectXmlNodeText("params/param[3]/value/struct/member[name=""title""]/value")
				logDescription=xmlPrc.SelectXmlNodeText("params/param[3]/value/struct/member[name=""description""]/value")
				logPostTime=xmlPrc.SelectXmlNodeText("params/param[3]/value/struct/member[name=""dateCreated""]/value/dateTime.iso8601")
			    'tagWords=xmlPrc.GetXmlNodeLength("params/param[3]/value/struct/member[name=""tagwords""]/value/array/data/value")
			    bolPublish=xmlPrc.SelectXmlNodeText("params/param[4]/value/boolean")
			    login2 userName,passWord
			    if stat_Admin then Call newPost(logTitle,logDescription,logPostTime,bolPublish) else Call returnError(0,"permitted to post a new log.")
			Case "metaWeblog.editPost":
			    if xmlPrc.GetXmlNodeLength("params/param[0]/value/string")>0 then 
			      logID=xmlPrc.SelectXmlNodeText("params/param[0]/value/string")
			    else
			      logID=xmlPrc.SelectXmlNodeText("params/param[0]/value/struct/member[name=""PostID""]/value")
			    end if
			    if len(toInt(logID))<1 then Call returnError(0,"parameter error.")
			    logID=int(toInt(logID))
			    userName=xmlPrc.SelectXmlNodeText("params/param[1]/value")
			    passWord=xmlPrc.SelectXmlNodeText("params/param[2]/value")
				logTitle=xmlPrc.SelectXmlNodeText("params/param[3]/value/struct/member[name=""title""]/value")
				logDescription=xmlPrc.SelectXmlNodeText("params/param[3]/value/struct/member[name=""description""]/value")
				logPostTime=xmlPrc.SelectXmlNodeText("params/param[3]/value/struct/member[name=""dateCreated""]/value/dateTime.iso8601")
			    'tagWords=xmlPrc.GetXmlNodeLength("params/param[3]/value/struct/member[name=""tagwords""]/value/array/data/value")
			    bolPublish=xmlPrc.SelectXmlNodeText("params/param[4]/value/boolean")
			    login2 userName,passWord
			    if stat_Admin then Call editPost(logID,logTitle,logDescription,logPostTime,bolPublish) else Call returnError(0,"permitted to post a new log.")
			Case "mt.setPostCategories":
			    if xmlPrc.GetXmlNodeLength("params/param[0]/value/string")>0 then 
			      logID=xmlPrc.SelectXmlNodeText("params/param[0]/value/string")
			    else
			      logID=xmlPrc.SelectXmlNodeText("params/param[0]/value/struct/member[name=""PostID""]/value")
			    end if
			    if len(toInt(logID))<1 then Call returnError(0,"parameter error.")
			    logID=int(toInt(logID))
			    userName=xmlPrc.SelectXmlNodeText("params/param[1]/value")
			    passWord=xmlPrc.SelectXmlNodeText("params/param[2]/value")
				logCategoryId=xmlPrc.SelectXmlNodeText("params/param[3]/value/array/data/value[0]/struct/member[name=""categoryId""]/value")
 			    login2 userName,passWord
			    if stat_Admin then Call setPostCategories(logID,logCategoryId) else Call returnError(0,"permitted to set post categories.")
			Case "metaWeblog.getPost":
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
			    if stat_Admin then Call getPost(logID) else Call returnError(0,"permitted to set post categories.")
			Case "metaWeblog.newMediaObject":
			    userName=xmlPrc.SelectXmlNodeText("params/param[1]/value")
			    passWord=xmlPrc.SelectXmlNodeText("params/param[2]/value")
				fileBits=xmlPrc.SelectXmlNodeText("params/param[3]/value/struct/member[name=""bits""]/value")
		        fileName=xmlPrc.SelectXmlNodeText("params/param[3]/value/struct/member[name=""name""]/value")
 			    login2 userName,passWord
			    if stat_Admin then Call newMediaObject(fileName,fileBits) else Call returnError(0,"permitted to upload file.")
			Case "blogger.deletePost":
			    if xmlPrc.GetXmlNodeLength("params/param[1]/value/string")>0 then 
			      logID=xmlPrc.SelectXmlNodeText("params/param[1]/value/string")
			    else
			      logID=xmlPrc.SelectXmlNodeText("params/param[1]/value/struct/member[name=""PostID""]/value")
			    end if
			    if len(toInt(logID))<1 then Call returnError(0,"parameter error.")
			    logID=int(toInt(logID))
			    userName=xmlPrc.SelectXmlNodeText("params/param[2]/value")
			    passWord=xmlPrc.SelectXmlNodeText("params/param[3]/value")
 			    login2 userName,passWord
			    if stat_Admin then Call deletePost(logID) else Call returnError(0,"permitted to delete log.")
 			 Case Else
				xmlPrc.CloseXml()
				Call returnError(0,strAction)
		End Select 
 else
	xmlPrc.CloseXml()
	Call returnError(0,"action Error2.")
end if

'=================Function In XML-PRC=========================

'----------------------return response Error---------------------------
Function returnError(faultCode,faultString)
	Response.Clear
	Response.write "<?xml version=""1.0"" encoding=""UTF-8""?><methodResponse>"
	Response.write "<fault><value><struct>"
	Response.write "<member><name>faultCode</name><value><int>"&faultCode&"</int></value></member>"
	Response.write "<member><name>faultString</name><value><string>"&faultString&"</string></value></member>"
	Response.write "</struct></value></fault>"
	Response.write "</methodResponse>"
	Response.End
End Function

'----------------------blogger.getUsersBlogs---------------------------
Function getUsersBlogs()
	Response.Clear
	Response.write "<?xml version=""1.0"" encoding=""UTF-8""?>"
	Response.write "<methodResponse><params><param><value><array><data><value><struct>"
	Response.write "<member><name>url</name><value><string><![CDATA["&SiteURL&"]]></string></value></member>"
	Response.write "<member><name>blogid</name><value><string><![CDATA["&CookieName&"]]></string></value></member>"
	Response.write "<member><name>blogName</name><value><string><![CDATA["&UnCheckStr(SiteName)&"]]></string></value></member>"
	Response.write "</struct></value></data></array></value></param></params></methodResponse>"
	Response.End
End Function


'----------------------metaWeblog.getRecentPosts---------------------------
Function getRecentPosts(intNum)
    dim RecentPosts
	
	Dim RS,dbRow,i
	
	SQL="SELECT TOP "&intNum&" L.log_ID,L.log_Title,L.log_Author,L.log_Content,L.log_PostTime,L.log_edittype,C.cate_Name,L.log_IsDraft FROM blog_Content AS L,blog_Category AS C WHERE C.cate_ID=L.log_cateID ORDER BY L.log_PostTime DESC"
	Set RS=Conn.ExeCute(SQL)
	if RS.EOF or RS.BOF then
		ReDim dbRow(0,0)
	 else
		dbRow=RS.getrows()
	end if
	RS.close
	set RS=nothing
	call CloseDB
	
	RecentPosts="<?xml version=""1.0"" encoding=""UTF-8""?><methodResponse><params><param><value><array><data>"
    
	if ubound(dbRow,1)<>0 then
		for i=0 to ubound(dbRow,2)
			RecentPosts = RecentPosts & "<value><struct>"
		    RecentPosts = RecentPosts & "<member><name>title</name><value><string><![CDATA["&dbRow(1,i)&"]]></string></value></member>"
		    if dbRow(5,i)=1 then
			    t=AddSiteURL(UBBCode(HTMLEncode(dbRow(3,i)),0,0,0,1,1))
			    RecentPosts = RecentPosts & "<member><name>description</name><value><string><![CDATA["&t&"]]></string></value></member>"
		     else
			    t=AddSiteURL(UnCheckStr(dbRow(3,i)))
			    RecentPosts = RecentPosts & "<member><name>description</name><value><string><![CDATA["&t&"]]></string></value></member>"
		    end if
		    RecentPosts = RecentPosts & "<member><name>dateCreated</name><value><dateTime.iso8601>"&DateToStr(dbRow(4,i),"y-m-dTH:I:S")&"</dateTime.iso8601></value></member>"
		    RecentPosts = RecentPosts & "<member><name>categories</name><value><array><data><value><string>"&dbRow(6,i)&"</string></value></data></array></value></member>"
		    RecentPosts = RecentPosts & "<member><name>tagwords</name><value><string><![CDATA[df34,err]]></string></value></member>"
		    RecentPosts = RecentPosts & "<member><name>postid</name><value><string><![CDATA["&dbRow(0,i)&"]]></string></value></member>"
		    RecentPosts = RecentPosts & "<member><name>userid</name><value><string><![CDATA["&dbRow(2,i)&"]]></string></value></member>"
		    RecentPosts = RecentPosts & "<member><name>link</name><value><string><![CDATA["&SiteURL&"/default.asp?id="&dbRow(0,i)&"]]></string></value></member>"
		    RecentPosts = RecentPosts & "<member><name>permaLink</name><value><string><![CDATA["&SiteURL&"/default.asp?id="&dbRow(0,i)&"]]></string></value></member>"
			RecentPosts = RecentPosts & "</struct></value>"
		next
	end if
	RecentPosts = RecentPosts & "</data></array></value></param></params></methodResponse>"
	
	'SaveToFile RecentPosts,"debug\recent2_"&randomStr(3)&"_"&Year(now)&Month(now)&Day(now)&Hour(now)&Minute(now)&Second(now)&".txt"
	
	Response.Clear
	Response.write RecentPosts
End Function

'----------------------metaWeblog.getPost---------------------------
Function getPost(lID)
    dim RecentPosts
	Dim RS,dbRow,i
	SQL="SELECT TOP 1 L.log_ID,L.log_Title,L.log_Author,L.log_Content,L.log_PostTime,L.log_edittype,C.cate_Name,L.log_IsDraft FROM blog_Content AS L,blog_Category AS C WHERE C.cate_ID=L.log_cateID And L.log_ID="&lID&" ORDER BY L.log_PostTime DESC"
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
	
	RecentPosts="<?xml version=""1.0"" encoding=""UTF-8""?><methodResponse><params><param>"
	i=0
			RecentPosts = RecentPosts & "<value><struct>"
		    RecentPosts = RecentPosts & "<member><name>title</name><value><string><![CDATA["&dbRow(1,i)&"]]></string></value></member>"
		    if dbRow(5,i)=1 then
			    t=AddSiteURL(UBBCode(HTMLEncode(dbRow(3,i)),0,0,0,1,1))
			    RecentPosts = RecentPosts & "<member><name>description</name><value><string><![CDATA["&t&"]]></string></value></member>"
		     else
			    t=AddSiteURL(UnCheckStr(dbRow(3,i)))
			    RecentPosts = RecentPosts & "<member><name>description</name><value><string><![CDATA["&t&"]]></string></value></member>"
		    end if
		    RecentPosts = RecentPosts & "<member><name>dateCreated</name><value><dateTime.iso8601>"&DateToStr(dbRow(4,i),"y-m-dTH:I:S")&"</dateTime.iso8601></value></member>"
		    RecentPosts = RecentPosts & "<member><name>categories</name><value><array><data><value><string>"&dbRow(6,i)&"</string></value></data></array></value></member>"
		    RecentPosts = RecentPosts & "<member><name>tagwords</name><value><string><![CDATA[df34,err]]></string></value></member>"
		    RecentPosts = RecentPosts & "<member><name>postid</name><value><string><![CDATA["&dbRow(0,i)&"]]></string></value></member>"
		    RecentPosts = RecentPosts & "<member><name>userid</name><value><string><![CDATA["&dbRow(2,i)&"]]></string></value></member>"
		    RecentPosts = RecentPosts & "<member><name>link</name><value><string><![CDATA["&SiteURL&"/default.asp?id="&dbRow(0,i)&"]]></string></value></member>"
		    RecentPosts = RecentPosts & "<member><name>permaLink</name><value><string><![CDATA["&SiteURL&"/default.asp?id="&dbRow(0,i)&"]]></string></value></member>"
			RecentPosts = RecentPosts & "</struct></value>"
	RecentPosts = RecentPosts & "</param></params></methodResponse>"
	response.write RecentPosts
End Function

'----------------------mt.getCategoryList/metaWeblog.getCategories---------------------------
function getCategories()
	dim Categories
	Categories = "<?xml version=""1.0"" encoding=""UTF-8""?><methodResponse><params><param><value><array><data>"
    Dim Arr_Category,Category_Len,i
 	CategoryList(1)
    Arr_Category=Application(CookieName&"_blog_Category")
    if ubound(Arr_Category,1)=0 then Call returnError(0,"no Categories")
	Category_Len=ubound(Arr_Category,2)
	
	For i=0 to Category_Len
		  if not Arr_Category(4,i) then
		    Categories = Categories & "<value><struct>"
		    Categories = Categories & "<member><name>description</name><value><string><![CDATA["&Arr_Category(1,i)&"]]></string></value></member>"
		    Categories = Categories & "<member><name>httpUrl</name><value><string><![CDATA["&SiteURL&"/default.asp?cateID="&Arr_Category(0,i)&"]]></string></value></member>"
		    Categories = Categories & "<member><name>rssUrl</name><value><string><![CDATA["&SiteURL&"/feed.asp?cateID="&Arr_Category(0,i)&"]]></string></value></member>"
		    Categories = Categories & "<member><name>title</name><value><string><![CDATA["&Arr_Category(1,i)&"]]></string></value></member>"
		    Categories = Categories & "<member><name>categoryId</name><value><string><![CDATA["&Arr_Category(0,i)&"]]></string></value></member>"
		    Categories = Categories & "<member><name>categoryName</name><value><string><![CDATA["&Arr_Category(1,i)&"]]></string></value></member>"
		    Categories = Categories & "</struct></value>"
		end if
	Next
	Categories = Categories & "</data></array></value></param></params></methodResponse>"
	Response.write Categories
end function

'----------------------metaWeblog.newPost---------------------------
function newPost(lTitle,lDescription,lPostTime,lPublish)
      '====get Last Category
	  Dim Arr_Category,Category_Len,i,lastCID
	  CategoryList(1)
	  Arr_Category=Application(CookieName&"_blog_Category")
	  if ubound(Arr_Category,1)=0 then Call returnError(0,"no Categories")
	  Category_Len=ubound(Arr_Category,2)
	  For i=0 to Category_Len
		  if not Arr_Category(4,i) then
		   lastCID=Arr_Category(0,i)
		   exit for
		end if
      Next
	  dim newPostStr
      dim lArticle,postLog
      set lArticle=new logArticle
         lArticle.categoryID = lastCID
         lArticle.logTitle = lTitle
         lArticle.logEditType = 0
         lArticle.logIsDraft = not cbool(lPublish)
         lArticle.logMessage = lDescription
         lArticle.logPubTime = Now()
         lArticle.logIntroCustom = 0
         lArticle.logAuthor = memName
	 	 postLog = lArticle.postLog
    set lArticle=nothing
   if postLog(2)<0 then 
	    Call returnError(0,postLog(1))
    else
		newPostStr = "<?xml version=""1.0"" encoding=""UTF-8""?><methodResponse><params><param>"
		newPostStr = newPostStr & "<value><string>"&postLog(2)&"</string></value>"
		newPostStr = newPostStr & "</param></params></methodResponse>"
		response.write newPostStr
    end if
end function

'----------------------metaWeblog.editPost---------------------------
function editPost(lID,lTitle,lDescription,lPostTime,lPublish)
      '====get Last Category
	  dim editPostStr
      dim lArticle,editLog
      set lArticle=new logArticle
      editLog = lArticle.getLog(lID)
	  if editLog(0)<0 then  Call returnError(0,"Can't find log.")
	  
         lArticle.logTitle = lTitle
         lArticle.logEditType = 0
         lArticle.logIsDraft = not cbool(lPublish)
         lArticle.logMessage = lDescription
         lArticle.logPubTime = Now()
         lArticle.logIntroCustom = 0
	 	 editLog = lArticle.editLog(lID)
    set lArticle=nothing
   if editLog(2)<0 then 
	    Call returnError(0,editLog(1))
    else
		editPostStr = "<?xml version=""1.0"" encoding=""UTF-8""?><methodResponse><params><param>"
		editPostStr = editPostStr & "<value><struct>"
		editPostStr = editPostStr & "<member><name>PostID</name><value><string>"&editLog(2)&"</string></value></member>"
		editPostStr = editPostStr & "</struct></value>"
		editPostStr = editPostStr & "</param></params></methodResponse>"
		response.write editPostStr
    end if
end function



'----------------------mt.setPostCategories---------------------------
function setPostCategories(lID,lCID)
    dim returnStr
    dim lArticle,editLog
    set lArticle=new logArticle
    editLog = lArticle.getLog(lID)
    set lArticle=nothing
   if editLog(0)<0 then
     Call returnError(0,"Can't find log.")
   else
         if IsInteger(lCID) then
	         dim lastCID
	         lastCID=conn.Execute("select top 1 log_cateID from blog_Content where log_ID="&lID)(0)
	         Conn.ExeCute("UPDATE blog_Category SET cate_count=cate_count-1 where cate_ID="&lastCID)
	         Conn.ExeCute("UPDATE blog_Category SET cate_count=cate_count+1 where cate_ID="&lCID)
	         Conn.ExeCute("UPDATE blog_Content SET log_cateID="&lCID&" where log_ID="&lID)
    	 end if
    
    if blog_postFile then
		 set lArticle = new ArticleCache
		 lArticle.SaveCache
		 set lArticle = nothing
         PostArticle lID
	 end if
         
	returnStr="<?xml version=""1.0"" encoding=""UTF-8""?><methodResponse><params><param><value><boolean>1</boolean></value></param></params></methodResponse>"
	response.write returnStr
   end if
end function

function deletePost(lID)
     dim lArticle,DeleteLog,returnStr
     set lArticle=new logArticle
	 DeleteLog=lArticle.deleteLog(lID)
     set lArticle=nothing
     if DeleteLog(0)=0 then
	 	 returnStr="<?xml version=""1.0"" encoding=""UTF-8""?><methodResponse><params><param><value><boolean>1</boolean></value></param></params></methodResponse>"
	 	 response.write returnStr
	 else
	     Call returnError(0,"Can't delete log.")
     end if
end function

'----------------------metaWeblog.newMediaObject"--------------------------

function newMediaObject(fName,fBits) 
 'On Error Resume Next
   IF stat_FileUpLoad=false Then Call returnError(0,"permitted to upload file.")
	Dim upl,FSOIsOK
		FSOIsOK=1
		Set upl=Server.CreateObject("Scripting.FileSystemObject")
		If Err<>0 Then
			Err.Clear
		    Set upl=Nothing
		     Call returnError(0,"Can't create folder.")
		End If
		Dim D_Name
			D_Name="month_"&DateToStr(Now(),"ym")
			If upl.FolderExists(Server.MapPath("attachments/"&D_Name))=False Then
				upl.CreateFolder Server.MapPath("attachments/"&D_Name)
			End If
    
    dim FileExt,i
    FileExt=""
    for i=len(fName) to 1 step -1
     if mid(fName,i,1)="." then exit for
     FileExt = mid(fName,i,1) & FileExt
    next
    dim tStream,base64M
    set base64M = new base64
    set tStream = Server.CreateObject("adodb.stream")
    tStream.Type = 1
	tStream.Mode = 3
	tStream.Open
	tStream.Position = 0
    tStream.write base64M.decode(fBits)
    
    if tStream.size > Int(UP_FileSize) then
      tStream.close
      set tStream = nothing
      set base64M = nothing
      Call returnError(0,"permitted to upload file.")
    end if
    if IsvalidFile(UCase(FixName(FileExt))) = False then
      tStream.close
      set tStream = nothing
      set base64M = nothing
      Call returnError(0,"permitted to upload file.")
    end if
    
    dim fullPath
	fName=randomStr(1)&Year(now)&Month(now)&Day(now)&Hour(now)&Minute(now)&Second(now)&"."&FileExt
	fullPath="attachments/"&D_Name&"/"&fName
    tStream.SaveToFile Server.MapPath(fullPath)
		If Err<>0 Then
		    tStream.close
		    set tStream = nothing
		    set base64M = nothing
		    Call returnError(0,err.Description)
			Err.Clear
		End If
      tStream.close
      set tStream = nothing
      set base64M = nothing
	response.write "<?xml version=""1.0"" encoding=""UTF-8""?><methodResponse><params><param><value><struct><member><name>url</name><value><string>"&fullPath&"</string></value></member></struct></value></param></params></methodResponse>"

end function
'------------------------------------------------------------------
Function bin2str(binstr)
	Dim varlen,clow,ccc,skipflag,i
	'中文字符Skip标志
	skipflag=0
	ccc = ""
	If Not IsNull(binstr) Then
		varlen=LenB(binstr)
		For i=1 To varlen
			If skipflag=0 Then
				clow = MidB(binstr,i,1)
				'判断是否中文的字符
				If AscB(clow)>127 Then
					'AscW会把二进制的中文双字节字符高位和低位反转，所以要先把中文的高低位反转
					ccc =ccc & Chr(AscW(MidB(binstr,i+1,1) & clow))
					skipflag=1
				Else
					ccc = ccc & Chr(AscB(clow))
				End If
			Else
				skipflag=0
			End If
		Next
	End If
	bin2str = ccc
End Function

function AddSiteURL(ByVal Str)
  If IsNull(Str) Then
  	AddSiteURL = ""
  	Exit Function 
  End If

  Dim re
  Set re=new RegExp
  With re
    .IgnoreCase =True
    .Global=True
    
    .Pattern="<img (.*?)src=""(?!(http|https)://)(.*?)"""
    str = .replace(str,"<img $1src=""" & SiteURL & "$3""")
    
    .Pattern="<a (.*?)href=""(?!(http|https|ftp|mms|rstp)://)(.*?)"""
    str = .replace(str,"<a $1href=""" & SiteURL & "$3""")
  End With
  Set re=Nothing

  AddSiteURL=Str
End Function

function toInt(str)
  dim tmS,i
  tmS=""
  toInt=""
	 for i=1 to len(str)
	   if asc(mid(str,i,1))>47 and asc(mid(str,i,1))<58 then 
	     tmS=tmS&mid(str,i,1)
	   end if 
	 next
	 toInt=tmS
end function

'=====================base64 encode/decode==============
class base64
  Private objXmlDom
  Private objXmlNode
  
  Private Sub Class_Initialize()
      Set objXmlDom = Server.CreateObject(getXMLDOM())
  end sub
  
  Private Sub Class_Terminate()
	  Set objXmlDom =nothing
  end sub

  public function encode(AnsiCode)
    encode=""
	Set objXmlNode = objXmlDom.createElement("file")
	objXmlNode.datatype = "bin.base64"
    objXmlNode.nodeTypedvalue = AnsiCode
    encode = objXmlNode.text
	Set objXmlNode =nothing
  end function
  
  public function decode(base64Code)
    decode=""
	Set objXmlNode = objXmlDom.createElement("file")
	objXmlNode.datatype = "bin.base64"
	objXmlNode.text = base64Code
    decode = objXmlNode.nodeTypedvalue
 	Set objXmlNode =nothing
  end function
end class
%>
