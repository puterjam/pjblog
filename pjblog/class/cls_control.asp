<%
'==================================
'  后台库文件
'    更新时间: 2006-5-21
'==================================

'----------- 显示操作信息 ----------------------------
sub getMsg 
	 if session(CookieName&"_ShowMsg")=true then
	  response.write ("<div id=""msgInfo"" align=""center""><img src=""images/Control/aL.gif"" style=""margin-bottom:-11px;""/><span class=""alertTxt"">" & session(CookieName&"_MsgText") & "</span><img src=""images/Control/aR.gif"" style=""margin-bottom:-11px;""/></div>")
	  response.write ("<script>setTimeout('hiddenMsg()',3000);function hiddenMsg(){document.getElementById('msgInfo').style.display='none';}</script>")
	  session(CookieName&"_ShowMsg")=false
	  session(CookieName&"_MsgText")=""
	 end If
end sub
  
'----------- 获得路径的文件信息 ----------------------------
function getPathList(pathName)
 dim FSO,ServerFolder,getInfo,getInfos,tempS
 getInfo=""
		Set FSO=Server.CreateObject("Scripting.FileSystemObject")
		
		Set ServerFolder=FSO.GetFolder(Server.MapPath(pathName))
			Dim ServerFolderList,ServerFolderEvery
			Set ServerFolderList=ServerFolder.SubFolders
			tempS=""
			For Each ServerFolderEvery IN ServerFolderList
                getInfo=getInfo&tempS&ServerFolderEvery.Name
                tempS="*"
			Next
            getInfo=getInfo&"|"
			Dim ServerFileList,ServerFileEvery
			Set ServerFileList=ServerFolder.Files
			tempS=""
			For Each ServerFileEvery IN ServerFileList
                getInfo=getInfo&tempS&ServerFileEvery.Name
                tempS="*"
			Next
	Set FSO=Nothing
	getInfos=split(getInfo,"|")
	getPathList=getInfos
end function

'----------- 获取文件信息 ----------------------------
function getFileInfo(FileName)
 dim FSO,File,FileInfo(3)
 Set FSO=Server.CreateObject("Scripting.FileSystemObject")
 if FSO.FileExists(Server.MapPath(FileName)) then
   Set File=FSO.GetFile(Server.MapPath(FileName))
   FileInfo(0)=File.Size
   if FileInfo(0)/1000>1 then 
     FileInfo(0)=int(FileInfo(0)/1000)&" KB"
    else
     FileInfo(0)=FileInfo(0)&" Bytes"
   end if
   FileInfo(1)=lcase(right(FileName,4))
   FileInfo(2)=File.DateCreated
   FileInfo(3)=File.Type 
 end if
   getFileInfo=FileInfo
 Set FSO=Nothing
end function

'----------- 获取文件图标 ----------------------------
function getFileIcons(str) 
 dim FileIcon,Target
 Select Case str
  case ".jpg"
   FileIcon="jpg.gif"
  case ".gif"
   FileIcon="gif.gif"
  case ".bmp"
   FileIcon="bmp.gif"
  case ".png"
   FileIcon="png.gif"
 case ".zip"
   FileIcon="zip.gif"  
 case ".rar"
   FileIcon="rar.gif"  
 case ".swf"
   FileIcon="swf.gif"  
 case ".mdb"
   FileIcon="mdb.gif"  
 case ".doc"
   FileIcon="doc.gif"  
 case ".xls"
   FileIcon="xls.gif"  
 case ".pdf"
   FileIcon="pdf.gif"  
 case ".mbk"
   FileIcon="mbk.gif"
 case ".mp3"
   FileIcon="mp3.gif"
 case ".wmv"
   FileIcon="wma.gif"
 case ".wma"
   FileIcon="wma.gif"
 case else
   FileIcon="unknow.gif"
 end Select
 getFileIcons="<img border=""0"" src=""images/file/"&FileIcon&""" style=""margin:4px 3px -3px 0px""/>"
end function


'----------- 获得目标大小 ----------------------------
Function GetTotalSize(GetLocal,GetType)
	On Error Resume Next
	Dim FSO
	Set FSO=Server.CreateObject("Scripting.FileSystemObject")
	IF Err<>0 Then
		Err.Clear
		GetTotalSize="Fail"
	Else
		Dim SiteFolder
		IF GetType="Folder" Then
			Set SiteFolder=FSO.GetFolder(GetLocal) 
		Else
			Set SiteFolder=FSO.GetFile(GetLocal) 
		End IF
		GetTotalSize=SiteFolder.Size
		IF GetTotalSize>1024*1024 Then
		GetTotalSize=GetTotalSize/1024/1024
		IF inStr(GetTotalSize,".") Then GetTotalSize = Left(GetTotalSize,inStr(GetTotalSize,".")+2)
			GetTotalSize=GetTotalSize&" MB"
		Else
			GetTotalSize=Fix(GetTotalSize/1024)&" KB"
		End IF
		
		Set SiteFolder=Nothing
	End IF
	Set FSO=Nothing
End Function

Function CopyFiles(TempSource,TempEnd) '复制文件
    Dim CopyFSO
    Set CopyFSO = Server.CreateObject("Scripting.FileSystemObject")
	IF CopyFSO.FileExists(TempEnd) then
       CopyFiles="目标备份文件 <b>" & TempEnd & "</b> 已存在，请先删除!"
       Set CopyFSO=Nothing
       Exit Function
    End If
    IF CopyFSO.FileExists(TempSource) Then
    Else
       CopyFiles="要复制的源数据库文件 <b>"&TempSource&"</b> 不存在!"
       Set CopyFSO=Nothing
       Exit Function
    End If
    CopyFSO.CopyFile TempSource,TempEnd
    CopyFiles="已经成功复制文件 <b>"&TempSource&"</b> 到 <b>"&TempEnd&"</b>"
    Set CopyFSO = Nothing
End Function

Function DeleteFiles(FilePath) '删除文件
	On Error Resume Next
    Dim FSO
    Set FSO=Server.CreateObject("Scripting.FileSystemObject")
    IF FSO.FileExists(FilePath) Then
		FSO.DeleteFile FilePath,True
		if err then
			Set FSO = Nothing
			DeleteFiles = false
			err.clear
			exit function
		end if
		DeleteFiles = True
    Else
		DeleteFiles = false
    End IF
    Set FSO = Nothing
End Function

Function FileExist(FilePath) '判断是否存在文件
  FileExist=false
  Dim FSO
  Set FSO=Server.CreateObject("Scripting.FileSystemObject")
  FilePath=Server.MapPath(FilePath)
  IF FSO.FileExists(FilePath) Then FileExist=true
end Function

function DisI(b)
 if b then
   response.write "<span style=""color:#00cc00""><b>支持</b></span>"
  else
   response.write "<span style=""color:#FF0000""><b>不支持</b></span>"
 end if
end function

function getModName '获得模块名字
  dim Bmodules,Bmodule,tempB
   tempB=""
   Set blog_module=Conn.ExeCute("SELECT * FROM blog_module")
   do until blog_module.eof
    Bmodules=Bmodules&tempB&blog_module("name")
    tempB="|&|"
    blog_module.movenext
   loop
   Bmodule=split(Bmodules,"|&|")
   set blog_module=nothing
   getModName=Bmodule
end function


'----------- 释放网站缓存 ----------------------------
Function FreeApplicationMemory
    on error resume next
	Response.Write "释放网站缓存数据列表：<div style='padding:5px 5px 5px 10px;'>"
	Dim Thing
	For Each Thing IN Application.Contents
		IF Left(Thing,Len(CookieName)) = CookieName Then
			Response.Write "<span style='color:#666'>" & thing & "</span><br/>"
			IF isObject(Application.Contents(Thing)) Then
				Application.Contents(Thing).Close
				Set Application.Contents(Thing) = Nothing
				Application.Contents(Thing) = Null
			ElseIF isArray(Application.Contents(Thing)) Then
				Set Application.Contents(Thing) = Nothing
				Application.Contents(Thing) = Null
			Else
				Application.Contents(Thing) = Null
			End IF
		End IF
	Next
	response.write "</div>"
End Function

'----------- 更新缓存信息 ----------------------------
Function FreeMemory
  call reloadcache
  	 if blog_postFile then
	     dim lArticle
		 set lArticle = new ArticleCache
		 lArticle.SaveCache
		 set lArticle = nothing
	 end if
End Function

function InstallPlugingSetting(Path,ModName,action) '安装插件配置文件
   InstallPlugingSetting=-1
   dim SettingXML,KeyLen,tempI,ModSetting,KeyName,KeyValue
   If action="delete" Then 
    conn.Execute "delete * from blog_ModSetting where set_ModName='"&ModName&"'"
    InstallPlugingSetting=0
    exit function 
   End If
   
   If action="insert" Then
	   dim modSetC,tmpC
	   tmpC=conn.execute("select count(set_id) as sC from [blog_ModSetting] where set_ModName='"&ModName&"'")(0)
       if tmpC>0 then 
		   InstallPlugingSetting=0
		   exit function
       end if
	   set SettingXML=new PXML
	   SettingXML.XmlPath=Path
	   SettingXML.Open
		   if SettingXML.getError=0 Then
			    KeyLen = int(SettingXML.GetXmlNodeLength("PluginOptions/Key"))
			    For tempI=0 To KeyLen-1
			          KeyName=SettingXML.GetAttributes("PluginOptions/Key","name",tempI)
			          KeyValue=SettingXML.GetAttributes("PluginOptions/Key","value",tempI)
			          ModSetting=Array(Array("set_ModName",ModName),Array("set_KeyName",KeyName),Array("set_KeyValue",KeyValue))
			          DBQuest "blog_ModSetting",ModSetting,"insert"
			    Next
		   end If
   End If
   InstallPlugingSetting=0
end Function

sub delPlugins(PluginsName) '删除插件
   conn.execute("delete * from blog_module where name='"&PluginsName&"'")
end sub

sub delModule(ModID) '删除模块
	  conn.execute("DELETE * from blog_module where id="&moduleID)
end sub

sub SavehtmlCode(HCode,modID) '写HTML代码到数据库
     conn.execute("update blog_module set HtmlCode='"& HCode & "' where id=" & modID)
end sub

 sub PostLink()
   Dim LoadTemplate,Temp,SaveArticle
    LoadTemplate=LoadFromFile("Template/Link.asp")
    if LoadTemplate(0)=0 then '读取成功后写入信息
         Temp=LoadTemplate(1)
          dim blog_Links,ImgLink,TextLink
          set blog_Links=conn.execute("select * from blog_Links where link_IsShow=true order by link_Order asc")
          do until blog_Links.eof
           if len(blog_Links("link_Image"))>0 then
              ImgLink=ImgLink&"<a href="""&blog_Links("link_URL")&""" target=""_blank""><img src="""&blog_Links("link_Image")&""" alt="""&blog_Links("link_Name")&""" border=""0"" style=""margin:3px;width:88px;height:31px""/></a>"
            else
              TextLink=TextLink&"<div class=""link"" style=""width:108px;float:left;overflow:hidden;margin-right:8px;height:24px;line-height:180%""><a href="""&blog_Links("link_URL")&""" target=""_blank"" title="""&blog_Links("link_Name")&""">"&blog_Links("link_Name")&"</a></div>"
           end if
           blog_Links.movenext
          loop
          Temp=Replace(Temp,"<$ImgLink$>",ImgLink)
          Temp=Replace(Temp,"<$TextLink$>",TextLink)
          SaveArticle=SaveToFile(Temp,"post/link.html")
    end if
 end sub
 
 
'=========================================
'安装插件
'=========================================
 sub InstallPlugins 
  On Error Resume Next
  dim PlugName,PlugTitle,PlugType,PlugSortID,PlugHtmlCode,PlugPluginPath,PlugSettingXML,MenuID,ConfigPath
  dim MainItem,SubItem,CreateTableSQL,UpdateTableSQL
  dim tmpC
   MenuID=-1
   ConfigPath=""
   Bmodule=getModName
   PluginsFolder=checkstr(Request.QueryString("Plugins"))
   set PluginsXML=New PXML
   PluginsXML.XmlPath="Plugins/"&PluginsFolder&"/install.xml"
   PluginsXML.open
   if PluginsXML.getError=0 then
     	 if len(PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginName"))>0 and IsvalidPlugins(PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginName")) and (not IsvalidValue(Bmodule,PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginName"))) and IsvalidValue(TypeArray,PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginType")) then
		            PlugName = PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginName")
		            PlugTitle = PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginTitle")
		            PlugType = PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginType")
		            PlugSortID = int(conn.execute("select count(*) from blog_module")(0)) * -1
		            PlugHtmlCode = PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginHtmlCode")
		            PlugSettingXML = PluginsXML.SelectXmlNodeText("PluginInstall/main/SettingFile")
		            ConfigPath = PluginsXML.SelectXmlNodeText("PluginInstall/main/ConfigPath")
		            If Err Then
		               ConfigPath=""
		               Err.clear
		            End If
		            
		            PlugPluginPath = ""
		            if PlugType="function" Then '判断模块插件
		              tempI=conn.Execute("select count(*) from blog_Category")(0)
		              PlugPluginPath = PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginPath")
		              if len(PlugPluginPath)>0 then '插入插件菜单
			              MainItem=Array(Array("cate_Name",PlugTitle),Array("cate_Intro",PlugTitle),Array("cate_OutLink",True),Array("cate_URL","LoadMod.asp?plugins="&PlugName),Array("cate_icon","images/icons/1.gif"),Array("cate_Order",tempI))
			              DBQuest "blog_Category",MainItem,"insert"
			              CategoryList(2)
			              MenuID=conn.Execute("select top 1 cate_ID from blog_Category order by cate_ID desc")(0)
			          end if
		            End If
		            
		            '写入模块所需要的SQL语句
		           CreateTableSQL = PluginsXML.SelectXmlNodeText("PluginInstall/main/CreateTableSQL")
		           UpdateTableSQL = PluginsXML.SelectXmlNodeText("PluginInstall/main/UpdateTableSQL")
		              
		           if len(CreateTableSQL)>0 then conn.Execute(CreateTableSQL)
		           if len(UpdateTableSQL)>0 then conn.Execute(UpdateTableSQL)
                   if len(PlugSettingXML)>0 And InstallPlugingSetting("Plugins/"&PluginsFolder&"/"&PlugSettingXML,PlugName,"insert")=0 then '检测并安装插件配置文件
                       MainItem=array(array("name",PlugName),array("title",PlugTitle),array("type",PlugType),array("SortID",PlugSortID),array("HtmlCode",PlugHtmlCode),array("PluginPath",PlugPluginPath),array("IsInstall",-1),array("InstallFolder",PluginsFolder),array("ConfigPath",ConfigPath),array("SettingXML",PlugSettingXML),Array("CateID",MenuID))
                       Dim ModSetTemp
                       Set ModSetTemp=New ModSet
                       ModSetTemp.Open PlugName
                       ModSetTemp.ReLoad()
                    else
                       MainItem=Array(array("name",PlugName),array("title",PlugTitle),array("type",PlugType),array("SortID",PlugSortID),array("HtmlCode",PlugHtmlCode),array("PluginPath",PlugPluginPath),array("IsInstall",-1),array("InstallFolder",PluginsFolder),array("ConfigPath",ConfigPath),Array("CateID",MenuID))
                   end If
                   
                   DBQuest "blog_module",MainItem,"insert" '插入插件信息
                   SubItemLen = int(PluginsXML.GetXmlNodeLength("PluginInstall/SubItem/item"))
                   
                   for tempI=0 to SubItemLen-1 '安装插件辅助信息
                     PlugType = PluginsXML.SelectXmlNodeItemText("PluginInstall/SubItem/item/PluginType",tempI)
                        if not PlugType="function" then
                             PlugTitle = PluginsXML.SelectXmlNodeItemText("PluginInstall/SubItem/item/PluginTitle",tempI)
                             PlugSortID = int(conn.execute("select count(*) from blog_module")(0)) * -1
                             PlugHtmlCode = PluginsXML.SelectXmlNodeItemText("PluginInstall/SubItem/item/PluginHtmlCode",tempI)
                             SubItem=array(array("name",PlugName&"SubItem"&(tempI+1)),array("title",PlugTitle),array("type",PlugType),array("SortID",PlugSortID),array("HtmlCode",PlugHtmlCode))
                             DBQuest "blog_module",SubItem,"insert" 
                        end if
                   next
           else
             session(CookieName&"_ShowMsg")=true
             Session(CookieName&"_MsgText")="<font color=""#ff0000"">安装失败</font> 安装插件和其他插件造成冲突，或插件已经安装!"
             Response.Redirect("ConContent.asp?Fmenu=Skins&Smenu=PluginsInstall")
         end if
       PluginsXML.CloseXml()
       log_module(2)
       FixPlugins(0)
       session(CookieName&"_ShowMsg")=true
       Session(CookieName&"_MsgText")="<font color=""#ff0000"">"&PlugName&"</font> 插件安装完成并且移动到 <a href=""ConContent.asp?Fmenu=Skins&Smenu=Plugins"" style=""color:#0000ff"">已装插件管理</a>"
       Response.Redirect("ConContent.asp?Fmenu=Skins&Smenu=PluginsInstall")
     else
       session(CookieName&"_ShowMsg")=True
       session(CookieName&"_MsgText")="安装插件时发生错误,错误代码: <font color=""#ff0000"">"&PluginsXML.getError&"</font>"
       PluginsXML.CloseXml()
       Response.Redirect("ConContent.asp?Fmenu=Skins&Smenu=PluginsInstall")
   end if
end sub

'====================写入插件ASP代码======================
sub FixPlugins(fixType)
  dim PlugName,PlugTitle,PlugType,PlugSortID,PlugHtmlCode,PlugPluginPath,PlugASPCode,PlugSettingXML
  dim MainItem,SubItem,CreateTableSQL,UpdateTableSQL,Temp,ASPCode
   ASPCode=""
   set Bmodule=conn.execute("select InstallFolder from blog_module where IsInstall=true")
   if Bmodule.eof or Bmodule.bof then
	       If CheckObjInstalled("ADODB.Stream") Then SaveToFile "","plugins.asp" 
   else
    do until Bmodule.eof
		   PluginsFolder=Bmodule("InstallFolder")
		   set PluginsXML=New PXML
		   PluginsXML.XmlPath="Plugins/"&PluginsFolder&"/install.xml"
		   PluginsXML.open
		   if PluginsXML.getError=0 then
		            PlugName = PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginName")
		            PlugASPCode = PluginsXML.SelectXmlNodeText("PluginInstall/main/PluginASPCode")
					Temp="&lt;%'---- ASPCode For "&PlugName&" ----%&gt;"
				    Temp=Replace(Temp,"&lt;","<")
				    Temp=Replace(Temp,"&gt;",">")
		   			ASPCode=ASPCode & Temp & Chr(10)
		            if len(PlugASPCode)>0 then
			            ASPCode=ASPCode & PlugASPCode & Chr(10)
					end if
		            SubItemLen = int(PluginsXML.GetXmlNodeLength("PluginInstall/SubItem/item"))
		            for tempI=0 to SubItemLen-1 
		               PlugType = PluginsXML.SelectXmlNodeItemText("PluginInstall/SubItem/item/PluginType",tempI)
		               if not PlugType="function" then
				            PlugASPCode = PluginsXML.SelectXmlNodeItemText("PluginInstall/SubItem/item/PluginASPCode",tempI)
							Temp="&lt;%'---- ASPCode For "&PlugName&"SubItem"&(tempI+1)&" ----%&gt;"
						    Temp=Replace(Temp,"&lt;","<")
						    Temp=Replace(Temp,"&gt;",">")
				   			ASPCode=ASPCode & Temp & Chr(10)
				            if len(PlugASPCode)>0 then
					            ASPCode=ASPCode & PlugASPCode & Chr(10)
			               end if
		               end if
		            next
		       PluginsXML.CloseXml()
		       
	       log_module(2)
	       If CheckObjInstalled("ADODB.Stream") Then SaveToFile ASPCode,"plugins.asp" 
	     else
	       if cbool(fixType) then
		       session(CookieName&"_ShowMsg")=True
		       session(CookieName&"_MsgText")="安装插件时发生错误,错误代码: <font color=""#ff0000"">"&PluginsXML.getError&"</font>"
		       PluginsXML.CloseXml()
		       Response.Redirect("ConContent.asp?Fmenu=Skins&Smenu=Plugins")
	       end if
	   end if
	   Bmodule.movenext
   loop
  end if
       if cbool(fixType) then
	       session(CookieName&"_ShowMsg")=true
	       Session(CookieName&"_MsgText")="插件已经修复完成!"
	       Response.Redirect("ConContent.asp?Fmenu=Skins&Smenu=Plugins")
       end if
end sub

'====================删除日志方法======================
function DeleteLog(LogID)
  dim logDB,comCount,post_CateID
     set logDB=conn.execute("select log_IsDraft,log_CateID from blog_content where log_ID="&LogID)
     if logDB.eof and logDB.bof then
         DeleteLog=0
     else
	     comCount=Conn.ExeCute("select count(blog_ID) FROM blog_Comment WHERE blog_ID="&LogID)(0)
	     response.write comCount&"<br/>"
	     Conn.ExeCute("update blog_Info set blog_CommNums=blog_CommNums-"&comCount)  
	     if logDB("log_IsDraft")=false then
			 Conn.ExeCute("UPDATE blog_Info SET blog_LogNums=blog_LogNums-1")
			 Conn.ExeCute("UPDATE blog_Category SET cate_count=cate_count-1 where cate_ID="&logDB("log_CateID"))
		 end if
	     Conn.ExeCute("DELETE * FROM blog_Comment WHERE blog_ID="&LogID)
	     Conn.ExeCute("DELETE * FROM blog_Content WHERE log_ID="&LogID)
	     DeleteFiles Server.MapPath("post/"&DelID&".asp")
         DeleteLog=1
	 end if
end function


function bc(t,s)
 dim tl,sl,i
 bc=false
 sl=len(s)
 tl=len(t)
 if tl< sl then bc=true:exit function
 for i=1 to sl
  if mid(t,i,1)<>mid(s,i,1) then bc=true:exit function
 next
end function

sub saveFilterKey(keyword)
         dim tempStr,saveXML,keywords,i
         tempStr = "<?xml version=""1.0"" ?><spam>"
	     keywords=split(keyword,", ")
	     for i=0 to ubound(keywords)
		     if len(trim(keywords(i)))>0 then tempStr = tempStr & "<key><![CDATA[" & trim(keywords(i)) & "]]></key>"
	     next
		 tempStr = tempStr & "</spam>"
		 saveXML = SaveToFile (tempStr,"spam.xml")
		 session(CookieName&"_ShowMsg")=true
		 if saveXML(0)<>0 then session(CookieName&"_MsgText")=saveXML(1) else session(CookieName&"_MsgText")="过滤列表保存成功."
 		 Response.Redirect("ConContent.asp?Fmenu=Comment&Smenu=spam")
end sub

sub saveReFilterKey()
         dim tempStr,saveXML,keywords,i
         tempStr = "<?xml version=""1.0"" ?><reg>"
         tempStr = tempStr & "<text><![CDATA[" & Request.form("testArea") & "]]></text>"
		 
	     for i=1 to Request.form("des").count
		     tempStr = tempStr & "<key des=""" & htmlE(Request.form("des").item(i)) & """ re=""" & htmlE(Request.form("re").item(i)) & """ times=""" & Request.form("times").item(i) & """/>"
	     next
	     
		 tempStr = tempStr & "</reg>"

	     saveXML = SaveToFile (tempStr,"reg.xml")
		 
		 session(CookieName&"_ShowMsg")=true
		 if saveXML(0)<>0 then session(CookieName&"_MsgText")=saveXML(1) else session(CookieName&"_MsgText")="过滤列表保存成功."
 		 Response.Redirect("ConContent.asp?Fmenu=Comment&Smenu=reg")
end sub

function htmlE(reString)
	Dim Str:Str=reString
	If Not IsNull(Str) Then
 		Str = Replace(Str, "&", "&amp;")
   		Str = Replace(Str, ">", "&gt;")
		Str = Replace(Str, "<", "&lt;")
    	Str = Replace(Str, """", "&quot;")
    	Str = Replace(Str, "\", "\\")
		htmlE = Str
	end if
end function 

'-----------获取分类标题-------------------
function categoryTitle()
   dim Fmenu,Smenu,cTitle
   Fmenu = Request.QueryString("Fmenu")
   Smenu = Request.QueryString("Smenu")
   set cTitle=Server.CreateObject("Scripting.Dictionary")
   cTitle.Add "General.Misc" , "站点基本设置 - 初始化数据"
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
   
   cTitle.Add "Status." , "服务器配置信息"
   
   cTitle.Add "welcome." , "欢迎使用PJBlog2"

   categoryTitle = cTitle(Fmenu & "." & Smenu)
   set cTitle = nothing
end function
%>