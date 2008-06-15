<%
'===========PBlog2 UBB代码转换代码==========
'      Author:PuterJam
'         Copryright PBlog2 
'         Update: 2005-12-29
'===========================================
Function UBBCode(ByVal strContent,DisSM,DisUBB,DisIMG,AutoURL,AutoKEY)
	If isEmpty(strContent) Or isNull(strContent) Then
        Exit Function
	Else
		Dim re, strMatchs, strMatch, rndID,tmpStr1,tmpStr2,tmpStr3,tmpStr4
		Set re=new RegExp
		re.IgnoreCase =True
		re.Global=True
		IF AutoURL=1 Then
			re.Pattern="([^=\]][\s]*?|^)(http|https|rstp|ftp|mms|ed2k)://([A-Za-z0-9\.\/=\?%\-_~`@':+!]*)"
			Set strMatchs=re.Execute(strContent)
			For Each strMatch in strMatchs
				tmpStr1=strMatch.SubMatches(0)
				tmpStr2=strMatch.SubMatches(1)
				tmpStr3=checkURL(strMatch.SubMatches(2))
				strContent=replace(strContent,strMatch.Value,tmpStr1&"<a href="""&tmpStr2&"://"&tmpStr3&""" target=""_blank"" rel=""external"">"&tmpStr2&"://"&tmpStr3&"</a>",1,-1,0)
			Next
			're.Pattern="(^|\s)(www\.\S+)"
			'strContent=re.Replace(strContent,"$1<a href=""http://$2"" target=""_blank"">$2</a>")
		End IF
		
		'防止xss注入
		strContent = replace(strContent,"expression","e&#173;xpression",1,-1,0)

		IF Not DisUBB=1 Then
			IF Not DisIMG=1 Then
          			re.Pattern="(\[img\])(.[^\]]*)\[\/img\]"
					Set strMatchs=re.Execute(strContent)
					For Each strMatch in strMatchs
						tmpStr1=(strMatch.SubMatches(1))
						strContent=replace(strContent,strMatch.Value,"<img src="""&tmpStr1&""" border=""0"" alt=""""/>",1,-1,0)
					Next
					
          			re.Pattern="\[img=(left|right|center|absmiddle|)\](.[^\]]*)(\[\/img\])"
					Set strMatchs=re.Execute(strContent)
					For Each strMatch in strMatchs
						tmpStr1=strMatch.SubMatches(0)
						tmpStr2=checkURL(strMatch.SubMatches(1))
						strContent=replace(strContent,strMatch.Value,"<img align="""&tmpStr1&""" src="""&tmpStr2&""" border=""0"" alt=""""/>",1,-1,0)
					Next
      			
          			re.Pattern="\[img=(\d*|),(\d*|)\](.[^\]]*)\[\/img\]"
					Set strMatchs=re.Execute(strContent)
					For Each strMatch in strMatchs
						tmpStr1=strMatch.SubMatches(0)
						tmpStr2=strMatch.SubMatches(1)
						tmpStr3=checkURL(strMatch.SubMatches(2))
						strContent=replace(strContent,strMatch.Value,"<img width="""&tmpStr1&""" height="""&tmpStr2&""" src="""&tmpStr3&""" border=""0"" alt=""""/>",1,-1,0)
					Next
					
          			re.Pattern="\[img=(\d*|),(\d*|),(left|right|center|absmiddle|)\](.[^\]]*)(\[\/img\])"
					Set strMatchs=re.Execute(strContent)
					For Each strMatch in strMatchs
						tmpStr1=strMatch.SubMatches(0)
						tmpStr2=strMatch.SubMatches(1)
						tmpStr3=strMatch.SubMatches(2)
						tmpStr4=checkURL(strMatch.SubMatches(3))
						strContent=replace(strContent,strMatch.Value,"<img width="""&tmpStr1&""" height="""&tmpStr2&""" align="""&tmpStr3&""" src="""&tmpStr4&""" border=""0"" alt=""""/>",1,-1,0)
					Next
				else
          			re.Pattern="(\[img\])(.[^\]]*)\[\/img\]"
					Set strMatchs=re.Execute(strContent)
					For Each strMatch in strMatchs
						tmpStr1=checkURL(strMatch.SubMatches(1))
						strContent=replace(strContent,strMatch.Value,"<a href="""&tmpStr1&""" target=""_blank"" title="""&tmpStr1&"""><img src=""images/image.gif"" alt="""" style=""margin:0px 2px -3px 0px"" border=""0""/>查看图片</a>",1,-1,0)
					Next
					
          			re.Pattern="\[img=(left|right|center|absmiddle|)\](.[^\]]*)(\[\/img\])"
					Set strMatchs=re.Execute(strContent)
					For Each strMatch in strMatchs
						tmpStr1=strMatch.SubMatches(0)
						tmpStr2=checkURL(strMatch.SubMatches(1))
						strContent=replace(strContent,strMatch.Value,"<a href="""&tmpStr2&""" target=""_blank"" title="""&tmpStr2&"""><img src=""images/image.gif"" alt="""" style=""margin:0px 2px -3px 0px"" border=""0""/>查看图片</a>",1,-1,0)
					Next
      			
          			re.Pattern="\[img=(\d*|),(\d*|)\](.[^\]]*)\[\/img\]"
					Set strMatchs=re.Execute(strContent)
					For Each strMatch in strMatchs
						tmpStr1=strMatch.SubMatches(0)
						tmpStr2=strMatch.SubMatches(1)
						tmpStr3=checkURL(strMatch.SubMatches(2))
						strContent=replace(strContent,strMatch.Value,"<a href="""&tmpStr3&""" target=""_blank"" title="""&tmpStr3&"""><img src=""images/image.gif"" alt="""" style=""margin:0px 2px -3px 0px"" border=""0""/>查看图片</a>",1,-1,0)
					Next
					
          			re.Pattern="\[img=(\d*|),(\d*|),(left|right|center|absmiddle|)\](.[^\]]*)(\[\/img\])"
					Set strMatchs=re.Execute(strContent)
					For Each strMatch in strMatchs
						tmpStr1=strMatch.SubMatches(0)
						tmpStr2=strMatch.SubMatches(1)
						tmpStr3=strMatch.SubMatches(2)
						tmpStr4=checkURL(strMatch.SubMatches(3))
						strContent=replace(strContent,strMatch.Value,"<a href="""&tmpStr4&""" target=""_blank"" title="""&tmpStr4&"""><img src=""images/image.gif"" alt="""" style=""margin:0px 2px -3px 0px"" border=""0""/>查看图片</a>",1,-1,0)
					Next				
			End IF
			
'-----------多媒体标签----------------
		  re.Pattern="\[(swf|wma|wmv|rm|ra|qt)(=\d*?|)(,\d*?|)\]([^<>]*?)\[\/(swf|wma|wmv|rm|ra|qt)\]"
	      Set strMatchs=re.Execute(strContent)
	      dim strType,strWidth,strHeight,strSRC,TitleText
			For Each strMatch in strMatchs
				RAndomize
				  strType=strMatch.SubMatches(0)
				  if strType="swf" then
				   TitleText="<img src=""images/flash.gif"" alt="""" style=""margin:0px 2px -3px 0px"" border=""0""/>Flash动画"
				  elseif strType="wma" then
				   TitleText="<img src=""images/music.gif"" alt="""" style=""margin:0px 2px -3px 0px"" border=""0""/>播放音频文件"
				  elseif strType="wmv" then
				   TitleText="<img src=""images/mediaplayer.gif"" alt="""" style=""margin:0px 2px -3px 0px"" border=""0""/>播放视频文件"		 
				  elseif strType="rm" then
				   TitleText="<img src=""images/realplayer.gif"" alt="""" style=""margin:0px 2px -3px 0px"" border=""0""/>播放real视频流文件"		 
				  elseif strType="ra" then
				   TitleText="<img src=""images/realplayer.gif"" alt="""" style=""margin:0px 2px -3px 0px"" border=""0""/>播放real音频流文件"		 
				  elseif strType="qt" then
				   TitleText="<img src=""images/mediaplayer.gif"" alt="""" style=""margin:0px 2px -3px 0px"" border=""0""/>播放mov视频文件"		 
				  end if
	              strWidth=strMatch.SubMatches(1)
	              strHeight=strMatch.SubMatches(2)
	              if (len(strWidth)=0) then 
	                strWidth="400"
	                else
	               strWidth=right(strWidth,(len(strWidth)-1))
	              end if
	              if (len(strHeight)=0) then
	               strHeight="300"
	                else
	              strHeight=right(strHeight,(len(strHeight)-1))
	              end if
	              strSRC=checkURL(strMatch.SubMatches(3))
				rndID="temp"&Int(100000 * Rnd)
				strContent= Replace(strContent,strMatch.Value,"<div class=""UBBPanel""><div class=""UBBTitle"">"&TitleText&"</div><div class=""UBBContent""><a id="""+rndID+"_href"" href=""javascript:MediaShow('"+strType+"','"+rndID+"','"+strSRC+"','"+strWidth+"','"+strHeight+"')""><img name="""+rndID+"_img"" src=""images/mm_snd.gif"" style=""margin:0px 3px -2px 0px"" border=""0"" alt=""""/><span id="""+rndID+"_text"">在线播放</span></a><div id="""+rndID+"""></div></div></div>")
			Next
			Set strMatchs=nothing
		
			re.Pattern="(\[mid\])(.[^\]]*)\[\/mid\]"
			strContent= re.Replace(strContent,"<embed src=""$2"" height=""45"" width=""314"" autostart=""0""></embed>")	
'-----------常规标签----------------
			re.Pattern = "\[url=(.[^\]]*)\](.[^\[]*)\[\/url]" 
			Set strMatchs=re.Execute(strContent)
			For Each strMatch in strMatchs
				tmpStr1=checkURL(strMatch.SubMatches(0))
				tmpStr2=strMatch.SubMatches(1)
				strContent=replace(strContent,strMatch.Value,"<a target=""_blank"" href="""&tmpStr1&""" rel=""external"">"&tmpStr2&"</a>",1,-1,0)
			Next

			re.Pattern = "\[url](.[^\[]*)\[\/url]"
			Set strMatchs=re.Execute(strContent)
			For Each strMatch in strMatchs
				tmpStr1=checkURL(strMatch.SubMatches(0))
				strContent=replace(strContent,strMatch.Value,"<a target=""_blank"" href="""&tmpStr1&""" rel=""external"">"&tmpStr1&"</a>",1,-1,0)
			Next
						
			re.Pattern = "\[ed2k=([^\r]*?)\]([^\r]*?)\[\/ed2k]"
			Set strMatchs=re.Execute(strContent)
			For Each strMatch in strMatchs
				tmpStr1=checkURL(strMatch.SubMatches(0))
				tmpStr2=strMatch.SubMatches(1)
				strContent=replace(strContent,strMatch.Value,"<img border="""" src=""images/ed2k.gif"" alt=""""/><a target=""_blank"" href="""&tmpStr1&""">"&tmpStr2&"</a>",1,-1,0)
			Next
			
			re.Pattern = "\[ed2k]([^\r]*?)\[\/ed2k]"
			Set strMatchs=re.Execute(strContent)
			For Each strMatch in strMatchs
				tmpStr1=checkURL(strMatch.SubMatches(0))
				strContent=replace(strContent,strMatch.Value,"<img border="""" src=""images/ed2k.gif"" alt=""""/><a target=""_blank"" href="""&tmpStr1&""">"&tmpStr1&"</a>",1,-1,0)
			Next
						
			re.Pattern = "\[email=(.[^\]]*)\](.[^\[]*)\[\/email]"
			Set strMatchs=re.Execute(strContent)
			For Each strMatch in strMatchs
				tmpStr1=checkURL(strMatch.SubMatches(0))
				tmpStr2=strMatch.SubMatches(1)
				strContent=replace(strContent,strMatch.Value,"<a href=""mailto:"&tmpStr1&""">"&tmpStr2&"</a>",1,-1,0)
			Next

			
			re.Pattern = "\[email](.[^\[]*)\[\/email]"
			Set strMatchs=re.Execute(strContent)
			For Each strMatch in strMatchs
				tmpStr1=checkURL(strMatch.SubMatches(0))
				strContent=replace(strContent,strMatch.Value,"<a href=""mailto:"&tmpStr1&""">"&tmpStr1&"</a>",1,-1,0)
			Next

'-----------字体格式----------------
    		re.Pattern="\[align=(\w{4,6})\]([^\r]*?)\[\/align\]"
			strContent=re.Replace(strContent,"<div align=""$1"">$2</div>")
  		    re.Pattern="\[color=(#\w{3,10}|\w{3,10})\]([^\r]*?)\[\/color\]"
			strContent=re.Replace(strContent,"<span style=""color:$1"">$2</span>")
    		re.Pattern="\[size=(\d{1,2})\]([^\r]*?)\[\/size\]"
			strContent=re.Replace(strContent,"<span style=""font-size:$1pt"">$2</span>")
    		re.Pattern="\[font=([^\r]*?)\]([^\r]*?)\[\/font\]"
			strContent=re.Replace(strContent,"<span style=""font-family:$1"">$2</span>")
			re.Pattern="\[b\]([^\r]*?)\[\/b\]"
			strContent=re.Replace(strContent,"<strong>$1</strong>")
    		re.Pattern="\[i\]([^\r]*?)\[\/i\]"
			strContent=re.Replace(strContent,"<i>$1</i>")
    		re.Pattern="\[u\]([^\r]*?)\[\/u\]"
			strContent=re.Replace(strContent,"<u>$1</u>")
    		re.Pattern="\[s\]([^\r]*?)\[\/s\]"
			strContent=re.Replace(strContent,"<s>$1</s>")
    		re.Pattern="\[sup\]([^\r]*?)\[\/sup\]"
			strContent=re.Replace(strContent,"<sup>$1</sup>")
    		re.Pattern="\[sub\]([^\r]*?)\[\/sub\]"
			strContent=re.Replace(strContent,"<sub>$1</sub>")
    		re.Pattern="\[fly\]([^\r]*?)\[\/fly\]"
			strContent=re.Replace(strContent,"<marquee width=""90%"" behavior=""alternate"" scrollamount=""3"">$1</marquee>")
  
		
'-----------特殊标签----------------
			re.Pattern = "\[down=(.[^\]]*)\](.[^\[]*)\[\/down]"
			Set strMatchs=re.Execute(strContent)
			For Each strMatch in strMatchs
				tmpStr1=checkURL(strMatch.SubMatches(0))
				tmpStr2=strMatch.SubMatches(1)
				strContent=replace(strContent,strMatch.Value,"<img src=""images/download.gif"" alt=""下载文件"" style=""margin:0px 2px -4px 0px""/> <a href="""&tmpStr1&""" target=""_blank"">"&tmpStr2&"</a>",1,-1,0)
			Next		
		
			re.Pattern = "\[down\](.[^\[]*)\[\/down]"
			Set strMatchs=re.Execute(strContent)
			For Each strMatch in strMatchs
				tmpStr1=checkURL(strMatch.SubMatches(0))
				strContent=replace(strContent,strMatch.Value,"<img src=""images/download.gif"" alt=""下载文件"" style=""margin:0px 2px -4px 0px""/> <a href="""&tmpStr1&""" target=""_blank"">下载此文件</a>",1,-1,0)
			Next

			re.Pattern = "\[mDown=(.[^\]]*)\](.[^\[]*)\[\/mDown]"
			Set strMatchs=re.Execute(strContent)
			For Each strMatch in strMatchs
				tmpStr1=checkURL(strMatch.SubMatches(0))
				tmpStr2=strMatch.SubMatches(1)
				if len(memName)>0 then
					strContent=replace(strContent,strMatch.Value,"<img src=""images/download.gif"" alt=""下载文件"" style=""margin:0px 2px -4px 0px""/> <a href="""&tmpStr1&""" target=""_blank"">"&tmpStr2&"</a>",1,-1,0)
				 else
					strContent=replace(strContent,strMatch.Value,"<img src=""images/download.gif"" alt=""只允许会员下载"" style=""margin:0px 2px -4px 0px""/> 该文件只允许会员下载! <a href=""login.asp"">登录</a> | <a href=""register.asp"">注册</a>",1,-1,0)
				end if
			Next

			re.Pattern = "\[mDown\](.[^\[]*)\[\/mDown]"
			Set strMatchs=re.Execute(strContent)
			For Each strMatch in strMatchs
				tmpStr1=checkURL(strMatch.SubMatches(0))
				if len(memName)>0 then
					strContent=replace(strContent,strMatch.Value,"<img src=""images/download.gif"" alt=""下载文件"" style=""margin:0px 2px -4px 0px""/> <a href="""&tmpStr1&""" target=""_blank"">下载此文件</a>",1,-1,0)
				 else
					strContent=replace(strContent,strMatch.Value,"<img src=""images/download.gif"" alt=""只允许会员下载"" style=""margin:0px 2px -4px 0px""/> 该文件只允许会员下载! <a href=""login.asp"">登录</a> | <a href=""register.asp"">注册</a>",1,-1,0)
				end if
			Next

		'CC Video 支持
		re.Pattern="\[cc\](.*?)\[\/cc\]"
		strContent= re.Replace(strContent,"<embed src=""http://union.bokecc.com/$1"" width=""438"" height=""387"" type=""application/x-shockwave-flash""></embed>")
			re.Pattern="\[code\](.*?)\[\/code\]"
			strContent= re.Replace(strContent,"<div class=""UBBPanel""><div class=""UBBTitle""><img src=""images/code.gif"" style=""margin:0px 2px -3px 0px"" alt=""程序代码""/> 程序代码</div><div class=""UBBContent"">$1</div></div>")

			re.Pattern="\[quote\](.*?)\[\/quote\]"
			strContent= re.Replace(strContent,"<div class=""UBBPanel""><div class=""UBBTitle""><img src=""images/quote.gif"" style=""margin:0px 2px -3px 0px"" alt=""引用内容""/> 引用内容</div><div class=""UBBContent"">$1</div></div>")
			re.Pattern="\[quote=(.[^\]]*)\](.*?)\[\/quote\]"
			strContent= re.Replace(strContent,"<div class=""UBBPanel""><div class=""UBBTitle""><img src=""images/quote.gif"" style=""margin:0px 2px -3px 0px"" alt=""引用来自 $1""/> 引用来自 $1</div><div class=""UBBContent"">$2</div></div>")

			re.Pattern="\[hidden\](.*?)\[\/hidden\]"
			if len(memName)>0 then		
				strContent= re.Replace(strContent,"<div class=""UBBPanel""><div class=""UBBTitle""><img src=""images/quote.gif"" style=""margin:0px 2px -3px 0px"" alt=""显示被隐藏内容""/> 显示被隐藏内容</div><div class=""UBBContent"">$1</div></div>")
			 else
				strContent= re.Replace(strContent,"<div class=""UBBPanel""><div class=""UBBTitle""><img src=""images/quote.gif"" style=""margin:0px 2px -3px 0px"" alt=""隐藏内容""/> 隐藏内容</div><div class=""UBBContent"">该内容已经被作者隐藏,只有会员才允许查阅 <a href=""login.asp"">登录</a> | <a href=""register.asp"">注册</a></div></div>")
			end if
			
			re.Pattern="\[hidden=(.[^\]]*)\](.*?)\[\/hidden\]"
			if len(memName)>0 then		
				strContent= re.Replace(strContent,"<div class=""UBBPanel""><div class=""UBBTitle""><img src=""images/quote.gif"" style=""margin:0px 2px -3px 0px"" alt=""显示被隐藏内容 $1""/> 显示被隐藏内容来自 $1</div><div class=""UBBContent"">$2</div></div>")
			 else
				strContent= re.Replace(strContent,"<div class=""UBBPanel""><div class=""UBBTitle""><img src=""images/quote.gif"" style=""margin:0px 2px -3px 0px"" alt=""隐藏内容 $1""/> 隐藏内容</div><div class=""UBBContent"">该内容已经被作者隐藏,只有会员才允许查阅 <a href=""login.asp"">登录</a> | <a href=""register.asp"">注册</a></div></div>")
			end if
		
			IF Not DisIMG=1 Then
				re.Pattern="\[html\](.*?)\[\/html\]"
				Set strMatchs=re.Execute(strContent)
				For Each strMatch in strMatchs
					RAndomize
					rndID="temp"&Int(100000 * Rnd)
					strContent=Replace(strContent,strMatch.Value,"<div class=""UBBPanel""><div class=""UBBTitle""><img src=""images/html.gif"" style=""margin:0px 2px -3px 0px""> HTML代码</div><div class=""UBBContent""><TEXTAREA rows=""8"" id="""&rndID&""">"&UBBFilter(HTMLDecode(strMatch.SubMatches(0)))& "</TEXTAREA><br/><INPUT onclick=""runEx('"&rndID&"')""  type=""button"" value=""运行此代码""/> <INPUT onclick=""doCopy('"&rndID&"')""  type=""button"" value=""复制此代码""/><br/> [Ctrl+A 全部选择 提示：你可先修改部分代码，再按运行]</div></div>",1,-1,0)
				Next
				Set strMatchs=nothing
			end if
			
'-----------List标签----------------
			strContent = Replace(strContent,"[list]","<ul>")
			re.Pattern = "\[list=(.[^\]]*)\]"
			strContent = re.Replace(strContent,"<ul style=""list-style-type:$1"">")
			re.Pattern = "\[\*\](.[^\[]*)(\n|)"
			strContent = re.Replace(strContent,"<li>$1</li>")
			strContent = Replace(strContent,"[/list]","</ul>")
	End IF

'-----------表情图标----------------
		IF Not DisSM=1 Then
			dim log_Smilies,log_SmiliesContent
			For Each log_Smilies IN Arr_Smilies
				log_SmiliesContent=Split(log_Smilies,"|")
				strContent=Replace(strContent,log_SmiliesContent(2)," <img src=""images/smilies/"&log_SmiliesContent(1)&""" border=""0"" style=""margin:0px 0px -2px 0px"" alt=""""/>")
			Next
		End IF
		
'-----------关键词识别----------------
		IF AutoKEY=1 Then
			dim log_Keywords,log_KeywordsContent
			For Each log_Keywords IN Arr_Keywords
				log_KeywordsContent=Split(log_Keywords,"$|$")
				IF log_KeywordsContent(3)<>"None" Then
					strContent=Replace(strContent,log_KeywordsContent(1),"<a href="""&log_KeywordsContent(2)&""" target=""_blank""><img src=""images/keywords/"&log_KeywordsContent(3)&""" border=""0"" alt=""""/> "&log_KeywordsContent(1)&"</a>")
				Else
					strContent=Replace(strContent,log_KeywordsContent(1),"<a href="""&log_KeywordsContent(2)&""" target=""_blank"">"&log_KeywordsContent(1)&"</a>")
				End IF
			Next
		End IF

		Set re=Nothing

		UBBCode=strContent
	End IF
End Function
%>