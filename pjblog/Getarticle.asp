<!--#include file="BlogCommon.asp" -->
<!--#include file="common/function.asp" -->
<!--#include file="common/library.asp" -->
<!--#include file="common/upfile.asp" -->
<!--#include file="common/cache.asp" -->
<!--#include file="common/checkUser.asp" -->
<%
Response.expires=-1 
Response.AddHeader"pragma","no-cache" 
Response.AddHeader"cache-control","no-store" 
Session.CodePage=65001
Dim id
id=request("id")
blog_postFile = Cint(Checkxss(request.QueryString("blog_postFile")))
if id<>"" and  isnumeric(id) then   
  Dim wbc_tag,Rs,i,RsT,OutPut,i2,i3,i4,str,ifMore,i2_2,total_rela,page,pagestr,thispage
  Set Rs=conn.execute("Select * from blog_Content where log_ID="&id&"")
  if request("page")<>"" and isnumeric(request("page"))  then
     page=Cint(request("page"))
	 if page<1 then page=1
  Else
     page=1
  End if
  thispage=page
  OutPut=""
  total_rela=0
  i2_2=10 '每页显示No.
  i3=(page-1)*i2_2
  if request("act")="more" then i2_2=-1
  ifMore=false
  if not rs.eof then
     wbc_tag=rs("log_tag")
     if wbc_tag<>"" then
        wbc_tag=split(wbc_tag,"}")
		i=0
		i2=0
		'find total
		DO until i>Ubound(wbc_tag)
           Set RsT=conn.execute("Select log_Title,log_id,log_ViewNums from  blog_Content where log_tag like '%"&wbc_tag(i)&"}%' and log_ID<>"&id&" order by log_PostTime desc")
		   if not RsT.eof then
		      Do until Rst.eof
                  str=split(RsT(0),"(*##*)")(0)
				  If instr(OutPut,str)=0 and wbc_tag(i)<>"" then
		             total_rela=total_rela+1
				  End if
				  Rst.movenext
			  Loop
		   End if
		   Set RsT=nothing
		   i=i+1
		Loop
		'pages
		if int((total_rela/i2_2))< total_rela/i2_2 then
		   page=(total_rela/i2_2)+1
		 Else
		   page=int(total_rela/i2_2)
		 end if		
		 pagestr=""
		 i=1
		 do until i>page
		   if thispage=i then 
		      pagestr=pagestr&"<strong style='text-decoration:none' title='当前页'>["&i&"]</strong>"
		   Elseif abs(thispage-i)<5 then
		      pagestr=pagestr&"<a style='cursor:pointer'  title='转到第"&i&"页' onclick=check('Getarticle.asp?id="&id&"&page="&i&"&blog_postFile="&blog_postFile&"','wbc_tag','wbc_tag') >["&i&"]</a>"
		   end if
		   i=i+1
		 Loop
		 if page>1 then pagestr=" <div style='margin-top:10px;width:350px;'><strong>分页:</strong> "&pagestr&""
		'show results
		i=0
		i4=0
		Dim urlLink
		DO until i>Ubound(wbc_tag)
           Set RsT=conn.execute("Select log_Title,log_id,log_ViewNums from  blog_Content where log_tag like '%"&wbc_tag(i)&"}%' and log_ID<>"&id&" order by log_PostTime desc")
		   if not RsT.eof then
		      Do until Rst.eof  or i4=i2_2
                  str=split(RsT(0),"(*##*)")(0)
				  If instr(OutPut,str)=0 and wbc_tag(i)<>"" then
				     if i2>(i3-1) then
					    i4=i4+1
						if blog_postFile = 2 then
							urlLink = Alias(RsT(1))
						else 
							urlLink = "article.asp?id="&RsT(1)
						end if
					    OutPut=OutPut&"   <li><a href='"&urlLink&"'>"&str&"["&RsT(2)&"]</a></li>"
					 end if
					 i2=i2+1
				  End if
				  Rst.movenext
			  Loop
              if not RsT.eof then Ifmore=true
		   End if
		   Set RsT=nothing
		   i=i+1
		Loop
	 End if
  End if
  output=output&pagestr
  If Ifmore or thispage<>1 then
     OutPut=OutPut&"<br/><strong>模式:</strong> <a style='cursor:pointer'  onclick=check('Getarticle.asp?id="&id&"&act=more&blog_postFile="&blog_postFile&"','wbc_tag','wbc_tag') >全部显示[共"&total_rela&"个相关文章]</a></div>"
  elseif i2>(i2_2-1)  then
     OutPut=OutPut&"<div style='margin-top:10px;width:350px;'><strong>模式:</strong> <a style='cursor:pointer'  onclick=check('Getarticle.asp?id="&id&"&blog_postFile="&blog_postFile&"','wbc_tag','wbc_tag') >分页显示[共"&total_rela&"个相关文章]</a></div>"
  end if 
  OutPut=replace(OutPut,chr(39),chr(34))
  response.Write OutPut
  Set rs=nothing
end if
'*************************************
'自定义路径 by evio
'*************************************
Function Alias(id)
    dim cname,ccate,chtml,ccateID,ccateExec,cnames,ctype,cc
	set cc=conn.execute("select top 1 log_CateID,log_cname,log_ctype from blog_Content where log_ID="&id)
	
	ccateID = cc(0)
	cname = cc(1)
	ctype = cc(2)

	set ccateExec=conn.execute("select Cate_Part from blog_Category where cate_ID="&ccateID)

	If not ccateExec.EOF and not ccateExec.bof Then
		ccate = ccateExec(0).value
	end if
	
	if ccate="" or ccate=empty or ccate=null or len(ccate)=0 then
		ccate="article/"
	else
		ccate="article/"&ccate&"/"
	end if
	
	if cname="" or cname=empty or cname=null or len(cname)=0 then
		cnames=trim(id)
	else
		cnames=cname
	end if
	
	if ctype="0" then
		chtml="htm"
	elseif ctype="1" then
		chtml="html"
	else
		chtml="asp"
	end if
	
	chtml="."&chtml
	set ccateExec = nothing
	set cc = nothing
	Alias=ccate&cnames&chtml
End Function
%>