<!--#include file="BlogCommon.asp" -->
<!--#include file="common/function.asp" -->
<!--#include file="common/library.asp" -->
<!--#include file="common/upfile.asp" -->
<!--#include file="common/cache.asp" -->
<!--#include file="common/checkUser.asp" -->
<%
'********************************************
'相关文章功能重构。
'create by hayden
'09.06.06
'********************************************
Response.expires=-1 
Response.AddHeader"pragma","no-cache" 
Response.AddHeader"cache-control","no-store" 
Session.CodePage=65001
Dim Rs,tagsql,related_tag,urlLink,OutPut
Dim id : id = CheckStr(request.querystring("id"))
Dim Perpage : Perpage = 10
CurPage = CheckStr(request.querystring("page"))
Dim act : act = CheckStr(request.querystring("act"))
Dim isrel : isrel = False
If Not IsInteger(CurPage) Then 
	CurPage = 1
Else
	CurPage= CInt(CurPage)
End If 
Dim totalPut,mpage

if IsInteger(id) Then getart gettag
If Not isrel Then response.write "无相关文章！"	'这里可以试着放上广告代码，在没有相关文章的情况下才显示。
Set conn=Nothing
'response.write SQLQueryNums

'获取相关文章TAG
Function gettag()
	gettag = ""
	Set Rs=conn.execute("Select log_tag from blog_Content where log_ID="&id&"")
	SQLQueryNums=SQLQueryNums+1
	if not rs.eof then
		related_tag=rs(0)
		if not isblank(related_tag) then
			gettag = "(log_tag like '%" & Replace(related_tag,"}{","}%' or log_tag like '%{")&"%') "
		End If
	End if
	Set rs=nothing
End Function 

'获取相关文章HTML
Function getart(ByVal strtagsql)
	If Len(strtagsql)=0 Then Exit Function 
	Dim RsT,i,tgthml
	set RsT=server.createobject("adodb.recordset")
	tagsql = "Select log_Title,log_id,log_ViewNums from blog_Content where "&strtagsql&" and log_ID<>"&id&" and log_IsDraft=False and log_CateID not  in (select  cate_id from blog_Category where cate_secret=True ) order by log_PostTime desc"
	RsT.open tagsql,conn,1,1
	SQLQueryNums=SQLQueryNums+1
	if RsT.eof Then
		Set RsT=Nothing
		Exit Function
	End If 
	If act<>"more" Then
		totalPut=RsT.recordcount
		RsT.move  (CurPage-1)*Perpage
		RsT.pagesize=Perpage '得到每页数
		mpage=RsT.pagecount     '得到总页数
		i = 0
	End If 
	do while not RsT.eof 
		if blog_postFile = 2 then
			urlLink = caload(RsT(1))
		else 
			urlLink = "article.asp?id="&RsT(1)
		end if
		tgthml=tgthml&"   <li><a href='"&urlLink&"'>"&RsT(0)&"["&RsT(2)&"]</a></li>"
		If act<>"more" Then 
			i=i+1
			if i>=Perpage then exit Do
		End If 
		Rst.movenext
	Loop
	'OutPut = "<ul>"&OutPut&"</ul>"
	If act<>"more" Then tgthml = tgthml & subpage(totalPut,mpage,CurPage,Perpage)
	If act<>"more" And CurPage>1 then
		tgthml=tgthml&"<br/><strong>模式:</strong> <a style='cursor:pointer'  onclick=""Related(" & id & ", 2, 'related_tag');"" >全部显示[共"&totalPut&"个相关文章]</a></div>"
	ElseIf  act="more" Then
		tgthml=tgthml&"<div style='margin-top:10px;width:350px;'><strong>模式:</strong> <a style='cursor:pointer'  onclick=""Related(" & id & ", 2, 'related_tag');"" >分页显示[共"&totalPut&"个相关文章]</a></div>"
	end if 
	tgthml=replace(tgthml,chr(39),chr(34))
	response.Write tgthml
	isrel = True

	Set RsT=Nothing
End Function 

'参数依次　记录总数，总页数，当前页数，每页记录数
Function subpage(ByVal totalPut,ByVal mpage,ByVal CurPage,ByVal Perpage)
	Dim tphtml,pi
	tphtml = "<strong>分页:</strong>"
	If CurPage-5>1 Then 
		tphtml = tphtml & "<a style=""cursor:pointer""  title=""转到第1页"" onclick=""Related(" & id & ", 2, 'related_tag');"" >[1]</a>..."
	End If 
	For pi=CurPage-5 To CurPage+5
		If pi>0 And pi< mpage+1 Then 
			If pi = Int(CurPage) Then
				tphtml = tphtml & "<strong style=""text-decoration:none"" title=""当前页"">["&pi&"]</strong>"
			Else
				tphtml = tphtml & "<a style=""cursor:pointer""  title=""转到第"&pi&"页"" onclick=""Related(" & id & ", 2, 'related_tag');"" >["&pi&"]</a>"
			End If 
		End If 
	Next
	If CurPage+5<Int(mpage) Then 
		tphtml = tphtml & "...<a style=""cursor:pointer""  title=""转到第"&mpage&"页"" onclick=""Related(" & id & ", 2, 'related_tag');"" >["&mpage&"]</a>"
	End If 
	subpage = tphtml
end Function
%>