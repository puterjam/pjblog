<%
function GetIp()
	'����ͻ������˴������������Ӧ����ServerVariables("HTTP_X_FORWARDED_FOR")����
	dim clientIp : clientIp = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
	if clientIp = "" then
		'����ͻ���û�ô���Ӧ����Request.ServerVariables("REMOTE_ADDR")����
		clientIp = Request.ServerVariables("REMOTE_ADDR")
	end if
	GetIp = clientIp
end function
'QUERY_STRING
function paramAdd(strParam,strAdd,value)
	dim arrParam,intI,intJ
	arrParam=split(strParam,"&")
	strAdd=lcase(strAdd)	
	intJ=Ubound(arrParam)
	paramAdd=""
	for intI=0 to intJ
		if inStr(1,lcase(arrParam(intI)),strAdd&"=")<>1 then paramAdd=paramAdd&arrParam(intI)&"&"
	next
	paramAdd=paramAdd&strAdd&"="&value
end function
'JS����
function EncodeJs(byVal str)
	if isNull(str) then
		EncodeJs=""
		exit function
	end if
	str=replace(str,chr(10),"")
	str=replace(str,chr(13),"\n")
	str=replace(str,"\","\\")
	str=replace(str,"""","\""")
	str=replace(str,"'","\'")
	EncodeJs=str	
end function
'ͨ�÷�ҳʵ��
function pageDivide(byRef rsCount,byRef pageSize,byRef pageCount,byRef pageNo)
	if not rs.EOF then
		rsCount		= rs.RecordCount
		rs.PageSize = pageSize
		pageCount	= rs.PageCount
		pageNo		= Request.QueryString("page")
		if isNumeric(pageNo) then
			pageNo	= cint(pageNo)
		else
			pageNo	= 1
		end if
		if pageNo<1 then pageNo=1
		if pageNo>pageCount then pageNo=pageCount
		rs.AbsolutePage=pageNo
		pageDivide = true
	else
		rsCount	   = 0
		pageCount  = 0
		pageNo	   = 0
		pageDivide = false
	end if
End function
'ͨ�÷�ҳ��ʾ
Sub pageTurn(byVal rsCount,byVal pageSize,byVal pageCount,byVal pageNo)
	Response.Write "<script type=""text/javascript"">pageTurn("&rsCount&","&pageSize&","&pageCount&","&pageNo&","""&EncodeJs(strParam)&""");</script>"
End Sub

dim strParam,strDconn
strParam = Request.ServerVariables("QUERY_STRING")
strDconn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("UploadClass.mdb")
%>				


