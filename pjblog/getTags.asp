<!--#include file="const.asp" -->
<!--#include file="conn.asp" -->
<!--#include file="common/function.asp" -->
<!--#include file="common/library.asp" -->
<!--#include file="common/cache.asp" -->
<%
tags(2)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="UTF-8">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
	<meta http-equiv="Content-Language" content="UTF-8" />
	<title>插入已经使用的Tag</title>
	<link rel="icon" href="favicon.ico" type="image/x-icon" />
	<link rel="shortcut icon" href="favicon.ico" type="image/x-icon" />
	<style>
	 body{margin:0px;background:#F1F1E3;border-width:0px}
	 #Top{background:#E3E3C7;border-bottom:1px solid #D5D59D;padding:8px;color:#737357;font-size:18px}
	 #Mid{height:250px;overflow:auto}
	 #Bottom{background:#E3E3C7;border-top:1px solid #D5D59D;padding:8px;color:#737357;text-align:right}
	 input{border:1px solid #737357;color:#3B3B1F;background:#C7C78F;font-size:12px;}
	 .tagA:link,.tagA:visited{
	   display:block;
	   background:#D7D79F;
	   padding:4px;
	   font-size:12px;
	   color:#3B3B1F;
	   margin:4px;
	   border:1px solid #D7D79F;
	   text-decoration:none;
	 }
	 .tagA:hover{
	   background:#EFEFDA;
	   border:1px solid #D7D79F;
	 }
	</style>
	<script>
	   function addTag(tagName) {
	     if (opener) {
	       var getTagObj=opener.document.forms[0].tags
	       var tags
	       if (getTagObj.value.length>0) {
	         tags=getTagObj.value.split(",")
	         for (i=0;i<tags.length;i++){
	           if (tags[i].toLowerCase()==tagName.toLowerCase()) return 
	         }
	         getTagObj.value+=","+tagName
	       }
	       else{
	         getTagObj.value+=tagName
	       }
	     }
	   }
	</script>
</head>
	<body scroll="no">
		<div id="Top"><b>插入已经使用的Tag</b></div>
	     <div id="Mid">
           <%
Dim log_Tag, log_TagItem
For Each log_TagItem IN Arr_Tags
    log_Tag = Split(log_TagItem, "||")

%>
	       <a href="#" class="tagA" onClick="addTag('<%=log_Tag(1)%>')" title="插入<%=log_Tag(1)%>"><%=log_Tag(1)%> (<%=log_Tag(2)%>)</a>
	       <%Next%>
	     </div>
		<div id="Bottom">
		  <input type="button" value="关闭" onClick="window.close()"/>
		</div>
	</body>
</html>
