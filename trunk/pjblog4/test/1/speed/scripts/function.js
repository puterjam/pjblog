//Encode
function EncodeUrl(str){	
	var strUrl=str.replace(/\r/g,"%0A");
	strUrl=strUrl.replace(/\n/g,"%0D");
	strUrl=strUrl.replace(/\"/g,"%22");
	strUrl=strUrl.replace(/\'/g,"%27");
	return strUrl;
}
//QUERY_STRING
function paramAdd(strParam,strAdd,value){
	strAdd=strAdd.toLowerCase();
	var arrParam=strParam.split("&");
	var j=arrParam.length;
	if(strParam=="")j=-1;
	var paramAdd="";
	for(var i=0;i<j;i++){
		if(arrParam[i].toLowerCase().indexOf(strAdd+"=")!=0){
			paramAdd+=arrParam[i]+"&";
		}
	}
	return EncodeUrl(paramAdd+strAdd+"="+value);
}
function pageTurn(rsCount,pageSize,pageCount,pageNo,strParam){
	if(pageCount>1){
	with(document){
		write("<table width=\"100%\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\"><tr><td>页次：");
		write(pageNo+"/"+pageCount+" 每页："+pageSize+" 记录数："+rsCount);
		write("</td><td align=\"right\">");
		var intFrom=pageNo-5;
		if(intFrom<1)intFrom=1;
		var intTo=intFrom+10;
		if(intTo>pageCount){
			intTo=pageCount;
			if(pageCount>10)
				intFrom=intTo-10;
			else
				intFrom=1;
		}
		write("分页："	);
		if(intFrom>1)
			write("<a href=\"?"+paramAdd(strParam,"page",1)+"\"><<</a> ");
		for(var i=intFrom;i<=intTo;i++){
			if(i==pageNo)
				write("<font class=\"highlight\">"+i+"</font> ");
			else
				write("<a href=\"?"+paramAdd(strParam,"page",i)+"\">"+i+"</a> ");
		}
		if(intTo<pageCount){
			write("<a href=\"?"+paramAdd(strParam,"page",pageCount)+"\">>></a> ");
		}
		write("转到：<input type=\"text\" size=4 id=\"objPageNo\" style=\"text-align:center\" value=\""+pageNo+"\" ");
		write("onkeydown=\"if(event.keyCode==13){location.replace('?'+paramAdd('"+strParam+"','page',this.value));}\">");
		write("<input type=\"button\" value=\"GO\" onclick=\"location.replace('?'+paramAdd('"+strParam+"','page',objPageNo.value));\">");
		write("</td></tr></table>");
	}
	}
}