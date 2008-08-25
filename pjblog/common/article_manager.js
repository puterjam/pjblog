
var NVersion="0.2"
 	
function checkVersion(N,C){
	 var NV = parseInt(N.split(".").join(""))
	 var CV = parseInt(C.split(".").join(""))
	 if (!CV || CV<NV) {
		return true
	 }
	 return false
}

try{
	if (checkVersion(NVersion,curver)) 
	{ 
		document.write ("<a href='http://www.leoyung.com/article/7903.htm'><font color=red>最新版本：V"+ NVersion +" </font></a>");
	}
	else
	{
		document.write ("<a href='http://www.leoyung.com/article/7903.htm'>最新版本：V"+ NVersion +"</a>");
	}
}
catch(e){}
