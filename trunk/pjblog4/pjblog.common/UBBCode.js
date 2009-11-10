//|===========================|
//|   UBB编辑器JS代码 1.0     |
//|      作者:舜子(PuterJam)  |
//|   版权所有 2005           |
//|===========================|

var UBBBrowerInfo=new Object();
var sAgent=navigator.userAgent.toLowerCase();
UBBBrowerInfo.IsIE=sAgent.indexOf("msie")!=-1;
UBBBrowerInfo.IsGecko=!UBBBrowerInfo.IsIE;UBBBrowerInfo.IsNetscape=sAgent.indexOf("netscape")!=-1;
if (UBBBrowerInfo.IsIE){
	UBBBrowerInfo.MajorVer=navigator.appVersion.match(/MSIE (.)/)[1];
	UBBBrowerInfo.MinorVer=navigator.appVersion.match(/MSIE .\.(.)/)[1];}
else{
	UBBBrowerInfo.MajorVer=0;UBBBrowerInfo.MinorVer=0;
	};
	UBBBrowerInfo.IsIE55OrMore=UBBBrowerInfo.IsIE&&(UBBBrowerInfo.MajorVer>5||UBBBrowerInfo.MinorVer>=5);

var UBBScriptLoader=new Object();
UBBScriptLoader.IsLoading=false;
UBBScriptLoader.Queue=new Array();
UBBScriptLoader.AddScript=function(scriptPath){
	UBBScriptLoader.Queue[UBBScriptLoader.Queue.length]=scriptPath;
	//if (!this.IsLoading) this.CheckQueue();
	};
	
UBBScriptLoader.CheckQueue = function(){
	if (this.Queue.length>0){
		this.IsLoading=true;
		var sScriptPath=this.Queue[0];
		var oTempArray=new Array();
		for (i=1;i<this.Queue.length;i++) oTempArray[i-1]=this.Queue[i];
		this.Queue=oTempArray;
		var el;
		if (sScriptPath.lastIndexOf('.css')>0){
			 el=document.createElement('LINK');
			 el.rel='stylesheet';e.type='text/css';
			 el.href=sScriptPath;
		}else	{
			 el=document.createElement("script");
			 el.type="text/javascript";
			 el.language="javascript";
			 el.src=sScriptPath;
		};
		document.getElementsByTagName("head")[0].appendChild(el);
						
		var oEvent = function(){
			if (this.tagName=='LINK'||!this.readyState||this.readyState=='loaded') UBBScriptLoader.CheckQueue();
		};
		
		if (el.tagName=='LINK'){
			if (UBBBrowserInfo.IsIE) el.onload=oEvent;else UBBScriptLoader.CheckQueue();
		}else{
			el.onload=el.onreadystatechange=oEvent;
		};
		

	}
	else
	{
		this.IsLoading=false;
		if (this.OnEmpty) this.OnEmpty();
	};
}


var EditMethod="normal"
var UBBTextArea

//UBBBrowerInfo.IsIE 判断是否是IE
//UBBBrowerInfo.IsGecko 判断是否是Gecko
//初试化代码

if (UBBBrowerInfo.IsIE){
 	UBBScriptLoader.AddScript('../pjblog.common/UBBCode_IE.js')
}

if (UBBBrowerInfo.IsGecko){
 	UBBScriptLoader.AddScript('../pjblog.common/UBBCode_Gecko.js')
}
UBBScriptLoader.CheckQueue();