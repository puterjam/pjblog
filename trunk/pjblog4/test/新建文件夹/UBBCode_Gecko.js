//|===========================|
//|   UBB编辑器JS代码 1.0     |
//|      作者:舜子(PuterJam)  |
//|   版权所有 2006           |
//|      For Mozilla          |
//|===========================|

function loadUBB(UBB_Content){
	window._ubbContent = UBB_Content;
	if (window._isUBBLoaded) {
		showUBB();
		return
	}
	initUBB("Message")
	document.getElementById("editorbody").style.display="";
	var me=document.getElementById("editMask")
	me.parentNode.removeChild(me);
	
	UBBScriptLoader.AddScript('../pjblog.model/cls_ubbtools.asp');
	UBBScriptLoader.CheckQueue();
}

function showUBB(){
	window._isUBBLoaded = true;
	document.getElementById("editorHead").innerHTML = ubbTools;
	document.getElementById("UBBSmiliesPanel").innerHTML = ubbSmile;
	

	UBBTextArea=document.getElementsByName(window._ubbContent);
	UBBTextArea[0].focus();
}
	

function initUBB(UBB_Content){
  UBBTextArea=document.getElementsByName(UBB_Content)
}

function AddText(str){
  UBBTextArea[0].value+=str
  UBBTextArea[0].focus()
}
function UBB_about(){
  alert("UBB编辑器 For PL-Blog2 1.0\n作者: 舜子(PuterJam)\n版权所有 2005\n浏览器: "+window.navigator.appName+"\n其他信息: "+window.navigator.appVersion)
} 

function UBB_bold(){
 if (EditMethod=="normal")
 {
	 var PopText
	  if (PopText=window.prompt(bold_normal,"")) {
		AddText("[b]"+PopText+"[/b]")
	  }
 }
 if (EditMethod=="expert")
 {
	 AddText("[b][/b]")
 }
} 

function UBB_italic(){
 if (EditMethod=="normal")
 {
	 var PopText
	  if (PopText=window.prompt(italic_normal,"")) {
		AddText("[i]"+PopText+"[/i]")
   }
 }
 
 if (EditMethod=="expert")
 {
	 AddText("[i][/i]")
 }
} 

function UBB_underline(){
 if (EditMethod=="normal")
 {
	 var PopText
	  if (PopText=window.prompt(underline_normal,"")) {
		AddText("[u]"+PopText+"[/u]")
   }
 }
 
 if (EditMethod=="expert")
 {
	 AddText("[u][/u]")
 }
} 

function UBB_deleteline(){
 if (EditMethod=="normal")
 {
	 var PopText
	  if (PopText=window.prompt(deleteline_normal,"")) {
		AddText("[s]"+PopText+"[/s]")
   }
}
 
 if (EditMethod=="expert")
 {
	 AddText("[s][/s]")
 }
}

function UBB_justifyleft(){
 if (EditMethod=="normal")
 {
	 var PopText
	  if (PopText=window.prompt(left_normal,"")) {
		AddText("[align=left]"+PopText+"[/align]")
   }
 }
 
 if (EditMethod=="expert")
 {
	 AddText("[align=left][/align]")
 }
} 

function UBB_justifycenter(){
 if (EditMethod=="normal")
 {
	 var PopText
	  if (PopText=window.prompt(center_normal,"")) {
		AddText("[align=center]"+PopText+"[/align]")
   }
 }
 
 if (EditMethod=="expert")
 {
	 AddText("[align=center][/align]")
 }
} 

function UBB_justifyright(){
 if (EditMethod=="normal")
 {
	 var PopText
	  if (PopText=window.prompt(right_normal,"")) {
		AddText("[align=right]"+PopText+"[/align]")
   }
 }
 
 if (EditMethod=="expert")
 {
	 AddText("[align=right][/align]")
 }
} 

//超链接
function UBB_link(){
 if (EditMethod=="normal")
 {
	 var PopText,PopUrl
	  if (PopText=window.prompt(link_normal,"")) {
	  	  if (PopUrl=window.prompt(link_normal_input,"http://")) AddText("[url="+PopUrl+"]"+PopText+"[/url]")
   }
 }
 
 if (EditMethod=="expert")
 {
	 AddText("[url][/url]")
 }
}

//会员下载
function UBB_mDown(){
 if (EditMethod=="normal")
 {
	 var PopText,PopUrl
	  if (PopText=window.prompt(mDown_normal,"")) {
	  	  if (PopUrl=window.prompt(mDown_normal_input,"http://")) AddText("[mDown="+PopUrl+"]"+PopText+"[/mDown]")
   }
 }
 
 if (EditMethod=="expert")
 {
	 AddText("[mDown][/mDown]")
 }
}

//eMule超链接
function UBB_ed2k(){
 if (EditMethod=="normal")
 {
	 var PopText,PopUrl
	  if (PopText=window.prompt(ed2k_normal,"")) {
	  	  if (PopUrl=window.prompt(ed2k_normal_input,"http://")) AddText("[ed2k="+PopUrl+"]"+PopText+"[/ed2k]")
   }
 }
 
 if (EditMethod=="expert")
 {
	 AddText("[ed2k][/ed2k]")
 }
}
//邮件
function UBB_mail(){
 if (EditMethod=="normal")
 {
	 var PopText,PopUrl
	  if (PopText=window.prompt(email_normal,"")) {
	  	  if (PopUrl=window.prompt(email_normal_input,"")) AddText("[email="+PopUrl+"]"+PopText+"[/email]")
   }
 }
 
 if (EditMethod=="expert")
 {
	 AddText("[email=][/email]")
 }
}
//帖图像
function UBB_image(){
 if (EditMethod=="normal")
 {
	 var PopText
	  if (PopText=window.prompt(image_normal,"")) {
		AddText("\n[img]"+PopText+"[/img]")
   }
 }
 
 if (EditMethod=="expert")
 {
	 AddText("[img][/img]")
 }
} 
//项目符号
function UBB_insertunorderedlist(){
 if (EditMethod=="normal")
 {
	 var PopText,PopType,PopTypeText
	  if (PopType=window.prompt(list_normal,"")) {
	     while ((PopType!="") && (PopType!="A") && (PopType!="a") && (PopType!="1") && (PopType!=null)) {
                        PopType=prompt(list_normal_error,"");               
                }
              if (PopType==""){
                  AddText("[list]")
              }
              else
              {
              	 switch (PopType){
              	 	 case "1":PopTypeText="decimal";break;
              	 	 case "a":PopTypeText="lower-alpha";break;
              	 	 case "A":PopTypeText="upper-alpha";break;
              	 	 default:PopTypeText=""
              	 }
              	 if (PopTypeText.length>0) {AddText("[list="+PopTypeText+"]")}else{AddText("[list]")}
              }          
           if (PopText=window.prompt(list_normal_input,"")) {
               AddText("\n [*]"+PopText)   	 
         	   while ((PopText!="") && (PopText!=null)) {
                     if (PopText=window.prompt(list_normal_input,"")) {
                           AddText("\n [*]"+PopText)
                 }
               }
           }
               AddText("\n[/list]")          
   }
 }
 
 if (EditMethod=="expert")
 {
	 AddText("[list]\n [*]\n[/list]")
 }
} 

//引用
function UBB_quote(){
 if (EditMethod=="normal")
 {
	 var PopText
	  if (PopText=window.prompt(quote_normal,"")) {
		AddText("[quote]"+PopText+"[/quote]")
   }
 }
 
 if (EditMethod=="expert")
 {
	 AddText("[quote][/quote]")
 }
}

//隐藏引用
function UBB_hidden(){
 if (EditMethod=="normal")
 {
	 var PopText
	  if (PopText=window.prompt(hidden_normal,"")) {
		AddText("[hidden]"+PopText+"[/hidden]")
   }
 }
 
 if (EditMethod=="expert")
 {
	 AddText("[hidden][/hidden]")
 }
}
//代码
function UBB_code(){
 if (EditMethod=="normal")
 {
	 var PopText
	  if (PopText=window.prompt(code_normal,"")) {
		AddText("[code]"+PopText+"[/code]")
   }
 }
 
 if (EditMethod=="expert")
 {
	 AddText("[code][/code]")
 }
}

//可运行HTML代码
function UBB_html(){
 if (EditMethod=="normal")
 {
	 var PopText
	  if (PopText=window.prompt(html_normal,"")) {
		AddText("[html]"+PopText+"[/html]")
   }
 }
 
 if (EditMethod=="expert")
 {
	 AddText("[html][/html]")
 }
}

//Flash
function UBB_flash(){
 if (EditMethod=="normal")
 {
	 var PopText
	  if (PopText=window.prompt(flash_normal,"")) {
		AddText("[swf]"+PopText+"[/swf]")
   }
 }
 
 if (EditMethod=="expert")
 {
	 AddText("[swf][/swf]")
 }
}

//音频
function UBB_music(){
 if (EditMethod=="normal")
 {
	 var PopText
	  if (PopText=window.prompt(wma_normal,"")) {
		AddText("[wma]"+PopText+"[/wma]")
   }
 }
 
 if (EditMethod=="expert")
 {
	 AddText("[wma][/wma]")
 }
}

//视频
function UBB_mediaplayer(){
 if (EditMethod=="normal")
 {
	 var PopText
	  if (PopText=window.prompt(wmv_normal,"")) {
		AddText("[wmv]"+PopText+"[/wmv]")
   }
 }
 
 if (EditMethod=="expert")
 {
	 AddText("[wmv][/wmv]")
 }
}
//Real媒体
function UBB_realplayer(){
 if (EditMethod=="normal")
 {
	 var PopText
	  if (PopText=window.prompt(rm_normal,"")) {
		AddText("[rm]"+PopText+"[/rm]")
   }
 }
 
 if (EditMethod=="expert")
 {
	 AddText("[rm][/rm]")
 }
}
function UBB_highlightcode(){
  popnew("common/Code/Code.htm","幻码JS版",500,430)
}

function UBB_htmlubb(){
  popnew("common/HTMLUBB.htm","HTML-UBB",500,400)
}

//插入表情
function UBB_smiley(){
  var smileyPos=new getPos('A_smiley')
  smileyPanel=document.getElementById('UBBSmiliesPanel')
  smileyPanel.style.left=smileyPos.Left+"px"
  smileyPanel.style.top=smileyPos.Top+"px"
  smileyPanel.style.visibility ="visible"
  document.body.addEventListener("click",CloseSmileyPanel,true)
}

function CloseSmileyPanel(){
  smileyPanel=document.getElementById('UBBSmiliesPanel')
  smileyPanel.style.visibility ="hidden"
  document.body.removeEventListener("click",CloseSmileyPanel,true)
}

function AddSmiley(str){
	AddText(str)
    CloseSmileyPanel()
}

function getPos(obj){
  this.Left=0
  this.Top=0
  var tempObj=document.getElementById(obj)
  while (tempObj.tagName.toLowerCase()!="body"){
  	 this.Left+=tempObj.offsetLeft
  	 this.Top+=tempObj.offsetTop
  	 tempObj=tempObj.offsetParent
  }
}