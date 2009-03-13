//|===========================|
//|   UBB编辑器JS代码 1.0     |
//|      作者:舜子(PuterJam)  |
//|   版权所有 2006           |
//|      For IE               |
//|===========================|
var SelectAllow=false
var UBBrange
 document.onselectstart=IESelectStart

function tellPoint()     
{  
//--获取坐标----   
   UBBrange = UBBTextArea[0].createTextRange()
   var oSel = document.selection.createRange()
   var textLength = UBBTextArea[0].innerText.length
   var line, Mchar, total, cl
   UBBrange.moveToPoint(oSel.offsetLeft, oSel.offsetTop)
   UBBrange.moveStart("character", -1*textLength)
   cl = UBBrange.getClientRects()
   line = cl.length
   total = UBBrange.text.length
   UBBrange.moveToPoint(cl[cl.length-1].left, cl[cl.length-1].top)
   UBBrange.moveStart("character", -1*textLength)
   Mchar = total - UBBrange.text.length
   if (oSel.offsetTop != cl[cl.length-1].top) {line++; Mchar = 0}
   else if (UBBTextArea[0].createTextRange().text.substr(UBBrange.text.length, 2) == "\r\n") Mchar -= 2
   UBBrange.moveToPoint(oSel.offsetLeft, oSel.offsetTop)
   UBBrange.moveStart("character",-UBBTextArea[0].value.length)
//-------------    
}

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
	
	UBBScriptLoader.AddScript('ubb_toolbar.asp');
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
  UBBTextArea[0].onclick=tellPoint
  UBBTextArea[0].onkeyup=tellPoint
}

function ctlent() {
	if((event.ctrlKey && window.event.keyCode == 13) || (event.altKey && window.event.keyCode == 83)) {
	  if (document.all.message.value.length>0) SubmitMsg();
	}
}
function AddText(str){
 if (!UBBrange)
  {
  UBBTextArea[0].value+=str
  UBBTextArea[0].focus()
  }else
  {
   UBBrange.text+=str
   UBBTextArea[0].focus()
   UBBrange.select()
  }
}
function IESelectStart(){
	//if (window.event.srcElement.tagName=="INPUT") {SelectAllow=false;return}
	//if (window.event.srcElement!=UBBTextArea[0]) {SelectAllow=false;return false}
  SelectAllow=true
}

//关于
function UBB_about(){
  alert("UBB编辑器 For PL-Blog2 1.0\n作者: 舜子(PuterJam)\n版权所有 2005\n浏览器: "+window.navigator.appName+"\n其他信息: "+window.navigator.appVersion+"\n系统语言: "+window.navigator.systemLanguage)
} 

//更改字体
function UBB_CFont(e){
 var UBBSelectrange=document.selection.createRange()
  if (SelectAllow && UBBSelectrange.text!=""){
   UBBSelectrange.text="[font="+e.value+"]"+UBBSelectrange.text+"[/font]"
   e.selectedIndex=0
   return
  }	
  AddText("[font="+e.value+"][/font]")
     e.selectedIndex=0
	  
}
//更改字体大小
function UBB_CFontSize(e){
 var UBBSelectrange=document.selection.createRange()
  if (SelectAllow && UBBSelectrange.text!=""){
   UBBSelectrange.text="[size="+e.value+"]"+UBBSelectrange.text+"[/size]"
   e.selectedIndex=0
   return
  }	
  AddText("[size="+e.value+"][/size]")
     e.selectedIndex=0
	  
}
//更改字体颜色
function UBB_CFontColor(e){
 var UBBSelectrange=document.selection.createRange()
  if (SelectAllow && UBBSelectrange.text!=""){
   UBBSelectrange.text="[color="+e.value+"]"+UBBSelectrange.text+"[/color]"
   e.selectedIndex=0
   return
  }	
  AddText("[color="+e.value+"][/color]")
     e.selectedIndex=0
	  
}

//粗体
function UBB_bold(){
 var UBBSelectrange=document.selection.createRange()
  if (SelectAllow && UBBSelectrange.text!=""){
   UBBSelectrange.text="[b]"+UBBSelectrange.text+"[/b]"
   return
  }	
  
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
//斜体
function UBB_italic(){
 var UBBSelectrange=document.selection.createRange()
  if (SelectAllow && UBBSelectrange.text!=""){
   UBBSelectrange.text="[i]"+UBBSelectrange.text+"[/i]"
   return
  }	
  
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
//下划线
function UBB_underline(){
 var UBBSelectrange=document.selection.createRange()
  if (SelectAllow && UBBSelectrange.text!=""){
   UBBSelectrange.text="[u]"+UBBSelectrange.text+"[/u]"
   return
  }	
  
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
//删除线
function UBB_deleteline(){
 var UBBSelectrange=document.selection.createRange()
  if (SelectAllow && UBBSelectrange.text!=""){
   UBBSelectrange.text="[s]"+UBBSelectrange.text+"[/s]"
   return
  }	
  
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
//向左对齐
function UBB_justifyleft(){
 var UBBSelectrange=document.selection.createRange()
  if (SelectAllow && UBBSelectrange.text!=""){
   UBBSelectrange.text="[align=left]"+UBBSelectrange.text+"[/align]"
   return
  }	
  
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
//居中
function UBB_justifycenter(){
 var UBBSelectrange=document.selection.createRange()
  if (SelectAllow && UBBSelectrange.text!=""){
   UBBSelectrange.text="[align=center]"+UBBSelectrange.text+"[/align]"
   return
  }	
  
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
//向右对齐
function UBB_justifyright(){
 var UBBSelectrange=document.selection.createRange()
  if (SelectAllow && UBBSelectrange.text!=""){
   UBBSelectrange.text="[align=right]"+UBBSelectrange.text+"[/align]"
   return
  }	
  
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
 var UBBSelectrange=document.selection.createRange()
  if (SelectAllow && UBBSelectrange.text!=""){
   UBBSelectrange.text="[url]"+UBBSelectrange.text+"[/url]"
   return
  }	
  
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
 var UBBSelectrange=document.selection.createRange()
  if (SelectAllow && UBBSelectrange.text!=""){
   UBBSelectrange.text="[mDown]"+UBBSelectrange.text+"[/mDown]"
   return
  }	
  
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
 var UBBSelectrange=document.selection.createRange()
  if (SelectAllow && UBBSelectrange.text!=""){
   UBBSelectrange.text="[ed2k]"+UBBSelectrange.text+"[/ed2k]"
   return
  }	
  
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
 var UBBSelectrange=document.selection.createRange()
  if (SelectAllow && UBBSelectrange.text!=""){
   UBBSelectrange.text="[email]"+UBBSelectrange.text+"[/email]"
   return
  }	
  
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
 var UBBSelectrange=document.selection.createRange()
  if (SelectAllow && UBBSelectrange.text!=""){
   UBBSelectrange.text="[img]"+UBBSelectrange.text+"[/img]"
   return
  }	
  
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
 var UBBSelectrange=document.selection.createRange()
  if (SelectAllow && UBBSelectrange.text!=""){
   UBBSelectrange.text="[list]\n [*]"+UBBSelectrange.text+"\n[/list]"
   return
  }	
  
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
                           AddText("[*]"+PopText)
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
 var UBBSelectrange=document.selection.createRange()
  if (SelectAllow && UBBSelectrange.text!=""){
   UBBSelectrange.text="[quote]"+UBBSelectrange.text+"[/quote]"
   return
  }	
  
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
 var UBBSelectrange=document.selection.createRange()
  if (SelectAllow && UBBSelectrange.text!=""){
   UBBSelectrange.text="[hidden]"+UBBSelectrange.text+"[/hidden]"
   return
  }	
  
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
 var UBBSelectrange=document.selection.createRange()
  if (SelectAllow && UBBSelectrange.text!=""){
   UBBSelectrange.text="[code]"+UBBSelectrange.text+"[/code]"
   return
  }	
  
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
 var UBBSelectrange=document.selection.createRange()
  if (SelectAllow && UBBSelectrange.text!=""){
   UBBSelectrange.text="[html]"+UBBSelectrange.text+"[/html]"
   return
  }	
  
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
 var UBBSelectrange=document.selection.createRange()
  if (SelectAllow && UBBSelectrange.text!=""){
   UBBSelectrange.text="[swf]"+UBBSelectrange.text+"[/swf]"
   return
  }	
  
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
 var UBBSelectrange=document.selection.createRange()
  if (SelectAllow && UBBSelectrange.text!=""){
   UBBSelectrange.text="[wma]"+UBBSelectrange.text+"[/wma]"
   return
  }	
  
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
 var UBBSelectrange=document.selection.createRange()
  if (SelectAllow && UBBSelectrange.text!=""){
   UBBSelectrange.text="[wmv]"+UBBSelectrange.text+"[/wmv]"
   return
  }	
  
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
 var UBBSelectrange=document.selection.createRange()
  if (SelectAllow && UBBSelectrange.text!=""){
   UBBSelectrange.text="[rm]"+UBBSelectrange.text+"[/rm]"
   return
  }	
  
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
  popnew("common/Code/Code.htm","幻码JS版",500,400)
}
function UBB_htmlubb(){
  popnew("common/HTMLUBB.htm","HTMLUBB",460,400)
}

//插入表情
function UBB_smiley(){
  var smileyPos=new getPos('A_smiley')
  smileyPanel=document.getElementById('UBBSmiliesPanel')
  smileyPanel.style.left=smileyPos.Left+"px"
  smileyPanel.style.top=smileyPos.Top+"px"
  smileyPanel.style.visibility ="visible"
  document.body.attachEvent("onclick",CloseSmileyPanel)
}

function CloseSmileyPanel(){
  smileyPanel=document.getElementById('UBBSmiliesPanel')
  smileyPanel.style.visibility ="hidden"
  document.body.detachEvent("onclick",CloseSmileyPanel)
}

function AddSmiley(str){
	AddText(str)
    CloseSmileyPanel()
}

function getPos(obj){
  this.Left=0
  this.Top=0
  var TempLeft
  var tempObj=document.getElementById(obj)
  while (tempObj.tagName.toLowerCase()!="body"){
  	 this.Left+=tempObj.offsetLeft
  	 this.Top+=tempObj.offsetTop
  	 tempObj=tempObj.offsetParent
  	 TempLeft+=tempObj.offsetLeft+","
  }
}