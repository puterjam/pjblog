//==========================================
// JS For PJBlog2 Background Control
//   By PuterJam 2006-6-10
//==========================================

	 function CheckMove(){
		  if (document.forms[0].source.value=="null") {
		   alert("请选择日志分类源")
		   document.forms[0].source.focus()
		   return false
		  }
		  if (document.forms[0].target.value=="null") {
		   alert("请选择需要移动到的目标日志分类")
		   document.forms[0].target.focus()
		   return false
		  }
		  if (document.forms[0].source.value==document.forms[0].target.value){
		   alert("源分类和目标分类一致无法移动")
		   return false
		  }
		return true
	  }

     function checkAll(e){
      for (i=0;i<document.forms[0].length;i++){
       if (document.forms[0][i].tagName=="INPUT"){
        if (document.forms[0][i].type=="checkbox") {document.forms[0][i].checked=!((!e)?(!window.cAll)?true:false:e.checked);document.forms[0][i].click();}
       }
      }
      if (!e) {window.cAll = !window.cAll}
     }
     
     function CheckDel(id){
	     if (confirm("确定要删除吗？\n\n假如你要删除的日志分类中包含日志，那么分类中的日志也会被删除")){
		  document.forms[0].DelCate.value=id
		  document.forms[0].whatdo.value="DelCate"
		  document.forms[0].submit()
		 }
	  }
	  
	  function CheckDelGroup(id){
	     if (confirm("确定要删除吗？\n\n假如你删除某个权限分组，那么该分组中的会员将会被重置为一般会员")){
		  document.forms[0].DelGroup.value=id
		  document.forms[0].whatdo.value="DelGroup"
		  document.forms[0].submit()
		 }
	  }
	  
    function addSpanKey(){
            var kWord = document.getElementById("keyWord")
            if (kWord.value.length<1 || kWord.value=="请输入关键字") {kWord.value="请输入关键字";kWord.style.backgroundColor="#ffc4a4";kWord.select();return false}
	        var splitWord=kWord.value.split(/[,\s]/g)
	        for (var i=0;i<splitWord.length;i++){
		        if(window.event)
			        document.getElementById("spamList").add(new Option(splitWord[i],splitWord[i]),0);
			       else
			        document.getElementById("spamList").insertBefore(new Option(splitWord[i],splitWord[i]),document.getElementById("spamList").firstChild);
		      }
            document.getElementById("spamList").selectedIndex = 0;
            kWord.value="";
            kWord.select()
            return false
    }
    
    function clearKey() {
           while (document.getElementById("spamList").selectedIndex>-1) {
             document.getElementById("spamList").remove(document.getElementById("spamList").selectedIndex)
           }
    }
    
    function submitList(){
          var spam=document.getElementById("spamList")
          var spamStr=""
          for (var i=0;i<spam.length;i++){
             spamStr += spam.options[i].value + ", ";
          }
          document.getElementById("keyList").value=spamStr;
          document.forms["filter"].submit();
    }

  function DelModule(id,isSystem){
    	if (isSystem=="True"){alert("系统内置模块无法删除！！")}
    	else{
    	 if (confirm("是否删除此模块？")){
    		  document.forms[0].DoID.value=id
    		  document.forms[0].whatdo.value="delModule"
    		  document.forms[0].submit()  
		  }
    	}
  }

  function DelPlugins(ID,Type){
       if (window.confirm("是否删除该插件？")) {
	       	if (window.confirm("保留和该插件相关的数据？")){
	      	  location="?Fmenu=Skins&Smenu=UnInstallPlugins&Keep=1&Plugins="+ID
	       	}
	       	else
	       	{
	      	  location="?Fmenu=Skins&Smenu=UnInstallPlugins&Keep=0&Plugins="+ID
	       	}
       }
  }
  
  function FixPlugins(){
	      location="?Fmenu=Skins&Smenu=FixPlugins"
  }
  
  function setSkin(path,Name){ 
		  document.forms[0].SkinPath.value=path
		  document.forms[0].SkinName.value=Name
		  document.forms[0].submit()
  }
  
 function delUser(id){
  if (confirm('是否删除该用户？')){
   document.forms[0].whatdo.value="DelUser"
   document.forms[0].DelID.value=id
   document.forms[0].submit()
   }
 }

 function Dellink(id){
	     if (confirm("确定要删除该友情链接吗？")){
		  document.forms["Link"].ALinkID.value=id
		  document.forms["Link"].whatdo.value="DelLink"
		  document.forms["Link"].submit()
		 }
 }
 
 function Toplink(id){
		  document.forms["Link"].ALinkID.value=id
		  document.forms["Link"].whatdo.value="TopLink"
		  document.forms["Link"].submit()
 }
 
 function ShowLink(id){
		  document.forms["Link"].ALinkID.value=id
		  document.forms["Link"].whatdo.value="ShowLink"
		  document.forms["Link"].submit()
 }

 function focusMe(o){
 	 o.inFocus=true;
 	 o.style.border="1px solid #AAC1E6";
 	 o.style.backgroundColor="#fff";
 	 o.style.cursor="text";
 	 o.style.overflowY="scroll";
	 reHeight(o);
 }
 
 function blurMe(o){
 	 o.inFocus=false;
	 o.style.border="0px #fff solid"
 	 o.style.padding="1px";
 	 o.style.backgroundColor="transparent";
 	 o.style.cursor="pointer";
 	 o.style.overflowY="hidden";
 	 o.style.height="33px";
 }
 function overMe(o){
 	 if (!o.inFocus) {
 	 	 o.style.border="1px solid #AAC1E6";
 	 	 o.style.padding="0px";
 	 	 o.style.backgroundColor="#F3F8FF";
 	 }
 }
 
 function outMe(o){
 	 if (!o.inFocus) {
	 	 o.style.border="0px #fff solid"
 	 	 o.style.padding="1px";
 	 	 o.style.backgroundColor="transparent";
  	 }
 }
 
 function checkMe(o){
 	 if (o.value.length>0) {
 	 	  o.previousSibling.innerHTML="回复内容:"
 	 }
 	 else{
 	 	  o.previousSibling.innerHTML="回复内容:<span class=\"tip\">(无回复留言)</span>"
 	 }
 }
 
 function reHeight(o){
  	 if (o.inFocus) {o.style.height = "120px";}
 }
 	 
 function highLight(o){
 	 o.parentNode.style.backgroundColor = (o.checked)?"#AAC1E6":"#E6EBF7"
 	 o.parentNode.parentNode.style.border = (o.checked)?"1px solid #AAC1E6":"1px solid #D6DEF1";
 }
 
 function DelModule(id,isSystem){
    	if (isSystem=="True"){alert("系统内置模块无法删除！！")}
    	else{
    	 if (confirm("是否删除此模块？")){
    		  document.forms[0].DoID.value=id
    		  document.forms[0].whatdo.value="delModule"
    		  document.forms[0].submit()  
		  }
    	}
  }
  
  function DelComment(){
    	 if (confirm("是否删除所选的项目？")){
    		  document.forms[0].doModule.value="DelSelect"
    		  document.forms[0].submit()  
		  }
  }
  
  function updateComment(){
    		  document.forms[0].doModule.value="Update"
    		  document.forms[0].submit()  
  }
  
  //regFilter
	 var strREHTML = '<input name="des" type="text" class="text" style="width:150px"/> ';
	  	 strREHTML += '<input name="re" type="text" class="text" style="width:300px"/> ';
		 strREHTML += '<input name="times" type="text" class="text" style="width:30px" maxlength="3"/> ';
		 strREHTML += '<img src="images/icon_del.gif" onclick="delRule(this)" alt="删除规则" style="cursor:pointer;margin-bottom:-2px"/> ';
		 strREHTML += '<span></span>';

	function delRule(o){
		var p = o.parentNode;
		p.parentNode.removeChild(p);
	}

	function testRules(){
		var rL = document.getElementById("reList");
		var ds = rL.getElementsByTagName("DIV");
		var ts = document.getElementById("testArea").value;
		var getError = false
		
		for (var i=0;i<ds.length;i++){
			var inp = ds[i].getElementsByTagName("INPUT");
			var sp = ds[i].getElementsByTagName("SPAN")[0];
			var r = rgExec(ts,inp[1].value,inp[2].value)
			if (r>0) {
				sp.style.color="#0000ff"
				sp.innerHTML = "get."
				sp.title = "示范文本中存在可以有匹配的项目．"
			}
			
			if (r == 0){
				sp.style.color="#008000"
				sp.innerHTML = "ok."
				sp.title = "正则式正常．"
			}
			
			if (r<0){
				sp.style.color="#f00"
				sp.innerHTML = "err."
				sp.title = "正则出现错误．"
				getError = true
			}
		}
		return getError
	}

	function submitReList(){
		  if (testRules()) {
		  	  alert("发现错误的正则式")
		  	  return
		  }
          document.forms["reFilter"].submit();
	}

	function addRules(){
		var rTitle = document.getElementById("reTitle");
		var rRules = document.getElementById("reRules");
		var rTimes = document.getElementById("reTimes");
		
		if (rTitle.value.length == 0) {
			alert("请输入规则描述");
			rTitle.select();
			return;
		}
		
		if (rRules.value.length == 0) {
			alert("请输入正则表达式");
			rRules.select();
			return;
		}
		
	 	try{
	 		var re  = new RegExp(rRules.value,"ig");
	 	}catch(e){
			alert(e.name + ": " + e.description);
			rRules.select();
			return;
	 	}
	 	
	 	if (rTimes.value.length == 0 || /[^0-9]/i.test(rTimes.value)) {
			alert("请输入最大匹配次数");
			rTimes.select();
			return;
		}
		addRule(rTitle.value,rRules.value,rTimes.value)
		rTitle.value = rRules.value = rTimes.value = "";
	}
	
	function addRule(rTitle,rRules,rTimes){
		var rL = document.getElementById("reList");
		var r = document.createElement("div");
		r.innerHTML = strREHTML;
		rL.appendChild(r);
		
		var inp = r.getElementsByTagName("INPUT");
		
		inp[0].value = rTitle;
		inp[1].value = rRules;
		inp[2].value = rTimes;
	}