// JavaScript Document
var Upload = {
	Code : function(){
		var c = "";
		c += "<div style=\"text-align:center\">";
		c += "<div id=\"MsgContent\">";
		c += "<div id=\"MsgHead\" style=\"text-align:right\"><span style=\"float:left\">PJBlog上传程式</span><a href=\"javascript:;\" onclick=\"$('pjblogdoupload').parentNode.removeChild($('pjblogdoupload'))\">关闭</a></div>";
		c += "<div id=\"MsgBody\">";
		c += "<div id=\"MessageText\"><div id=\"uploadContenter\" style=\"\"></div></div>";
		c += "</div>";
		c += "</div>";
		c += "</div>";
		c += "<iframe style=\"display:none\" name=\"pjblogUploader\"></iframe>";
		return c;
	},
	open : function(obj, inputID, ReturnType, Path, fcount){
		Box.selfWidth = true;
		Box.selefHeight = false;
		var AjaxUp = null;
		var c = Box.FollowBox($(inputID), 400, 0, 5, this.Code());
		this.showUpload(obj, inputID, ReturnType, Path, fcount);
		c.id = "pjblogdoupload";
		c.style.cssText += "; border:1px solid #000;";
	},
	showUpload : function(obj, inputID, ReturnType, Path, fcount){
		/*
			@ obj				对象 一般为 this
			@ inputID			文本框ID
			@ ReturnType		返回类型 0
			@ Path				保存路径
			@ fcount			最大上传文件数
		*/
		try{
			AjaxUp.reset();
		}catch(ex){}
		//创建Uploader，参数Contenter	字符串	上传控件的容器，程序自动给容易四面增加3px的padding
		AjaxUp = new AjaxProcesser("uploadContenter");
		
		//设置提交到的iframe名称
		AjaxUp.target = "pjblogUploader";
		
		AjaxUp.processID = "pjblog";
		
		//上传处理页面
		AjaxUp.url = "../pjblog.logic/log_upload.asp"; 
		
		//上传处理页面
		AjaxUp.MaxFileCount = (isNaN(fcount) ? 999: fcount);
		
		//保存目录
		AjaxUp.savePath = (Path ? Path: "../upload");  
		
		//上传成功时运行的程序
		AjaxUp.succeed = function(files){
			//下面遍历所有的文件，files是一个数组，数组元素的数目就是上传文件的个数，每个元素包含的信息为文件名字和文件大小
			var info = "";
			for(var i = 0 ; i < files.length ; i++){
				info += files[i].path + ";";
			}
			info = info.substr(0, info.length - 1);
			$(inputID).value = info; 
			setTimeout("$('pjblogdoupload').parentNode.removeChild($('pjblogdoupload'))", 1000);
		}
		
		//上传失败时运行的程序
		AjaxUp.faild=function(msg){
			alert("失败原因:" + msg)
		}
	}
}
/*主JS文件*/
function AjaxProcesser(objID){
	this.target = "";
	this.defaultStyle = false;
	this.interValID = 0;//计时器ID
	this.timeTick = 300; //进程查询时间间隔
	this.processID = "pjblog";//进程ID
	this.frm = null;//表单
	this.submit = null;//提交按钮\
	this.processIng = null;
	this.processBar = null;//进度条
	this.process = null;//进度
	this.processInfo = null;//进度详细信息
	this.uploader = null;//隐藏iframe
	this.split = null;//用于添加一个文件的标示
	this.appendTo = $(objID);//容器
	this.appendTo.style.cssText = "padding:3px";//容器样式
	this.files = {count:0};//文件集合
	this.createUploader();//创建AJAX上传对象
	this.startTime = 0;//上传开始时间	
	this.files = {count:0,list:{}};
	this.url = "";
	this.savePath = "";
	this.FileCount = 1;
	this.MaxFileCount = 999;
}
AjaxProcesser.prototype.succeed=function(a){
	return;
};
AjaxProcesser.prototype.faild=function(a){
	return;
};
AjaxProcesser.prototype.addFile=function(){  //对象方法--添加一个文件
	if(this.FileCount >= this.MaxFileCount){alert("超过允许的最大文件上传数量;"); return;}
	_this = this;
	var file = document.createElement("input");//创建一个文件域
	file.type = "file";
	file.name = "file" + getID();
	file.size = 40;
	if(!this.defaultStyle){file.style.cssText = "font-size:9pt;border:1px #dddddd solid;padding:3px 0px 1px 3px;height:20px;";}
	var b = document.createElement("br");
	this.frm.insertBefore(b, this.split);
	this.frm.insertBefore(file, this.split);//添加到表单	
	var remove = document.createElement("input");//生成一个移除按钮
	remove.value = "移除";
	remove.type = "button";
	if(!this.defaultStyle){remove.style.cssText = "font-size:9pt;border:1px #dddddd solid;padding:3px 3px 1px 3px;height:20px;margin-left:3px;";}
	remove.onclick = function(){
		_this.frm.removeChild(this.previousSibling.previousSibling);
		_this.frm.removeChild(this.previousSibling);
		_this.frm.removeChild(this);
		_this.FileCount--;
	};
	this.frm.insertBefore(remove,this.split);//添加到表单
	this.FileCount++;
};
AjaxProcesser.prototype.reset = function(){
	while(this.appendTo.childNodes){
		this.appendTo.removeChild(this.appendTo.lastChild);
	}
};
AjaxProcesser.prototype.createUploader = function(){
	_this = this;
	var frm = document.createElement("form");//创建form表单
	frm.method = "post";
	frm.encoding = "multipart/form-data";
	frm.style.cssText = "padding:0px;margin:0px;";
	var file = document.createElement("input");//创建一个文件域
	file.type = "file";
	file.name = "file" + getID();
	file.size = 40;
	if(!this.defaultStyle){file.style.cssText = "font-size:9pt;border:1px #A69588 solid;padding:3px 0px 1px 3px;height:20px;";}
	this.files[file.name] = file; //添加到文件集合
	frm.appendChild(file);//添加到表单
	var split = document.createElement("br")
	frm.appendChild(split);//创建一个换行,此换行作为添加文件的标示
	this.split = split;

	var button = document.createElement("input");//创建一个按钮,用于上传
	button.value = "上传";
	button.type = "button";
	if(!this.defaultStyle){button.style.cssText = "font-size:9pt;border:1px #A69588 solid;padding:3px 3px 1px 3px;height:20px;margin-top:3px;";}
	button.onclick = function(){
		_this.processID = "pjblog" + getID();
		var action = "";
		action = _this.url + "?path=" + _this.savePath + "&processid=" + _this.processID;
		_this.frm.action = action;
		_this.frm.target = _this.target;
		_this.frm.submit();
		_this.startTime = Date.parse(new Date());
		_this.processDiv.style.display = "block";
		_this.interValID = window.setInterval("_this.getProcess()", _this.timeTick);
	};
	var add=document.createElement("input");//创建一个按钮
	add.value = "添加文件";
	add.type = "button";
	if(!_this.defaultStyle){add.style.cssText = "font-size:9pt;border:1px #A69588 solid;padding:3px 3px 1px 3px;height:20px;margin-top:3px;margin-left:3px;";}
	add.onclick = function(){
		_this.addFile();
	};
	frm.appendChild(button);//把按钮添加到表单
	frm.appendChild(add);//把按钮添加到表单
	this.frm = frm;
	this.appendTo.appendChild(frm);//把表单添加到容器中
	
	var processDiv = document.createElement("div");//创建第二个容器来容纳信息
	processDiv.style.cssText = "display:none;padding:3px;font-size:9pt;border:1px #A69588 solid;width:406px;margin:5px 2px 2px 0px;";
	
	var processIng = document.createElement("div");//创建进度详细信息显示
	processIng.style.cssText = "padding:2px 5px 2px 1px;font-size:9pt;margin:0px;";
	processIng.innerHTML = "进度";
	this.processIng = processIng;
	processDiv.appendChild(processIng);//把进度详细信息显示添加到容器
	
	var processBar = document.createElement("div");//创建一个进度条
	processBar.style.cssText = "font-size:9pt;width:400px;padding:0px;margin:0px;height:auto;border:1px #dddddd solid;background-color:#eeeeee;";
	var process = document.createElement("div");//创建进度
	process.style.cssText = "font-size:9pt;text-align:center;background-color:#aaaaaa;width:0px;height:13px;padding:1px 0px 0px 2px;"
	//process.innerHTML="0.00%";
	this.process = process;
	processBar.appendChild(process);//把进度添加到进度条
	this.processBar = processBar;
	processDiv.appendChild(processBar);//把进度条添加到容器
	
	var processInfo = document.createElement("div");//创建进度详细信息显示
	processInfo.style.cssText = "padding:2px 5px 2px 1px;font-size:9pt;"
	this.processInfo = processInfo;
	processDiv.appendChild(processInfo);//把进度详细信息显示添加到容器
	this.processDiv = this.appendTo.appendChild(processDiv);
};

/*获取上传进程*/
AjaxProcesser.prototype.getProcess = function(){
	var json = unescape(cookie.GET(this.processID));
	//alert(unescape(json));
	msg = json.json();
	if (msg == null){return;}
	var pro = this.getInformation(msg);            //这里返回所有的上传信息,想显示那写信息可以自由决定
	var str = "";
	var img = "∵∴";
	if(pro.finish){img = "<span style=\"font-weight:bold;color:green;\">√ ";}
	if(pro.step == "faild"){img = "<span style=\"font-weight:bold;color:red;\">×";}
	this.processIng.innerHTML = str + img + pro.alt + "</span>";
	str = str + "总大小:" + reSize(pro.total);
	str = str + "&nbsp; <span style=\"color:green;\">已上传:" + reSize(pro.cur) + "</span>";
	str = str + "&nbsp; <span style=\"color:red;\">上传速度:" + pro.speed + "</span>";
	this.processInfo.innerHTML = str;
	this.process.innerHTML = pro.percent;
	this.process.style.width = Math.floor(398 * pro.process) + "px"; //显示进度
	if(pro.finish){
		this.frm.reset();
		window.clearInterval(_this.interValID);
		if(pro.step == "faild"){
			this.faild(pro.msg);
		}
		if(pro.step == "saved"){
			this.succeed(pro.msg);
		}
	}
//	_this = this;
//	Ajax({
//        url:"../getProcess.asp?processid=" + this.processID,
//        method:"POST",
//        oncomplete:function(msg){
//	    msg = msg.responseText.json();
//	    if(msg == null){return;}
//            var pro=_this.getInformation(msg);            //这里返回所有的上传信息,想显示那写信息可以自由决定
//            var str="";
//            var img="∵∴";
//            if(pro.finish){img="<span style=\"font-weight:bold;color:green;\">√ ";}
//            if(pro.step=="faild"){img="<span style=\"font-weight:bold;color:red;\">×";}
//            _this.processIng.innerHTML= str + img + pro.alt + "</span>";
//            str= str + "总大小:" + reSize(pro.total);
//            str= str + "&nbsp; <span style=\"color:green;\">已上传:" + reSize(pro.cur) + "</span>";
//			str= str + "&nbsp; <span style=\"color:red;\">上传速度:" + pro.speed + "</span>";
//            _this.processInfo.innerHTML= str;
//            _this.process.innerHTML=pro.percent;
//            _this.process.style.width=Math.floor(398 * pro.process) + "px"; //显示进度
//            if(pro.finish){
//				_this.frm.reset();
//				window.clearInterval(_this.interValID);
//				if(pro.step=="faild"){
//					_this.faild(pro.msg);
//				}
//				if(pro.step=="saved"){
//					_this.succeed(pro.msg);
//				}
//            }
//        }
//    });
};

/*获取上传信息*/
AjaxProcesser.prototype.getInformation=function(info){
    //信息对象,请不要修改
    var uploadInfo={
        ID:info.ID,         //上传的进程ID
        stepId:0,
        step:info.step,     //上传进程的英文提示
        DT:info.dt,         //上传进程时间
        total:info.total,   //上传的总数据大小(字节)
        cur:info.now,       //已经上传的数据大小
		speed:reSize(parseInt(info.now/((Date.parse(new Date())-this.startTime)/1000))) + "/秒",
        process:(Math.floor((info.now / info.total) * 100) / 100),  //上传进度的小数形式,用于进度条
        percent:(Math.floor((info.now / info.total) * 10000) / 100) + "%", //进程进度的百分比形式
        alt:"",             //上传进程的中文提示
        msg:"",             //用于显示额外信息,例如错误原因等
        finish:false        //是否已经完成
    };
    /*状态说明*/
    switch(info.step){
        case "":
            uploadInfo.alt="正在初始化上传...";
            uploadInfo.stepId=1;
            break;
        case "uploading":
            uploadInfo.alt="正在上传...";
            uploadInfo.stepId=2;
            break;
        case "uploaded":
            uploadInfo.alt="上传完毕,服务器处理数据中...";
            uploadInfo.stepId=3;
            break;
        case "processing":
            uploadInfo.alt="正在处理文件: " + info.description;
            uploadInfo.stepId=4;
            break;
        case "processed":
            uploadInfo.alt="处理数据完毕,准备保存文件...";
            uploadInfo.stepId=5;
            break;
        case "saving":
            uploadInfo.alt="正在保存文件: " + info.description;
            uploadInfo.stepId=6;
            break;
        case "saved":
            uploadInfo.alt="文件保存完毕,上传成功!";
			uploadInfo.msg=eval("[" + info.description.substr(0,info.description.length-1) + "]")
			uploadInfo.stepId=7;
            uploadInfo.finish=true;
            break;
        case "faild":
            uploadInfo.alt="上传失败!";
            uploadInfo.msg=info.description;
            uploadInfo.stepId=8;
            uploadInfo.finish=true;       
            break;
        default:
            uploadInfo.alt="无此操作!";
            uploadInfo.stepId=9;
            uploadInfo.finish=true;
    }
    return uploadInfo;
}


/*格式化数据*/
var reSize =function (num){
    var Size=parseInt(num);
    var res="";
    if(Size < 1024){
       res= Math.floor(Size * 100) /100 + "B"
    }else if(Size >= 1024 && Size < 1048576){
       res= Math.floor((Size / 1024) * 100) /100  + "KB"
    }else if(Size >= 1048576){
       res= Math.floor(((Size / 1024) / 1024) *  100) /100 + "MB"
    }
    return res;
};


/*生成上传ID*/
var getID=function (){
    var mydt=new Date();
    with(mydt){
        var y=getYear();if(y<10){y='0'+y}
        var m=getMonth()+1;if(m<10){m='0'+m}
        var d=getDate();if(d<10){d='0'+d}
        var h=getHours();if(h<10){h='0'+h}
        var mm=getMinutes();if(mm<10){mm='0'+mm}
        var s=getSeconds();if(s<10){s='0'+s}
    }
    var r="000" + Math.floor(Math.random() * 1000);
    r=r.substr(r.length-4);
    return y + m + d + h + mm + s + r;
};