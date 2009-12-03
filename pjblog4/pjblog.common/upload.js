// JavaScript Document
	var upload = {};
	upload.alowCount = 10;
	upload.fileCount = 2;
	upload.fileArray = new Array(); // [{id:*,name:*,path:*,size:*,type:*,ext:*}[, ....]]
/*
	' @ 	加载附件资源信息
*/
	upload.ReloadSource = function(i){
		var newArray = this.fileArray;
		if (newArray.length > 0){
			var c = newArray.length;
			var left = c - ((i - 1) * this.fileCount) - 1;
			var right = c - 1 - (i * this.fileCount) + 1;
			if (left > (c - 1)){left = c - 1};
			if (right < 0){right = 0}
			var s, k = "";
			$("#source").empty();
			for (s = left; s >= right ; s--){
				var vs = newArray[s];//取得附件信息
				var dv = document.createElement("div");
				k += "<div class=\"item\"><a href=\"javascript:upload.SourceClick(" + s + ");\" class=\"" + vs.ext + "\">" + vs.name + "</a></div>"
			}
			var PageStr = MultiPage(newArray.length, this.fileCount, i, "javascript:upload.ReloadSource({$page});", "float:none", "", false);
			$("#source").html(k)
			$("#SourcePage").html(PageStr);
		}else{
			$("#source").html("")
			$("#SourcePage").html("");
		}
	}
/*
	' @	加载出附件管理窗口
*/
	upload.out = function(html){
		var div = effect.MakeBox.open(html);
		jQuery(div).css({
					border  : "3px solid #C0D1EF"
		}).attr({id : "upload"});
	};
	
/*
	' @	 点击获得资源详细信息
*/
	upload.SourceClick = function(i){
		try{$("#opensee").remove()}catch(e){}
		var info = this.fileArray[i];
			jQuery(effect.windows.open($("#contentMain"), {
					position : 0,
					html     : this.viewChoose(info, i),
					offset   : 1
			}))
			.css({
					background 	: "#fff",
					margin		: "3px",
					border		: "1px solid #AAC1EA",
					padding		: "15px",
					width		: ($("#contentMain").outerWidth() - 37) + "px",
					height		: ($("#contentMain").outerHeight() - 2) + "px"
			})
			.attr("id", "opensee");
			$('#upload').draggable({disabled : true});
	}

/*
	' @	资源加载选择过程
*/
	upload.viewChoose = function(obj){
		var st = obj.type.split("/")[0];
		if (st == "image"){
			return this.Sourceimage(obj);
		}else{return this.AttentMent(obj);}
	}

/*
	' @	主函数,打开窗口方法
*/
	upload.open = function(){
		var div = document.createElement("div")
		jQuery(div).css({
					border : "1px solid #AAC1EA",
					width  : "600px"
		}).addClass("upload").html(this.html())
		this.out(div);
		// 加载资源
		this.ReloadSource(1);
		//$('input[name=uploadType]:checked').val()
		$("#postUpload").ajaxForm({
				dataType : "json",
				beforeSubmit : function(){
					var canUp = true;
					$("input[type*='file']").each(function(){
						if ($(this).val() == "" || $(this).val().length  == 0){canUp = false}							   
					});
					if (canUp) {
						$("#uploadinfo").text("正在上传, 请稍后..."); return true;
					}else{
						alert("上传的附件内容不能为空");
						return false;
					}
				},
				success : function(data){
					var t = data.Info;
					if (t.length > 0){
						$("#uploadinfo").text("共上传 " + t.length + " 个文件成功..");
						for (var n = 0 ; n < t.length ; n++){
							upload.fileArray.push(t[n]);
						}
						upload.ReloadSource(1);
						$("#postUpload").resetForm();
					}else{
						$("#uploadinfo").text("上传文件失败..")
					}
				}
		});
		// 绑定拖动
		$('#upload').draggable({handle : "#uploadTitle"});
	};

/*
	' @	增加上传按钮
*/
	upload.add = function(){
		var i = $(".file").length + 1;
		if (i > this.alowCount){alert("超过允许上传最大量"); return;}
		var file = document.createElement("input");
		var items = document.createElement("div");
		$("#fload")
			.append(
					jQuery(items)
						.addClass("item")
						.attr("id", "item" + i)
						.append(
								"<input type=\"checkbox\" value=\"" + i + "\" class=\"filedel\">"
						)
						.append(
								jQuery(file)
									.addClass("file")
									.attr({
										  size : "40",
										  name : "file-" + i,
										  type : "file"
									})
								)
			);
	}

/*
	' @	全选方法
*/
	upload.selectAll1 = function(){
		$(".filedel").each(function(){
			if (!$(this).attr("checked")){$(this).attr("checked", true)}				
		})
	}

/*
	' @	反选方法
*/
	upload.selectAll2 = function(){
		$(".filedel").each(function(){
			if ($(this).attr("checked")){$(this).attr("checked", false)}				
		})
	}

/*
	' @	自动匹配全选或者反选代码
*/
	upload.selectAll = function(obj){
		if (obj.checked){
			this.selectAll1();
		}else{this.selectAll2();}
	}

/*
	' @	删除上传框代码
*/
	upload.del = function(){
		$(".filedel").each(function(){
			if ($(this).attr("checked")){
				if ($(this).val() == 1){
					alert("不能删除第一个元素");
				} else {
					$("#item" + $(this).val()).remove();
				}
			}						
		})
	}

/*
	' @	更新附件方法
*/
	upload.update = function(obj, id){
		try{this.updateClose()}catch(e){}
		var _id = id, jsid = $("#jsid").val();
		jQuery(obj).attr({disabled : true});
		jQuery(effect.windows.open(jQuery(obj), {
				position	: 	5,
				html		:	this.updateString(obj, id),
				offset		:	1
		})).attr("id", "updateBox");
		//绑定上传
		$("#updatePostBox").ajaxForm({
				dataType : "json",
				beforeSubmit : function(){
					var canUp = true;
					$("#updatePostBox input[type*='file']").each(function(){
						if ($(this).val() == "" || $(this).val().length  == 0){canUp = false}							   
					});
					if (canUp) {
						jQuery(effect.windows.open($("#updateBox"), {
									position 	:  	0,
									html		: 	"<strong>正在上传, 请稍后...</strong>",
									offset		: 	0
						}))
						.addClass("opacity")
						.css({
									width		:	$("#updateBox").outerWidth() + "px",
									height		:	$("#updateBox").outerHeight() + "px",
									color		:	"#000000",
									"text-align":	"center",
									"line-height" : $("#updateBox").outerHeight() + "px"
						}).attr("id", "updateBox2");
					}else{
						alert("上传的附件内容不能为空");
						return false;
					}
				},
				success : function(data){
					var json = data.Info[0];
					var m = upload.fileArray;
					for (var n = 0 ; n < m.length ; n++){
						if (m[n].id == _id){
							m[n] = json;//更新缓存
						}
					}
					upload.ReloadSource(1); // 重新获取资源
					try{
						$("#updateBox2").remove();
						upload.updateClose();
					}catch(e){}
				}
		});
	}
	
	upload.updateClose = function(){
		$("#updateBox").remove();
		$("#b_update").attr({disabled : false});
	}
	
	upload.SingleDel = function(obj, id){
		jQuery(obj).attr("disabled", true);
		var _obj = obj, _id = id;
		$.getJSON(
			"../pjblog.logic/log_Upload.asp",
			{action : "singledel", id : id},
			function(data){
				if (data.Suc){
					var c = $("#jsid").val();
					upload.fileArray = $.grep(upload.fileArray, function(n, i){
									return n.id == parseInt(_id);				  
					}, true);
					upload.ReloadSource(1);
					try{$("#opensee").remove()}catch(e){}
				}else{
					alert(data.Info);
					jQuery(_obj).attr("disabled", false);
				}
			}
		);
	}

/*
	' @	JS分页代码 来源PJ3
*/	
	function MultiPage(Numbers, Perpage, Curpage, Url, Style, event, FirstShortCut){
		var _curPage = parseInt(Curpage);
		var _numbers = parseInt(Numbers);
		event = event || ""
		//var Url = (baseUrl || Request.ServerVariables("URL"))+Url_Add;
		
		var Page, Offset, PageI;
		//	If Int(_numbers)>Int(PerPage) Then
		
		Page = 9;
		Offset = 4;
		
		var Pages, FromPage, ToPage;
		
		if (_numbers % Perpage == 0) {
			Pages = parseInt(_numbers / Perpage);
		}else{
			Pages = parseInt(_numbers / Perpage) + 1;
		}
		if (Pages < 2){
			return "";
		}
		
		FromPage = _curPage - Offset;
		ToPage = _curPage + Page - Offset -1;
		
		if (Page>Pages) {
			FromPage = 1;
			ToPage = Pages;
		}else{
			if (FromPage<1) {
				ToPage = _curPage+1 - FromPage;
				FromPage = 1;
				if ((ToPage - FromPage)<Page && (ToPage - FromPage)<Pages) {ToPage = Page}
			}else if (ToPage>Pages) {
				FromPage = _curPage - Pages + ToPage;
				ToPage = Pages;
				if ((ToPage - FromPage)<Page && (ToPage - FromPage)<Pages) {FromPage = Pages - Page+1}
			}
		}
	
		//html start
		var pageCode = ['<div class="page" style="'+Style+'"><ul><li class="pageNumber">']; // & _curPage&"/"&Pages & " | "
		
		//第一页
		if (_curPage!=1 && FromPage>1) {pageCode.push('<a href="'+Url.replace("{$page}", 1)+' title="第一页"  style="text-decoration:none" ' + event + '>«</a> | ')}
			
		if (!FirstShortCut) {ShortCut = ' accesskey=","'}else{ ShortCut = ''}
		
		//上一页
		if (_curPage!=1) {pageCode.push('<a href="'+Url.replace("{$page}", (_curPage - 1))+'" title="上一页" style="text-decoration:none;"'+ShortCut+' ' + event + '>‹</a> | ')}
		
		//列表部分
		for (PageI = FromPage;PageI<=ToPage;PageI++){
			if (PageI!=_curPage) {
				pageCode.push('<a href="'+Url.replace("{$page}", PageI) + '">'+PageI+'</a> | ');
			}else{
				pageCode.push('<strong>'+PageI+'</strong>');
				if (PageI!=Pages) {pageCode.push(' | ')}
			}
		}
		
		if (!FirstShortCut) {ShortCut = ' accesskey="."'} else {ShortCut = ''}
		
		//下一页
		if (_curPage!=Pages) {pageCode.push('<a href="'+Url.replace("{$page}", (_curPage+1))+'" title="下一页" style="text-decoration:none"'+ShortCut+' ' + event + '>›</a>')}
		
		//最后一页
		if (_curPage!=Pages && ToPage<Pages) {pageCode.push(' | <a href="'+Url.replace("{$page}", (Pages))+'" title="最后一页" style="text-decoration:none" ' + event + '>»</a>')}
		
		//html end
		pageCode.push('</li></ul></div>');
		pageCode = pageCode.join("");
		FirstShortCut = true;
		
		return pageCode;
	}

/*
	' @	图片资源显示窗口代码
*/
	upload.Sourceimage = function(obj, id){
		var c = "";
		c += "<div id=\"openview\">";
		c += "<input type=\"hidden\" id=\"jsid\" value=\"" + id + "\">";
		c += "<div class=\"title\" style>";
		c += "<div class=\"left\"><img src=\"" + obj.path + "\"></div>";
		c += "<div class=\"right\"><div style=\"width:100%;height:50px\"></div><div><strong>" + obj.name + "</strong></div><div>图片类型 : " + obj.ext + " <span style=\"margin-left:20px\">图片大小 : " + obj.size + "<sub>Bytes</sub></span></div></div>";
		c += "</div>";
		c += "<div style=\"clear:both\"></div>";
		c += "<div class=\"body\">";
		c += "<div>图片路径 : <i>" + obj.path + "</i></div>";
		c += "<div>文件类型 : " + obj.type + "</div>";
		c += "<div><a href=\"javascript:void(0);\" onclick=\"upload.update(this, " + obj.id + ");\" id=\"b_update\">更新</a> | <a href=\"javascript:void(0);\" onclick=\"upload.SingleDel(this, " + obj.id + ")\">删除</a> | <a href=\"javascript:;\">插入</a></div>";
		c += "</div>";
		c += "<div class=\"rbottom\"><input type=\"button\" class=\"button\" value=\"取消\" onclick=\"$('#upload').draggable({disabled : false});$('#opensee').remove();\"></div>"
		c += "</div>";
		return c;
	}
	
/*
	' @	附件详细信息窗口代码
*/
	upload.AttentMent = function(obj, id){
		var c = "";
		c += "<div id=\"openview\">";
		c += "<input type=\"hidden\" id=\"jsid\" value=\"" + id + "\">";
		c += "<div class=\"title\" style=\"padding: 2px 20px 2px 20px\">";
		c += "<h4>附件详细信息</h4>";
		c += "<div>附件名 : " + obj.name + "</div>"
		c += "<div>附件大小 : " + obj.size + "<sub>Bytes</sub></div>"
		c += "<div>附件后缀 : " + obj.ext + "</div>"
		c += "</div>";
		c += "<div style=\"clear:both\"></div>";
		c += "<div class=\"body\">";
		c += "<div>文件真实路径 : <i>" + obj.path + "</i></div>";
		c += "<div>文件类型 : " + obj.type + "</div>";
		c += "<div><a href=\"javascript:void(0);\" onclick=\"upload.update(this, " + obj.id + ");\" id=\"b_update\">更新</a> | <a href=\"javascript:void(0);\" onclick=\"upload.SingleDel(this, " + obj.id + ")\">删除</a> | <a href=\"javascript:;\">插入</a></div>";
		c += "</div>";
		c += "<div class=\"rbottom\"><input type=\"button\" class=\"button\" value=\"取消\" onclick=\"$('#upload').draggable({disabled : false});$('#opensee').remove();\"></div>"
		c += "</div>";
		return c;
	}
	
/*
	' @	主窗口HTML代码
*/
	upload.html = function(){
		var c = "";
		var left = this.left();
		var right = this.right();
		c += "<form action=\"../pjblog.logic/log_Upload.asp?action=fast\" method=\"post\" enctype=\"multipart/form-data\" id=\"postUpload\">"; // form
		c += "<div class=\"title\" id=\"uploadTitle\"><span style=\"float:left\">附件管理</span><span class=\"close\" onclick=\"$('#upload').remove()\"></span><div style=\"clear:both\"></div></div>";
		c += "<div class=\"body\">"
		c += "<div class=\"left\">" + left + "</div>"
		c += "<div class=\"right\" id=\"contentMain\">" + right + "</div>";
		c += "<div style=\"clear:both\"></div>";
		c += "</div>";
		c += "<div class=\"bottom\"><input type=\"submit\" class=\"button\" value=\"上传\"><input type=\"button\" class=\"button\" value=\"添加\" onclick=\"upload.add()\"><input type=\"button\" class=\"button\" value=\"删除\" onclick=\"upload.del()\"><input type=\"button\" class=\"button\" value=\"取消\" onclick=\"$('#upload').remove()\"></div>"
		c += "</form>";
		return c;
	};
	
/*
	' @	主窗口左边代码
*/
	upload.left = function(){
		var c = "<div class=\"sourcetitle\">资源管理</div>";
		c += "<div class=\"source\" id=\"source\">";
		c += "</div>";
		c += "<div id=\"SourcePage\"></div>"
		return c;
	};
	
/*
	' @	主窗口右边代码
*/
	upload.right = function(){
		var c = "<div class=\"main\">";
		c += "<div class=\"ctop\"><span style=\"float:left\"><input type=\"checkbox\" value=\"1\" class=\"check\" onclick=\"upload.selectAll(this)\"> 全选</span><span style=\"float:right; color:#333\"><span style=\"cursor:pointer\">上传模式</span> <span style=\"cursor:pointer\">信息模式</span></span></div>";
		c += "<div class=\"cbottom\">";
		c += "<div style=\"text-align:right; padding-right:20px\"><span style=\"float:left;font-style:oblique; font-size:10px\">▶ CopyRight @ PJBlog 4</span><span id=\"uploadinfo\">&nbsp;</span></div>";
		// 模式开始
		
		c += "<div class=\"fload\" id=\"fload\" style=\"min-height:120px\">";
		c += "<div class=\"item\" id=\"item1\"><input type=\"checkbox\" value=\"1\" class=\"filedel\"><input type=\"file\" class=\"file\" size=\"40\" name=\"file-1\"></div>";
		c += "</div>";
		
		// 模式结束
		c += "</div>";
		c += "</div>";
		return c;
	};
	
	upload.updateString = function(obj, id){
		var c = "";
		c += "<form action=\"../pjblog.logic/log_Upload.asp?action=update&id=" + id + "\" method=\"post\" id=\"updatePostBox\" enctype=\"multipart/form-data\">"
		c += "<div><a href=\"javascript:void(0);\" onclick=\"upload.updateClose()\" style=\"margin-right:6px\"><img src=\"../images/check_error.gif\" border=\"0\"></a><input type=\"file\" style=\"width:240px; border:1px solid #ccc\"  name=\"evio\"><input type=\"submit\" class=\"button\" value=\"更新\"></div>";
		c += "</form>";
		return c;
	}