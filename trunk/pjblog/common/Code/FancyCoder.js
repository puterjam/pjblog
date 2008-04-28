/***********************************************
*      幻码JS版 - FancyCoder for JavaScript
*              version 1.0 beta
***********************************************/

//检查语法加亮规则定义，并做相应调整
//lang: 可为0个或多个语言ID，分别表示各个需要检查的语言，如果不写则表示检查所有
function FCCheckSyntaxDef(/*lang, ...*/) {
	//设定语言列表
	if (arguments.length > 0) {
		var langList = {};
		for (var i = arguments.length - 1; i >= 0; i--) {
			if (FCSyntaxDef[arguments[i]] != null) langList[arguments[i]] = true;
		}
	} else {
		var langList = FCSyntaxDef;
	}
	//依次检查各组语言
	for (var lang in langList) {
		var syntax = FCSyntaxDef[lang];
		//检查普通文本设置
		if (syntax.color == null) syntax.color = "#000000";
		if (syntax.style == null) syntax.style = "";
		else syntax.sytle = syntax.style.toLowerCase();
		//检查词定界符设置
		var delim = syntax.delimiters;
		if (delim == null) {
			syntax.delimiters = "~!@%^&*()-+=|\\/{}[]:;\"'<>,.? \t";
		} else if (delim.indexOf(" \t") < 0) {
			syntax.delimiters += " \t";
		}
		//检查注释设置
		if (syntax.comments != null) {
			if (typeof(syntax.comments) == "string") syntax.comments = syntax.comments.split(" ");
			if (syntax.cmtcolor == null) syntax.cmtcolor = "#00ff00";
			if (syntax.cmtstyle == null) syntax.cmtstyle = "";
			else syntax.cmtsytle = syntax.cmtstyle.toLowerCase();
		} else {
			syntax.comments = [];
		}
		//检查块设置
		var blocks = syntax.blocks;
		for (var classid in blocks) {
			var block = blocks[classid];
			if (block.color == null) block.color = "#00ffff";
			if (block.style == null) block.style = "";
			else block.style = block.style.toLowerCase();
			block.lines = block.lines == true;
		}
		//检查关键字设置
		var keywords = syntax.keywords;
		for (var classid in keywords) {
			var group = keywords[classid];
			if (group.color == null) group.color = "#0000ff";
			if (group.style == null) group.style = "";
			else group.style = group.style.toLowerCase();
			group.list = (" " + (group.list instanceof Array ? group.list.join(" ")
				: group.list) + " ").replace(/  +/g, " ");
		}
	}
}
//--------------------------------------------------------------

//创建CSS样式段
//classid: CSS样式段ID
//font: 字体
//size: 字体大小
//color: 字体颜色
//style: 字体风格
function FCMakeCSSClass(classid, color, style, font, size) {
	return "." + classid + " {\r\n\tcolor: " + color + ";\r\n"
		+ (style.indexOf("b") < 0 ? "" : "\tfont-weight: bold;\r\n")
		+ (style.indexOf("i") < 0 ? "" : "\tfont-style: italic;\r\n")
		+ (style.indexOf("u") < 0 ? "" : "\ttext-decoration: underline;\r\n")
		+ (font == null ? "" : "\tfont-family: " + font + ";\r\n")
		+ (size == null ? "" : "\tfont-size: " + size + ";\r\n")
		+ "}\r\n";
}
//--------------------------------------------------------------

//创建指定语言的CSS样式，返回转换好的CSS代码，如果语言不存在则返回null
//lang: 语法加亮规则的语言ID
//font: 所用字体，如果为对象，就用每个对象成员对应相应的classid
//size: 所用字体的大小，如果为对象，就用每个对象成员对应相应的classid
function FCMakeCSS(lang, font, size) {
	var syntax = FCSyntaxDef[lang];
	if (syntax == null) return null;
	var fontList = font instanceof Object;
	var sizeList = size instanceof Object;
	//定义普通文本样式
	var css = FCMakeCSSClass(lang + "_Default", syntax.color, syntax.style,
		fontList ? font.comment : font, sizeList ? size.comment : size);
	//定义注释样式
	if (syntax.comments.length > 0) {
		css += FCMakeCSSClass(lang + "_Comments", syntax.cmtcolor, syntax.cmtstyle,
			fontList ? font.comment : font, sizeList ? size.comment : size);
	}
	//定义块样式
	for (var classid in syntax.blocks) {
		var block = syntax.blocks[classid];
		css += FCMakeCSSClass(lang + "_" + classid, block.color, block.style,
			fontList ? font[classid] : font, sizeList ? size[classid] : size);
	}
	//定义关键词样式
	for (var classid in syntax.keywords) {
		var group = syntax.keywords[classid];
		css += FCMakeCSSClass(lang + "_" + classid, group.color, group.style,
			fontList ? font[classid] : font, sizeList ? size[classid] : size);
	}
	return css;
}
//--------------------------------------------------------------

//创建词缀，返回词缀定义{prefix:前缀,suffix:后缀}
//mode: 转换模式（0：<font>模式，1：<span>模式，2：<span>模式带css，3:[UBB]模式，默认0）
//classid: 加亮类别ID
//color: 加亮颜色
//style: 字体风格
function FCMakeAffix(mode, classid, color, style) {
	if (mode == 1 || mode == 2) {
		return {
			prefix : "<SPAN class='" + classid + "'>",
			suffix : "</SPAN>"
		};
	} else if (mode == 3) {
		var nb = style.indexOf("b") >= 0;
		var ni = style.indexOf("i") >= 0;
		var nu = style.indexOf("u") >= 0;
		return {
			prefix : "[color=" + color + "]" + (nb?"[b]":"") + (ni?"[i]":"") + (nu?"[u]":""),
			suffix : (nu?"[/u]":"") + (ni?"[/i]":"") + (nb?"[/b]":"") + "[/color]"
		};
	} else {
		var nb = style.indexOf("b") >= 0;
		var ni = style.indexOf("i") >= 0;
		var nu = style.indexOf("u") >= 0;
		return {
			prefix : "<FONT color='" + color + "'>" + (nb?"<B>":"") + (ni?"<I>":"") + (nu?"<U>":""),
			suffix : (nu?"</U>":"") + (ni?"</I>":"") + (nb?"</B>":"") + "</FONT>"
		};
	}
}
//--------------------------------------------------------------

//对给定字符串进行HTML编码
//str: 要被编码的字符串
function FCHtmlEncode(str) {
	return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
//--------------------------------------------------------------

//对给定字符串进行UBB编码
//str: 要被编码的字符串
function FCUbbEncode(str) {
	return str.replace(/\[/g, "[").replace(/\]/g, "]");
}
//--------------------------------------------------------------

//转换源代码到HTML加亮代码，返回转换后代码，如果语言不存在则返回null
//srcCode: 需要被转换的源代码
//lang: 转换所用的语法加亮规则的语言ID
//mode: 转换模式（默认0）
function FCTranslate(srcCode, lang, mode) {
	var syntax = FCSyntaxDef[lang];
	if (syntax == null) return null;
	var src = srcCode.split("\r\n");
	var encode = mode == 3 ? FCUbbEncode : FCHtmlEncode;
	//定义普通文本加亮词缀
	var affix = FCMakeAffix(mode, lang + "_Default", syntax.color, syntax.style);
	var defPref = affix.prefix, defSuff = affix.suffix;
	//定义注释加亮词缀
	var comments = syntax.comments;
	if (comments.length > 0) {
		var affix = FCMakeAffix(mode, lang + "_Comments", syntax.cmtcolor, syntax.cmtstyle);
		var cmtPref = affix.prefix, cmtSuff = affix.suffix;
	}
	//定义块加亮词缀
	var blkBegin = [], blkEnd = [], blkEsc = [], blkLines = [], blkPref = [], blkSuff = [];
	for (var classid in syntax.blocks) {
		var block = syntax.blocks[classid];
		blkBegin.push(block.begin);
		blkEnd.push(block.end);
		blkEsc.push(block.escape);
		blkLines.push(block.lines);
		var affix = FCMakeAffix(mode, lang + "_" + classid, block.color, block.style);
		blkPref.push(affix.prefix);
		blkSuff.push(affix.suffix);
	}
	//定义关键词加亮词缀表
	var keywords = [], kwPref = [], kwSuff = [];
	for (var classid in syntax.keywords) {
		var group = syntax.keywords[classid];
		keywords.push(group.list);
		var affix = FCMakeAffix(mode, lang + "_" + classid, group.color, group.style);
		kwPref.push(affix.prefix);
		kwSuff.push(affix.suffix);
	}
	//断词转换
	var delim = syntax.delimiters;
	for (var index = 0, index2 = 0; index < src.length; index++, index2++) {
		var code = src[index];
		var htmlCode = "";
		for (var pos1 = 0, pos2 = 0, ch = null, flag = 0; ch != ""; pos2++) {
			ch = code.substr(pos2, 1);
			if (ch != "" && flag == 0 && delim.indexOf(ch) < 0) continue;
			//如果为持续文本则继续，否则截断（flag：0文本，1空格，2标点）
			if (pos2 <= pos1) {
				flag = ch.match(/s/g) ? 1 : 2;
				continue;
			}
			var word = code.substr(pos1, pos2 - pos1); //截取词
			if (flag == 1) { //空格
				htmlCode += word;
			} else {
				if (flag == 2) { //标点
					//判别注释
					for (var i in comments) {
						if (code.substr(pos1, comments[i].length) != comments[i]) continue;
						htmlCode += cmtPref + encode(code.substr(pos1)) + cmtSuff;
						word = "";
						break;
					}
					if (word == "") break;
					//判别块
					for (var i in blkBegin) {
						if (code.substr(pos1, blkBegin[i].length) != blkBegin[i]) continue;
						var end = blkEnd[i], esc = blkEsc[i];
						for (pos2 = pos1 + blkBegin[i].length; pos2 = code.indexOf(end, pos2);) {
							if (pos2 < 0) {
								if (blkLines[i] && index < src.length - 1) {
									pos2 = code.length + 2;
									code += "\r\n" + src[++index];
									continue;
								}
								htmlCode += blkPref[i] + encode(code.substr(pos1)) + blkSuff[i];
								word = "";
								break;
							} else if (esc == null || code.substr(pos2 - esc.length, esc.length) != esc) {
								pos2 += end.length;
								break;
							}
							pos2 += end.length;
						}
						if (pos2 >= 0) {
							htmlCode += blkPref[i] + encode(code.substr(pos1, pos2 - pos1)) + blkSuff[i];
							flag = 0;
							pos1 = pos2;
							pos2--;
							word = "0";
						}
						break;
					}
					if (word == "") break;
					else if (word == "0") continue;
				}
				//关键字加亮
				var w = encode(word);
				for (var i in keywords) {
					if (keywords[i].indexOf(" " + word + " ") < 0) continue;
					htmlCode += kwPref[i] + w + kwSuff[i];
					word = "";
					break;
				}
				if (word != "") htmlCode += w;
			}
			flag = delim.indexOf(ch) < 0 ? 0 : ch.match(/s/g) ? 1 : 2;
			pos1 = pos2;
		}
		src[index2] = htmlCode;
	}
	src.splice(index2, src.length);
	return defPref + src.join("\r\n") + defSuff;
}
//--------------------------------------------------------------

//还原HTML加亮代码到源代码，返回还原后的代码
//htmlCode: 需要被还原的HTML代码
function FCRevert(htmlCode) {
}
//--------------------------------------------------------------

//生成预览窗口并居中屏幕显示
//title: 窗口标题
//content: 预览的HTML文本内容
//wndWidth: 窗口宽度（默认640）
//wndHeight: 窗口高度（默认480）
function FCPreview(title, content, wndWidth, wndHeight) {
	if (!(wndWidth > 0)) wndWidth = 640;
	if (!(wndHeight > 0)) wndHeight = 480;
	var left = screen.width/2 - wndWidth/2;
	var top = screen.height/2 - wndHeight/2;
	var previewWnd = window.open("", "FCPreviewWnd", "scrollbars=yes,resizable=yes,menubar=yes,"
		+ "width=" + wndWidth + ",height=" + wndHeight + ",left=" + left + ",top=" + top
		+ ",screenX=" + left + ",screenY=" + top);
	previewWnd.document.write("<HTML><HEAD><TITLE>" + title + "</TITLE></HEAD>\r\n<BODY leftmargin='0'"
		+ " topmargin='0' marginwidth='0' marginheight='0'><TABLE width='200'><TR><TD><PRE>\r\n"
		+ content + "\r\n</PRE></TD></TR></TABLE></BODY></HTML>");
}
//--------------------------------------------------------------

//生成语法加亮规则的选项列表
//selectLang: 默认选中的语言ID（默认选中第一项）
function FCSyntaxOptions(selectLang) {
	for (var i in FCSyntaxDef) {
		if (selectLang == null) selectLang == i;
		document.write('<OPTION value="' + i + '"' + (selectLang == i ? ' selected' : '')
			+ '>' + FCSyntaxDef[i].name + '</OPTION>');
	}
}
//--------------------------------------------------------------

//检测语法加亮规则定义集合
if (typeof(FCSyntaxDef) == "undefined") {
	FCSyntaxDef = {};
} else {
	FCCheckSyntaxDef();
}
//--------------------------------------------------------------
