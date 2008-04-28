/*************************************
*     Document Syntax Definition
*************************************/
FCSyntaxDef ["DOC"] = {
	name		: "Document",
	delimiters	: "~!%^&*()-+=|\\/{}[]:;\"'<>,.?◎－·（）：",
	comments	: "----- *****",
	cmtcolor	: "#008080",
	blocks		: {
		Title : {
			name	: "标题",
			color	: "#A80000",
			style	: "bu",
			begin	: "◎",
			end		: "◎"
		},
		Title2 : {
			name	: "2级标题",
			color	: "#006600",
			style	: "b",
			begin	: "－",
			end		: "："
		},
		Title3 : {
			name	: "3级标题",
			color	: "#5050A0",
			begin	: "·",
			end		: "："
		},
		Explanation : {
			name	: "附加解说",
			color	: "#800080",
			begin	: "（",
			end		: "）",
			lines	: true
		},
		CodeBlock : {
			name	: "代码段",
			color	: "#0000FF",
			begin	: "//--code",
			end		: "//--code",
			lines	: true
		},
		HttpLink : {
			name	: "Http链接",
			color	: "#0080C0",
			begin	: "://",
			end		: " "
		},
		FtpLink : {
			name	: "Ftp链接",
			color	: "#0080C0",
			begin	: "ftp://",
			end		: " "
		},
		EmailLink : {
			name	: "电子邮件链接",
			color	: "#804000",
			begin	: "mailto:",
			end		: " "
		}
	}
};
//--------------------------------------------------------------
FCCheckSyntaxDef("DOC");