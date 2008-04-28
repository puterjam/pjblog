/*************************************
*       XML Syntax Definition
*************************************/
FCSyntaxDef ["XML"] = {
	name		: "XML",
	delimiters	: "<?/>!=\"",
	blocks		: {
		BlockComment : {
			name	: "块注释",
			color	: "#008080",
			begin	: "<!--",
			end		: "-->",
			lines	: true
		},
		BlockText : {
			name	: "块文本",
			color	: "#808080",
			begin	: "<![CDATA[",
			end		: "]]>",
			lines	: true
		},
/*		TagHead : {
			name	: "标签头",
			color	: "#0000ff",
			style	: "b",
			begin	: "<",
			end		: " "
		},
		TagHead2 : {
			name	: "标签头",
			color	: "#0000ff",
			style	: "b",
			begin	: "<",
			end		: "\t"
		},*/
		String : {
			name	: "字符串",
			color	: "#ff00ff",
			begin	: "\"",
			end		: "\""
		}
	},
	keywords	: {
		Tag : {
			name	: "标签",
			color	: "#0000ff",
			style	: "b",
			list	: "< > ? / :"
		},
		Attribute : {
			name	: "属性",
			color	: "#008000",
			list	: "="
		}
	}
};
//--------------------------------------------------------------
FCCheckSyntaxDef("XML");