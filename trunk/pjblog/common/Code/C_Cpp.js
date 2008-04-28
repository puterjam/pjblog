/*************************************
*      C/C++ Syntax Definition
*************************************/
FCSyntaxDef ["CPP"] = {
	name		: "C/C++",
	delimiters	: "~!%^&*()-+=|\\/{}[]:;\"'<>,.?",
	comments	: "//",
	cmtcolor	: "#008080",
	blocks		: {
		String : {
			name	: "×Ö·û´®",
			color	: "#ff00ff",
			begin	: "\"",
			end		: "\"",
			escape	: "\\"
		},
		Char : {
			name	: "×Ö·û",
			color	: "#4080ff",
			begin	: "'",
			end		: "'",
			escape	: "\\"
		},
		BlockComment : {
			name	: "¿é×¢ÊÍ",
			color	: "#008080",
			begin	: "/*",
			end		: "*/",
			lines	: true
		}
	},
	keywords	: {
		Operator : {
			name	: "ÔËËã·û",
			color	: "#0000ff",
			list	: "+ - = / % & > < ^ ! | : [ ] { } ( ) , ; . ~ * ?"
		},
		Keyword : {
			name	: "Óï·¨¹Ø¼ü×Ö",
			color	: "#800000",
			style	: "b",
			list	: "auto break class catch case continue const default do delete else extern "
					+ "for friend goto inline if new namespace operator private protected public "
					+ "return sizeof switch static this try template throw typedef virtual while "
		},
		VarType : {
			name	: "±äÁ¿ÀàÐÍ¹Ø¼ü×Ö",
			color	: "#ff8000",
			style	: "b",
			list	: "char double enum float int long register short signed struct "
					+ "union unsigned void volatile "
		},
		Compiler : {
			name	: "±àÒë¹Ø¼ü×Ö",
			color	: "#008000",
			style	: "b",
			list	: "#define #error #include #elif #if #line #else #ifdef #pragma #endif #ifndef "
					+ "#undef defined __asm __fastcall __self __segment __based __segname __fortran "
					+ "__cdecl __huge __far __saveregs __export __pascal __near __loadds __interrupt "
					+ "__inline __multiple_inheritance __single_inheritance __virtual_inheritance "
		}
	}
};
//--------------------------------------------------------------
FCCheckSyntaxDef("CPP");