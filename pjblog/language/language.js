// -----------------------------------------------
//	Language Jscript For PJBlog Write By evio
//	Http://www.evio.name & Http://www.pjhome.net
// 	2009-09-19 13:15
//	CopyRight @ PJBlog Http://www.pjhome.net
// -----------------------------------------------
var lang = {
	/*
		Err Class. Deal with error description!
	*/
	Err : {
		Conn : "数据库连接出错，请检查连接字串!",
		info : function(t){
			var temp;
			switch (t)
			{
				case 1 : temp = "用户名已注册"; break;
				case 2 : temp = "用户名未注册"; break;
				case 3 : temp = "非法参数"; break;
				default : temp = "非法操作"; // 999 号错误
			}
			return temp;
		}
	},
	/*
		Action Class.
	*/
	Action : {
		Login : "登入",
		Logout : "退出系统",
		Register : "注册新用户"
	},
	
	/*
		Tip Class.
	*/
	Tip : {}
}