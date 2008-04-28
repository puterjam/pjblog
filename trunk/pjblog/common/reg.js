//正则表达过滤
function rgExec(text,reg,times){
	if (!reg) return -3 //reg is null
	try{
		var re = new RegExp(reg,"ig");
	}catch(e){
		return -1 //regExp error.
	}

	if (typeof text == "string"){
		var tm = text.match(re);
		if (tm && times<=tm.length){
			return 1
		}
	}else{
		return -2 //not string
	}
	return 0
}