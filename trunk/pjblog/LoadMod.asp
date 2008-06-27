<!--#include file="const.asp" -->
<!--#include file="conn.asp" -->
<!--#include file="common/function.asp" -->
<!--#include file="common/library.asp" -->
<!--#include file="common/cache.asp" -->
<%
'==================================
'  插件读取页面
'    更新时间: 2006-5-28
'==================================
log_module(1)
Call TransferMod '装载插件
Call CloseDB '关闭数据库

%>
<script Language="JScript" runat="server">
//检查插件是否成功安装
	function checkplugins(pl){
	    var pins = function_Plugin.split("$*$");
	    for (var i=0;i<pins.length;i++){
	      var plugItems = pins[i].split("%|%");
	      if (pl == plugItems[0]) return plugItems
	    }
	    return false
	}
  
//装载插件
	function TransferMod(){
		try{
			getPlugins = Request.QueryString("plugins");
			if (!getPlugins)
				Response.Write('<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" /><div style="font-size:12px;font-weight:bold;border:1px solid #006;padding:6px;background:#fcc">非法参数</div>');
		    
		    var loadMod = checkplugins(getPlugins);
			if (loadMod) {
			    var path = "Plugins/" + loadMod[2] + "/" + loadMod[1];
				Server.Transfer(path);
			}else{
			  	Response.Write('<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" /><div style="font-size:12px;font-weight:bold;border:1px solid #006;padding:6px;background:#fcc">无法找到相应的模块!</div>');
			}
		}catch(e){
		  	Response.Write('<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" /><div style="font-size:12px;font-weight:bold;border:1px solid #006;padding:6px;background:#fcc">加载模块发生异常!</div>');
		}
	}
</script>
