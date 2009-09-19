<script Language="JScript" runat="server">
//***************PJblog2 连接数据库*******************
// PJblog2 Copyright 2007
// Update:2007-5-28
//***************************************************

	function createConnection(strPath){
		try{
			Conn = Server.CreateObject("ADODB.Connection");
			Conn.connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Server.MapPath(strPath);
			Conn.open();
		}catch(e){
		    Response.Write('<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" /><div style="font-size:12px;font-weight:bold;border:1px solid #006;padding:6px;background:#fcc">' + lang.Err_Conn + '</div>');
		    Response.End
		}
	}
</script>
<%
Call createConnection(AccessFile)
%>
