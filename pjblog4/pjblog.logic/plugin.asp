<!--#include file = "../include.asp" -->
function fillSide(){
/* -------------------- pluginstart -------------------- */
/*---------------------------------plugin_Start_ins.category------------------------------------*/
<!--#include file="../pjblog.plugin/category/ins_category.asp" -->
/*---------------------------------plugin_End_ins.category------------------------------------*/
/*---------------------------------plugin_Start_ins.userpannel------------------------------------*/
<!--#include file="../pjblog.plugin/userpannel/ins_userpannel.asp" -->
/*---------------------------------plugin_End_ins.userpannel------------------------------------*/
/*---------------------------------plugin_Start_ins.bloginfo------------------------------------*/
<!--#include file="../pjblog.plugin/bloginfo/ins_bloginfo.asp" -->
/*---------------------------------plugin_End_ins.bloginfo------------------------------------*/
/* -------------------- pluginend -------------------- */
}
fillSide();
<%
Sys.Close
%>