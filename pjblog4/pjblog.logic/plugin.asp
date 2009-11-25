<!--#include file = "../include.asp" -->
<!--#include file = "../pjblog.data/cls_modset.asp" -->
<!--#include file = "../pjblog.data/cls_plus.asp" -->
function fillSide(){
/* -------------------- pluginstart -------------------- */
/*---------------------------------plugin_Start_sys.bloginfo.ins.bloginfo------------------------------------*/
<!--#include file="../pjblog.plugin/bloginfo/ins_bloginfo.asp" -->
/*---------------------------------plugin_End_sys.bloginfo.ins.bloginfo------------------------------------*/
/*---------------------------------plugin_Start_sys.category.ins.category------------------------------------*/
<!--#include file="../pjblog.plugin/category/ins_category.asp" -->
/*---------------------------------plugin_End_sys.category.ins.category------------------------------------*/
/*---------------------------------plugin_Start_NewTopArticle.NewTopArticle------------------------------------*/
<!--#include file="../pjblog.plugin/NewTopArticle/NewTopArticle.asp" -->
/*---------------------------------plugin_End_NewTopArticle.NewTopArticle------------------------------------*/
/*---------------------------------plugin_Start_sys.userpannel.ins.userpannel------------------------------------*/
<!--#include file="../pjblog.plugin/userpannel/ins_userpannel.asp" -->
/*---------------------------------plugin_End_sys.userpannel.ins.userpannel------------------------------------*/
/* -------------------- pluginend -------------------- */
}
fillSide();
<%
Sys.Close
%>