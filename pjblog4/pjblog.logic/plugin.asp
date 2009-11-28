<!--#include file = "../include.asp" -->
<!--#include file = "../pjblog.data/cls_modset.asp" -->
<!--#include file = "../pjblog.data/cls_plus.asp" -->
function fillSide(){
/* -------------------- pluginstart -------------------- */
/*---------------------------------plugin_Start_NewTopArticle.NewTopArticle------------------------------------*/
<!--#include file="../pjblog.plugin/NewTopArticle/NewTopArticle.asp" -->
/*---------------------------------plugin_End_NewTopArticle.NewTopArticle------------------------------------*/
/*---------------------------------plugin_Start_sys.category.Sidecategory------------------------------------*/
<!--#include file="../pjblog.plugin/category/ins_category.asp" -->
/*---------------------------------plugin_End_sys.category.Sidecategory------------------------------------*/
/*---------------------------------plugin_Start_sys.userpannel.SideUserPannel------------------------------------*/
<!--#include file="../pjblog.plugin/userpannel/ins_userpannel.asp" -->
/*---------------------------------plugin_End_sys.userpannel.SideUserPannel------------------------------------*/
/* -------------------- pluginend -------------------- */
}
fillSide();
<%
Sys.Close
%>