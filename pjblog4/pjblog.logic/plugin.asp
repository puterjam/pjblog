<!--#include file = "../include.asp" -->
<!--#include file = "../pjblog.data/cls_modset.asp" -->
<!--#include file = "../pjblog.data/cls_plus.asp" -->
function fillSide(){
/* -------------------- pluginstart -------------------- */
/*---------------------------------plugin_Start_NewTopArticle.NewTopArticle------------------------------------*/
<!--#include file="../pjblog.plugin/NewTopArticle/NewTopArticle.asp" -->
/*---------------------------------plugin_End_NewTopArticle.NewTopArticle------------------------------------*/
/* -------------------- pluginend -------------------- */
}
fillSide();
<%
Sys.Close
%>