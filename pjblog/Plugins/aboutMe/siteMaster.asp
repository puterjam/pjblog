<!--#include file="../../commond.asp" -->
<!--#include file="../../header.asp" -->
<!--#include file="../../plugins.asp" -->
<!--#include file="../../common/ModSet.asp" -->
<%' 如果需要用到UBB编辑器需要把 "../../common/UBBconfig.asp" 引用进来%>

 <!--内容-->
  <div id="Tbody">

   <div id="mainContent">
     <div id="innermainContent">
       <div id="mainContent-topimg"></div>
	   <%=content_html_Top%>
	   <%
	     Dim LoadModSet
	     Set LoadModSet=New ModSet
	     LoadModSet.open("AboutMeForPJBlog")
	     dim MyName,MyFace,Age,MyDay,Sex,MyBrood,MyStar,Address,MyDes,ILike
	     MyName=LoadModSet.getKeyValue("MyName")
	     MyFace=LoadModSet.getKeyValue("MyFace")
	     Age=LoadModSet.getKeyValue("Age")
	     MyDay=LoadModSet.getKeyValue("MyDay")
	     Sex=LoadModSet.getKeyValue("Sex")
	     MyBrood=LoadModSet.getKeyValue("MyBrood")
	     MyStar=LoadModSet.getKeyValue("MyStar")
	     Address=LoadModSet.getKeyValue("Address")
	     MyDes=LoadModSet.getKeyValue("MyDes")
	     ILike=LoadModSet.getKeyValue("ILike")
	   %>
	   <div id="Content_ContentList" class="content-width">
         <div class="Content">
         <div class="Content-top"><div class="ContentLeft"></div><div class="ContentRight"></div>
         <h1 class="ContentTitle">个人档案</h1>
         <h2 class="ContentAuthor">About Me</h2>
         </ul>
         </div>
         <div class="Content-body"><img src="<%=MyFace%>" alt="" border="0" align="left" style="margin-right:6px"/>
            <table border="0" cellspacing="0" cellpadding="0">
               <tr>
                <td class="commenttop" style="padding:2px;width:150px;"><nobr>Blog名称：</nobr></td><td  class="commenttop" style="padding:2px;"><%=SiteName%></td>
               </tr>
               <tr>
                <td class="commentcontent" style="padding:2px;width:150px;"><nobr>站长昵称：</nobr></td><td class="commentcontent" style="padding:2px;"><%=MyName%></td>
               </tr>
               <tr>
                <td class="commenttop" style="padding:2px;width:150px;"><nobr>年 龄：</nobr></td><td  class="commenttop" style="padding:2px;"><%=Age%></td>
               </tr>
               <tr>
                <td class="commentcontent" style="padding:2px;width:150px;"><nobr>生 日：</nobr></td><td class="commentcontent" style="padding:2px;"><%=MyDay%></td>
               </tr>
               <tr>
                <td class="commenttop" style="padding:2px;width:150px;"><nobr>性 别：</nobr></td><td  class="commenttop" style="padding:2px;"><%=Sex%></td>
               </tr>
               <tr>
                <td class="commentcontent" style="padding:2px;width:150px;"><nobr>血 型：</nobr></td><td class="commentcontent" style="padding:2px;"><%=MyBrood%></td>
               </tr>               
               <tr>
                <td class="commenttop" style="padding:2px;width:150px;"><nobr>星 座：</nobr></td><td  class="commenttop" style="padding:2px;"><%=MyStar%></td>
               </tr>
               <tr>
                <td class="commentcontent" style="padding:2px;width:150px;"><nobr>地 址：</nobr></td><td class="commentcontent" style="padding:2px;"><%=Address%></td>
               </tr>               
               <tr>
                <td class="commenttop" style="padding:2px;width:150px;"><nobr>个人说明：</nobr></td><td  class="commenttop" style="padding:2px;"><%=MyDes%></td>
               </tr>
               <tr>
                <td class="commentcontent" style="padding:2px;width:150px;"><nobr>兴趣爱好：</nobr></td><td class="commentcontent" style="padding:2px;"><%=ILike%></td>
               </tr>     
             </table>
         </div>
         <div class="Content-bottom">
         </div>
         </div>
</div>
	   <%=content_html_Bottom%>
       <div id="mainContent-bottomimg"></div>
   </div>
   </div>
   
   <div id="sidebar">
    <div id="innersidebar">
     <div id="sidebar-topimg"><!--工具条顶部图象--></div>
	  <%=side_html%>
     <div id="sidebar-bottomimg"></div>
   </div>
  </div>
 </div>
 <div style="font: 0px/0px sans-serif;clear: both;display: block"></div>
 <!--#include file="../../footer.asp" -->