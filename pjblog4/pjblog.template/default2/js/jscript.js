$(function(){
  $("a").bind("focus",function(){if(this.blur)this.blur();});
 
  $("#menu > li:first-child").addClass("first_menu");
  $("#menu > li:first-child.current_page_item, #menu > li:first-child.current-cat").addClass("first_menu_active");
  $("#menu > li:last-child").addClass("last_menu");
  $("#menu > li:last-child.current_page_item, #menu > li:last-child.current-cat").addClass("last_menu_active");
  $("#menu li ul li:has(ul)").addClass("parent_menu");

  $("#right_col ul > li:last-child").css({marginBottom:"0"});

  $("#comment_area ol > li:even").addClass("even_comment");
  $("#comment_area ol > li:odd").addClass("odd_comment");
  $(".even_comment > .children > li").addClass("even_comment_children");
  $(".odd_comment > .children > li").addClass("odd_comment_children");
  $(".even_comment_children > .children > li").addClass("odd_comment_children");
  $(".odd_comment_children > .children > li").addClass("even_comment_children");
  $(".even_comment_children > .children > li").addClass("odd_comment_children");
  $(".odd_comment_children > .children > li").addClass("even_comment_children");

  $("#guest_info input,#comment_textarea textarea")
    .focus( function () { $(this).css({borderColor:"#33a8e5"}) } )
    .blur( function () { $(this).css({borderColor:"#ccc"}) } );

  $(".search_tag").toggle(function(){
    $(".search_tag").addClass("active_search_tag");
    $("#tag_list ul").animate({ opacity: 'show', height:"show" }, {duration:1000 ,easing:"easeOutBack"});
    return false;
  },function(){
    $("#tag_list ul").fadeTo("slow", 0).slideUp("slow").fadeTo("fast", 1);
    $(".search_tag").removeClass("active_search_tag")
    return false;
  });

  $("#trackback_switch").click(function(){
    $("#comment_switch").removeClass("comment_switch_active");
    $(this).addClass("comment_switch_active");
    $("#comment_area").animate({opacity: 'hide'}, 0);
    $("#trackback_area").animate({opacity: 'show'}, 1000);
    return false;
  });

  $("#comment_switch").click(function(){
    $("#trackback_switch").removeClass("comment_switch_active");
    $(this).addClass("comment_switch_active");
    $("#trackback_area").animate({opacity: 'hide'}, 0);
    $("#comment_area").animate({opacity: 'show'}, 1000);
    return false;
  });

  $("#guest_info input,#comment_textarea textarea")
    .focus( function () { $(this).css({borderColor:"#33a8e5"}) } )
    .blur( function () { $(this).css({borderColor:"#ccc"}) } );

  $("blockquote").append('<div class="quote_bottom"></div>');

  $(".side_box:first").addClass("first_side_box");

});


function changefc(color){
  document.getElementById("search_input").style.color=color;
}


/*
dropdowm menu
*/

var menu=function(){
	var t=15,z=50,s=6,a;
	function dd(n){this.n=n; this.h=[]; this.c=[]}
	dd.prototype.init=function(p,c){
		a=c; var w=document.getElementById(p), s=w.getElementsByTagName('ul'), l=s.length, i=0;
		for(i;i<l;i++){
			var h=s[i].parentNode; this.h[i]=h; this.c[i]=s[i];
			h.onmouseover=new Function(this.n+'.st('+i+',true)');
			h.onmouseout=new Function(this.n+'.st('+i+')');
		}
	}
	dd.prototype.st=function(x,f){
		var c=this.c[x], h=this.h[x], p=h.getElementsByTagName('a')[0];
		clearInterval(c.t); c.style.overflow='hidden';
		if(f){
			p.className+=' '+a;
			if(!c.mh){c.style.display='block'; c.style.height=''; c.mh=c.offsetHeight; c.style.height=0}
			if(c.mh==c.offsetHeight){c.style.overflow='visible'}
			else{c.style.zIndex=z; z++; c.t=setInterval(function(){sl(c,1)},t)}
		}else{p.className=p.className.replace(a,''); c.t=setInterval(function(){sl(c,-1)},t)}
	}
	function sl(c,f){
		var h=c.offsetHeight;
		if((h<=0&&f!=1)||(h>=c.mh&&f==1)){
			if(f==1){c.style.filter=''; c.style.opacity=1; c.style.overflow='visible'}
			clearInterval(c.t); return
		}
		var d=(f==1)?Math.ceil((c.mh-h)/s):Math.ceil(h/s), o=h/c.mh;
		c.style.opacity=o; c.style.filter='alpha(opacity='+(o*100)+')';
		c.style.height=h+(d*f)+'px'
	}
	return{dd:dd}
}();
