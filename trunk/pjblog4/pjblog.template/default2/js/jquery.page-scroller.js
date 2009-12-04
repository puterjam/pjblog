/*  ================================================================================
 *
 *  jquery.page-scroller  -version 3.0.6
 *  http://coliss.com/articles/build-websites/operation/javascript/296.html
 *  (c) 2004-2009 coliss.com
 *  License:http://creativecommons.org/licenses/by/2.1/jp/
 *
================================================================================  */

/*  ================================================================================
TOC
============================================================
Set Adjustment
Page Scroller
============================================================
this script requires jQuery 1.3(http://jquery.com/)
================================================================================  */


/*  ================================================================================
Set Adjustment
================================================================================  */
var virtualTopId = "top",
    virtualTop,
    adjTraverser,
    adjPosition,
    callExternal = "pSc",
    delayExternal= 200;

/* example
======================================================================  */
//    virtualTop = 0;    // virtual top's left position = 0
//    virtualTop = 1;    // virtual top's left position = vertical movement
//    adjTraverser = 0;  // left position = 0
//    adjTraverser = 1;  // horizontal movement.
//    adjPosition = -26;

/*  ================================================================================
Page Scroller
================================================================================  */
eval(function(p,a,c,k,e,d){e=function(c){return(c<a?"":e(parseInt(c/a)))+((c=c%a)>35?String.fromCharCode(c+29):c.toString(36))};if(!''.replace(/^/,String)){while(c--)d[e(c)]=k[c]||e(c);k=[function(e){return d[e]}];e=function(){return'\\w+'};c=1;};while(c--)if(k[c])p=p.replace(new RegExp('\\b'+e(c)+'\\b','g'),k[c]);return p;}('(c($){7 D=$.E.D,C=$.E.C,G=$.E.G,A=$.E.A;$.E.1Q({C:c(){3(!6[0])1g();3(6[0]==i)b 1b.1P||$.1q&&5.B.1e||5.f.1e;3(6[0]==5)b((5.B&&5.1p=="1l")?5.B.1i:5.f.1i);b C.1n(6,1o)},D:c(){3(!6[0])1g();3(6[0]==i)b 1b.1T||$.1q&&5.B.1v||5.f.1v;3(6[0]==5)b((5.B&&5.1p=="1l")?5.B.1m:5.f.1m);b D.1n(6,1o)},G:c(){3(!6[0])b 11;7 k=5.M?5.M(6[0].z):5.1t(6[0].z);7 j=1u 1r();j.x=k.1j;1s((k=k.1a)!=12){j.x+=k.1j}3((j.x*0)==0)b(j.x);g b(6[0].z)},A:c(){3(!6[0])b 11;7 k=5.M?5.M(6[0].z):5.1t(6[0].z);7 j=1u 1r();j.y=k.19;1s((k=k.1a)!=12){j.y+=k.19}3((j.y*0)==0)b(j.y);g b(6[0].z)}})})(1Y);$(c(){$(\'a[F^="#"]\').1d(c(){7 h=R.21+R.20;7 H=((6.F).1Z(0,(((6.F).18)-((6.X).18)))).Q((6.F).1f("//")+2);3(h.I("?")!=-1)Y=h.Q(0,(h.I("?")));g Y=h;3(H.I("?")!=-1)Z=H.Q(0,(H.I("?")));g Z=H;3(Z==Y){d.V((6.X).1V(1));b 1O}});$("f").1d(c(){d.O()})});6.q=12;7 d={14:c(w){3(w=="x")b(($(5).C())-($(i).C()));g 3(w=="y")b(($(5).D())-($(i).D()))},13:c(w){3(w=="x")b(i.17||5.f.t||5.f.J.t);g 3(w=="y")b(i.1R||5.f.1J||5.f.J.1J)},S:c(l,m,v,p,o){7 q;3(q)P(q);7 1F=16;7 L=d.13(\'x\');7 N=d.13(\'y\');3(!l||l<0)l=0;3(!m||m<0)m=0;3(!v)v=$.1I.1N?10:$.1I.1W?8:9;3(!p)p=0+L;3(!o)o=0+N;p+=(l-L)/v;3(p<0)p=0;o+=(m-N)/v;3(o<0)o=0;7 U=u.1z(p);7 T=u.1z(o);i.1X(U,T);3((u.1A(u.1w(L-l))<1)&&(u.1A(u.1w(N-m))<1)){P(6.q);i.1x(l,m)}g 3((U!=l)||(T!=m))6.q=1B("d.S("+l+","+m+","+v+","+p+","+o+")",1F);g P(6.q)},O:c(){P(6.q)},1K:c(e){d.O()},V:c(n){d.O();7 r,s;3(!!n){3(n==1L){r=(K==0)?0:(K==1)?i.17||5.f.t||5.f.J.t:$(\'#\'+n).G();s=((K==0)||(K==1))?0:$(\'#\'+n).A()}g{r=(1C==0)?0:(1C==1)?($(\'#\'+n).G()):i.17||5.f.t||5.f.J.t;s=1E?($(\'#\'+n).A())+1E:($(\'#\'+n).A())}7 15=d.14(\'x\');7 W=d.14(\'y\');3(((r*0)==0)||((s*0)==0)){7 1G=(r<1)?0:(r>15)?15:r;7 1y=(s<1)?0:(s>W)?W:s;d.S(1G,1y)}g R.X=n}g d.S(0,0)},1c:c(){7 h=R.F;7 1H=h.1f("#",0);7 1h=h.1M(1k);3(!!1h){1D=h.Q(h.I("?"+1k)+4,h.18);1S=1B("d.V(1D)",1U)}3(!1H)i.1x(0,0);g b 11}};$(d.1c);',62,126,'|||if||document|this|var||||return|function|coliss||body|else|usrUrl|window|tagCoords|obj|toX|toY|idName|frY|frX|pageScrollTimer|anchorX|anchorY|scrollLeft|Math|frms|type|||id|top|documentElement|width|height|fn|href|left|anchorPath|lastIndexOf|parentNode|virtualTop|actX|getElementById|actY|stopScroll|clearTimeout|slice|location|pageScroll|posY|posX|toAnchor|dMaxY|hash|usrUrlOmitQ|anchorPathOmitQ||true|null|getWindowOffset|getScrollRange|dMaxX||pageXOffset|length|offsetTop|offsetParent|self|initPageScroller|click|clientWidth|indexOf|error|checkPageScroller|scrollWidth|offsetLeft|callExternal|CSS1Compat|scrollHeight|apply|arguments|compatMode|boxModel|Object|while|all|new|clientHeight|abs|scroll|setY|ceil|floor|setTimeout|adjTraverser|anchorId|adjPosition|spd|setX|checkAnchor|browser|scrollTop|cancelScroll|virtualTopId|match|mozilla|false|innerWidth|extend|pageYOffset|timerID|innerHeight|delayExternal|substr|opera|scrollTo|jQuery|substring|pathname|hostname'.split('|'),0,{}))