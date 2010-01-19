/**
 * FontEffect - jQuery plugin for font effect
 *
 * @author Alessandro Uliana (fonteffect@iofo.it)
 *
 * Dual licensed under the MIT and GPL licenses:
 * http://www.opensource.org/licenses/mit-license.php
 * http://www.gnu.org/licenses/gpl.html
 *
 * Demo and examples on
 * http://www.iofo.it/jquery/fonteffect/
 *
 * @requires jQuery v1.3.2
 * @version: 1.0.0 - 30/3/2009
 */
;(function($){
    //* Global Variables *************************************************************
    var FE={};
    FE.divcounter=0; // div counter for debug
    //* Tabella posizioni layer *************************************************************
    FE.tabpos = ["", "0001021020212212", "00010203041020304041424344142434", "000102030405061020304050606162636465162636465666"];
    //* Most common typeface(from http: //www.codestyle.org/css/font-family/index.shtml) **********
    FE.font = {
        serif:       "Georgia, 'Times New Roman', 'Century Schoolbook L', serif",
        sans_serif:  "Verdana, Helvetica, Arial, 'URW Gothic L', sans-serif",
        monospace:   "'Courier New', Courier, 'DejaVu Sans Mono', monospace",
        fantasy:     "Impact, Papyrus, fantasy",
        cursive:     "'Comic Sans MS' cursive"
        };
    //* Main Function *******************************************************************************
    $.fn.FontEffect = function(o){
        //* Defaults *********************************************************************
        var d = $.extend({
            outline             :false,
            outlineColor1       :"",
            outlineColor2       :"",
            outlineWeight       :1,    // 1=light, 2=normal, 3=bold
            mirror              :false,
            mirrorColor         :"#000",
            mirrorOffset        :-10,
            mirrorHeight        :50,
            mirrorDetail        :1,    // 1=high, 2=medium, 3=low
            mirrorTLength       :50,
            mirrorTStart        :0.2,
            shadow              :false,
            shadowColor         :"#aaa",
            shadowOffsetTop     :5,
            shadowOffsetLeft    :5,
            shadowBlur          :1,    // 1=none, 2=low, 3=high
            shadowOpacity       :0.1,
            gradient            :false,
            gradientColor       :"",
            gradientFromTop     :true,
            gradientPosition    :20,
            gradientLength      :50,
            gradientSteps       :20,
            proportional        :false,
            hideText            :false,
            debug               :false
        },  o);
        //* Main Loop ********************************************************************
        this.not(".JQFE").each(function(){
            //* Check and correct options ********************************************************************
            if(!d.outline &&
                !d.shadow  &&
                !d.mirror  &&
                !d.gradient) {d.outline=true;};
            if(d.outline){
                if(d.outlineColor1 == "" && d.outlineColor2 == ""){
                    d.outlineColor1=pickcontrast($(this).css("color"));
                    };
                if(d.outlineColor2 == "") d.outlineColor2=d.outlineColor1;
                };
            if(d.gradient && d.gradientColor == ""){d.gradientColor=pickcontrast($(this).css("color"));};
            //* get the element display option and change to inline ********************************************************************
            var userdisplay=$(this).css("display");
            var userposition=$(this).css("position");
            $(this).css({
                display: "inline",
                position:((userposition == "absolute")?"absolute": "relative")
                });
            //* Local Variables ********************************************************************
            var h=$(this).height();
            var w=$(this).width()*1.04;
            var W=w+"px";
            var H=h+"px";
            //var W=w.pxToEm({scope: this});
            //var H=h.pxToEm({scope: this});
            var t=$(this).html();
            //* Set Class and Options ********************************************************
            $(this)
                .data("options", d)
                .addClass("JQFE")
                .css({
                    width: W,
                    height: H,
                    display: userdisplay,
                    position:(($(this).css("position")!= "absolute")?"relative": "absolute"),
                    zoom: 1
                    });
            //* Create the MyContainer structure ***********************************************
            var MyContainer=$("<div></div>").css({ // need extra div for IE zindex bug
                width: W,
                height: H,
                position: "relative"
            });
            //* MyContainer For the central layer ***********************************************
            MyContainer.append(
                $("<div class='JQFEText'>"+t+"</div>").css({
                    display:  d.hideText ? "none" :  "inline" ,
                    width: W,
                    height: H,
                    position: "relative",
                    zIndex:   100
                })
            );
            //* MyContainer For the Upper Effect layer ***********************************************
            var alldivsup=$("<div></div>").css({
                width: W,
                height: H,
                left: "0px",
                position: "absolute",
                top:      parseInt($(this).css("paddingTop"))*0+"px",
                zIndex:   110
            });
            //* MyContainer For the Lower Effect layer ***********************************************
            var alldivsdown=$(alldivsup).clone().css({zIndex: 90});
            FE.divounter+= 4;
            $(this).html("");
            //* Mirror Effect ****************************************************************
            if(d.mirror){
                for(i=0;i<h*(d.mirrorHeight/100);i++){
                    if(d.proportional){
                        var css_top1    =(h+d.mirrorOffset+i*d.mirrorDetail).pxToEm({scope: this});
                        var css_height  =d.mirrorDetail.pxToEm({scope: this});
                        var css_top2    =((h*-1)+i*(100/d.mirrorHeight)).pxToEm({scope: this});
                    }
                    else{
                        var css_top1    =(h+d.mirrorOffset+i*d.mirrorDetail)+"px";
                        var css_height  =d.mirrorDetail+"px";
                        var css_top2    =((h*-1)+i*(100/d.mirrorHeight))+"px";
                    };
                    var css_opacity=d.mirrorTStart-(i*(d.mirrorTStart/((d.mirrorHeight/100)*d.mirrorTLength)));
                    var appo=$("<div class='JQFEMirror'></div>").css({
                        position: "absolute",
                        top:      css_top1,
                        height:   css_height,
                        width:    W,
                        overflow: "hidden"
                    }).append($("<div>"+t+"</div>").css({
                            position: "absolute",
                            color:    d.mirrorColor,
                            top:      css_top2,
                            opacity:  css_opacity
                        })
                        );
                    FE.divounter+= i*2;
                    // Skip Non Visible Layers ************************************************
                    if(css_opacity<0.01) break;
                    alldivsdown.append(appo);
                };
            };
            //* Outline Effect ***************************************************************
            if(d.outline){
                var totdiv =(d.outlineWeight)*8;
                var to=FE.tabpos[d.outlineWeight];
                for(i=0;i<totdiv;i++){
                    appo=$("<div class='JQFEOutline'>"+t+"</div>").css({
                        position: "absolute",
                        top:     (to.charAt(i*2)  -d.outlineWeight)+"px",
                        left:    (to.charAt(i*2+1)-d.outlineWeight)+"px",
                        width:    W,
                        color:   ((i<totdiv/2+d.outlineWeight)?d.outlineColor1: d.outlineColor2),
                        zIndex:  ((i>totdiv-totdiv/3)?20: 30)
                    });
                    FE.divounter+= i;
                    alldivsdown.append(appo);
                };
            };
            //* Shadow Effect ****************************************************************
            if(d.shadow){
                var totdiv =(d.shadowBlur)*8;
                var to=FE.tabpos[d.shadowBlur];
                for(i=0;i<totdiv;i++){
                    appo=$("<div class='JQFEShadow'>"+t+"</div>").css({
                        opacity:  d.shadowOpacity,
                        position: "absolute",
                        top:     (to.charAt(i*2)  -d.shadowBlur)+d.shadowOffsetTop +"px",
                        left:    (to.charAt(i*2+1)-d.shadowBlur)+d.shadowOffsetLeft+"px",
                        width:    W,
                        height:   H,
                        color:    d.shadowColor,
                        zIndex:   10
                    });
                    FE.divounter+= i;
                    alldivsdown.append(appo);
                };
            };
            //* Gradient Effect *************************************************************
            if(d.gradient){
                var step    = Math.round((h*(d.gradientLength*0.01))/d.gradientSteps);
                var postop  = h*(d.gradientPosition*0.01);
                var opa     =(1/d.gradientSteps);
                var gcolor  = d.gradientColor;
                /*
                if(!d.gradientFromTop){
                    gcolor=$(this).css("color");
                    $(this).css("color", d.gradientColor);
                }
                */
                for(i=0;i<d.gradientSteps;i++){
                    if(d.proportional){
                        css_top1   = (((i == 0)?0: postop)+i*step).pxToEm({scope: this});
                        css_height = (((i == 0)?postop: 0)+step  ).pxToEm({scope: this});
                        css_top2   = ((((i == 0)?0: postop)+i*step)*-1).pxToEm({scope: this});
                    }
                    else{
                        css_top1   = (((i == 0)?0: postop)+i*step)+"px";
                        css_height = (((i == 0)?postop: 0)+step  )+"px";
                        css_top2   = ((((i == 0)?0: postop)+i*step)*-1)+"px";
                    };
                    appo=$("<div class='JQFEGradient'></div>").css({
                        position: "absolute",
                        top:      css_top1,
                        height:   css_height,
                        left:     "0px",
                        width:    W,
                        overflow: "hidden"
                    }).append($("<div>"+t+"</div>").css({
                            width:    "100%",
                            position: "absolute",
                            top:      css_top2,
                            color:    gcolor,
                            opacity:  1-opa*i
                        })
                        );
                    FE.divounter+= i*2;
                    alldivsup.append(appo);
                };
            };
            //* End Effects ******************************************************************
            MyContainer.append(alldivsdown);
            MyContainer.append(alldivsup);
            //* Draw Effect ******************************************************************
            $(this).append(MyContainer);
        });//* Main Loop End *******************************************************************************
        //* Internal Functions ******************************************************************************
        function hex2rgb(hexcolor){
            hexcolor=hexcolor.substring(1);
            if(hexcolor.length == 3) hexcolor=hexcolor.charAt(0)+hexcolor.charAt(0)+hexcolor.charAt(1)+hexcolor.charAt(1)+hexcolor.charAt(2)+hexcolor.charAt(2);
            var rgbcolor="rgb("+parseInt(hexcolor.substring(0, 2), 16)+", "+parseInt(hexcolor.substring(2, 4), 16)+", "+parseInt(hexcolor.substring(4, 6), 16)+")";
            return(rgbcolor);
        };
        function chkColorString(col){
            // test if "col" is a valid html color definition string(rgb(n, n, n)||#fff|#ffffff)
            return(/(#([0-9A-Fa-f]{3,6})\b)|(rgb\(\s*\b([0-9]|[1-9][0-9]|1[0-9][0-9]|2[0-4][0-9]|25[0-5])\b\s*,\s*\b([0-9]|[1-9][0-9]|1[0-9][0-9]|2[0-4][0-9]|25[0-5])\b\s*,\s*\b([0-9]|[1-9][0-9]|1[0-9][0-9]|2[0-4][0-9]|25[0-5])\b\s*\))|(rgb\(\s*(\d?\d%|100%)+\s*,\s*(\d?\d%|100%)+\s*,\s*(\d?\d%|100%)+\s*\))/.test(col));
        };
        function pickcontrast(col){ //(try to) find the contrasting color
            if(chkColorString(col)){
                col = col.toUpperCase();
                if(col.charAt(0) == "#") col=hex2rgb(col);
                var appo=col.substring(4, col.length-1).split(", ");
                var g=255-parseInt(appo[0]);
                var b=255-parseInt(appo[1]);
                var r=255-parseInt(appo[2]);
                col="rgb("+r+", "+g+", "+b+")";
            };
        return(col);
        };
        return this; // chain...
    }; //* Main Function End *******************************************************************************
    //* External Functions *******************************************************************************
    $.fn.changeOptionsFE = function(newoptions){
        if(this){
            var oldoptions=$(this).data("options") || {};
            $.extend(oldoptions, newoptions);
            $(this).data("options", oldoptions);
        };
    };
    $.fn.redrawFE = function(newoptions){
        if(this){
            if(newoptions) $(this).changeOptionsFE(newoptions);
            $(this).removeFE();
            $(this).FontEffect($(this).data("options"));
        };
    };
    $.fn.removeFE = function(removeoptions){
        if(this && $(this).hasClass("JQFE")){
            var t=$(this).find("div[class='JQFEText']").html();
            $(this).removeClass("JQFE");
            if(removeoptions) $(this).data("options", {});
            $(this).find("div[class^='JQFE']").remove();
            $(this).html(t);
        };
    };
    //* External Functions End *******************************************************************************
})(jQuery);
//* End FontEffect Plugin ******************************************************************************
/*--------------------------------------------------------------------
 * javascript method:  "pxToEm"
 * by:
   Scott Jehl(scott@filamentgroup.com)
   Maggie Wachs(maggie@filamentgroup.com)
   http: //www.filamentgroup.com
 *
 * Copyright(c) 2008 Filament Group
 * Dual licensed under the MIT(filamentgroup.com/examples/mit-license.txt) and GPL(filamentgroup.com/examples/gpl-license.txt) licenses.
 *
--------------------------------------------------------------------*/
Number.prototype.pxToEm = String.prototype.pxToEm = function(settings){
    settings = $.extend({
        scope:  'body',
        reverse:  false
},  settings);
    var pxVal =(this  ==  '') ? 0 :  parseFloat(this);
    var scopeVal;
    var getWindowWidth = function(){
        var de = document.documentElement;
        return self.innerWidth ||(de && de.clientWidth) || document.body.clientWidth;
};
    if(settings.scope  ==  'body' && $.browser.msie &&(parseFloat($('body').css('font-size')) / getWindowWidth()).toFixed(1) > 0.0){
        var calcFontSize = function(){
            return(parseFloat($('body').css('font-size'))/getWindowWidth()).toFixed(3) * 16;
    };
        scopeVal = calcFontSize();
}
    else { scopeVal = parseFloat($(settings.scope).css("font-size")); };
    var result =(settings.reverse  ==  true) ?(pxVal * scopeVal).toFixed(2) + 'px' : (pxVal / scopeVal).toFixed(2) + 'em';
    return result;
};
/*一，增加阴影(Shadow)实例

1，javascript部分：$('#example1').FontEffect({ shadow:true })

2，CSS部分：#example1{ font-family :"Impact"; color :#000; font-size :3em; }

3，HTML部分：<DIV id='example1'> Shadow </DIV>

二，增加渐变(Gradient )实例

1，javascript部分：$('#example2').FontEffect({ gradient:true })

2，CSS部分：#example2{ font-family :"Impact"; color :#000; font-size :3em; }

3，HTML部分：<DIV id='example2'>  Gradient </DIV>

三，增加反射(Mirror)实例

1，javascript部分：$('#example3').FontEffect({ mirror:true })

2，CSS部分：#example3{ font-family :"Impact"; margin-bottom:30px; color :#000; font-size :3em; }

3，HTML部分：<DIV id='example3'>  Mirror </DIV>

附Font effect参数清单(可自定义设置相关效果)

outline :false, 
outlineColor1 :"", 
outlineColor2 :"", 
outlineWeight :1, 

mirror :false, 
mirrorColor :"",
mirrorOffset :-10, 
mirrorHeight :50, 
mirrorDetail :1,
mirrorTLength :50, 
mirrorTStart :0.2, 

shadow :false, 
shadowColor :"#000", 
shadowOffsetTop :5, 
shadowOffsetLeft:5, 
shadowBlur :1, 
shadowOpacity :0.1, 

gradient :false, 
gradientColor :"", 
gradientPosition:20, 
gradientLength :50, 
gradientSteps :20, 
hideText :false*/