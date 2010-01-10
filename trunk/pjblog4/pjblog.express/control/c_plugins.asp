<%
Public Sub c_plugins
%>
<style type="text/css">
.skin{width:700px; margin:0 auto; height:486px;}
.skin ul{ list-style:none;}
.skin ul li.item{ height:110px; margin:30px 0px;}
.skin ul li.item .zone{ height:110px}
.skin ul li.item .zone .left{ float:left; height:110px; background:url(../../images/Control/skin_first_bg.gif) no-repeat 0px 4px; width:150px; position:relative;}
.skin ul li.item .zone .left img{ position:absolute; width:86px; height:90px; top:9px; left:7px;}
.skin ul li.item .zone .right{ float:left; width:460px;}

.skin ul li.item .zone .right .set{ line-height:30px; height:40px; font-weight:700; color:#0F5FBB}
.skin ul li.item .zone .right .net{ height:50px; color:#999}
.skin ul li.item .zone .right .upload{ height:20px;}

</style>
	<div class="skin">
    	<ul>
        	<li class="item">
            	<div class="zone">
                	<div class="left">
                    	<img src="../../images/Control/plus_no.gif" />
                    </div>
                    <div class="right">
                    	<div class="set">未安装插件</div>
                        <div class="net">上传的插件都放在 ＂<strong style=" font-style:oblique">/pjblog.plugin/</strong>＂ 下, 可以通过本模块对这些插件进行安装.插件安装过程请耐心等待, 尽量避免刷新页面, 导致插件安装中断, 使程序无法运行.</div>
                        <div class="upload"><a href="">进入模块</a> </div>
                    </div>
                    <div class="clear"></div>
                </div>
            </li>
            <li class="item">
            	<div class="zone">
                	<div class="left">
                    	<img src="../../images/Control/plus_online.gif" />
                    </div>
                    <div class="right">
                    	<div class="set">在线安装插件</div>
                        <div class="net">网上所有官方认证审核的插件, 可以在线安装这些插件. 同时可以通过<strong>插件设置</strong>模块进入后发布自己的插件. 一经审核, 在所有博客程序上都能找到你的作品!</div>
                        <div class="upload"><a href="">进入模块</a> </div>
                    </div>
                    <div class="clear"></div>
                </div>
            </li>
            <li class="item">
            	<div class="zone">
                	<div class="left">
                    	<img src="../../images/Control/plus_set.gif" />
                    </div>
                    <div class="right">
                    	<div class="set">插件设置</div>
                        <div class="net">所有已安装的插件都可以在这里找到, 并且对其做相应的设置.同时也可以发布插件, 停止插件和卸载插件.</div>
                        <div class="upload"><a href="">进入模块</a> </div>
                    </div>
                    <div class="clear"></div>
                </div>
            </li>
        </ul>
    </div>
<%
End Sub
%>