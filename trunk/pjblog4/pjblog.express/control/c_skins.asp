<%
Public Sub c_skins
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
                    	<img src="../../images/Control/skin_set.gif" />
                    </div>
                    <div class="right">
                    	<div class="set">主题设置</div>
                        <div class="net">选择主题和样式, 导出主题, 发布主题. 同时你可以为每套主题制定不同的插件支持和选择不同的风格样式.</div>
                        <div class="upload"><a href="?m=skin&s=set">进入模块</a> </div>
                    </div>
                    <div class="clear"></div>
                </div>
            </li>
            <li class="item">
            	<div class="zone">
                	<div class="left">
                    	<img src="../../images/Control/skin_net.gif" />
                    </div>
                    <div class="right">
                    	<div class="set">在线安装主题</div>
                        <div class="net">网上所有官方认证审核的主题, 可以在线安装这些主题. 同时可以通过<strong>主题设置</strong>模块进入后发布自己的主题. 一经审核, 在所有博客程序上都能找到你的作品!</div>
                        <div class="upload"><a href="?m=skin&s=net">进入模块</a> </div>
                    </div>
                    <div class="clear"></div>
                </div>
            </li>
            <li class="item">
            	<div class="zone">
                	<div class="left">
                    	<img src="../../images/Control/skin_upload.gif" />
                    </div>
                    <div class="right">
                    	<div class="set">主题导入(PBD格式文传)</div>
                        <div class="net">不需要通过FTP进行主题上传, 上传的文件格式(PBD)将被解压后释放. 同时你可以即时管理你的PBD文件, 进行删除或者添加等应用.</div>
                        <div class="upload"><a href="?m=skin&s=upd">进入模块</a> </div>
                    </div>
                    <div class="clear"></div>
                </div>
            </li>
        </ul>
    </div>
<%
End Sub
%>