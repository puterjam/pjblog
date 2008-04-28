<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/TR/WD-xsl">
	<xsl:template match="/">
		<HTML xmlns:xsl="http://www.w3.org/TR/WD-xsl">
		<HEAD>
		<TITLE><xsl:value-of select="response/rss/channel/title"/></TITLE>
		<style>
		 body{font-size:12px;text-align:center;margin:10px;background:#DFE4EA;FILTER: progid:DXImageTransform.Microsoft.Gradient(gradientType=0,startColorStr=#F7F8F9,endColorStr=#E3E7EA);}
		 h1{font-size:16px;text-align:left;padding:5px;}
		 h3{font-size:14px;margin:3px;background:url(images/icon_trackback.gif) no-repeat;padding-left:20px;}
		 a:link,a:visited{color:#47586D}
		 a:hover{color:#5F7692}
		 .main{margin:auto;text-align:left;background:#fff;padding:3px;padding-top:6px;border:1px solid #47586D;width:99%;FILTER: progid:DXImageTransform.Microsoft.Shadow(direction=135,color=#999999,strength=3);}
		 .content {padding:5px 5px 10px 25px;border-bottom:1px dotted #47586D}
		 .info{padding:4px;text-align:right;color:#666}
		 .day{font-size:11px;color:#666;font-weight:100}
		</style>
		</HEAD>
		<BODY>
			  <xsl:apply-templates select="response/error"/>
		</BODY>
		</HTML>
	</xsl:template>
	
	
	<xsl:template match="response/error">
		<xsl:choose> 
		    <xsl:when test=".[value()$eq$0]">
				    <h1><a>
					    <xsl:attribute name="href"><xsl:value-of select='/response/rss/channel/link'/></xsl:attribute>
					    <xsl:value-of select="/response/rss/channel/title"/> - Trackback List
				    </a></h1>
					<div class="main">
						      <xsl:apply-templates select="/response/rss/channel"/>
					<div class="info"><a><xsl:attribute name="href"><xsl:value-of select='/response/rss/channel/link'/></xsl:attribute>
								    <img src="images/urlInTop.gif" border="0" alt="返回" style="margin:0 2px -3px 0"/> 返回日志</a></div>
					</div>
			</xsl:when>
		    <xsl:otherwise>
		       <div class="main"><b><img src="images/ico_skdaq.gif" border="0" alt="返回" style="margin:0 2px -2px 0"/> <xsl:value-of select='//message'/></b></div>
		    </xsl:otherwise>
		</xsl:choose>
	</xsl:template>
	
	<xsl:template match="/response/rss/channel">
		<xsl:for-each select="item">
			<div>
				<h3><a>
					<xsl:attribute name="href"><xsl:value-of select='link'/></xsl:attribute>
					<xsl:attribute name="target">_blank</xsl:attribute>
					<xsl:value-of select="title"/></a> <span class="day"><xsl:value-of select="pubDate"/></span></h3>
				<div class="content"><xsl:value-of select="description"/></div>
			</div>
		</xsl:for-each>
	</xsl:template>
	
</xsl:stylesheet>