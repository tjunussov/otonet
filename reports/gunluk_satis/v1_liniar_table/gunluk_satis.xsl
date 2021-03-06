<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:exslt="http://exslt.org/common" exclude-result-prefixes="exslt msxsl">
	
	<xsl:output method="html" />

	<xsl:strip-space elements="*" />
    
    <xsl:decimal-format name="euro" infinity="0" decimal-separator="," NaN="0" grouping-separator="&#160;"/>
    
    <xsl:variable name="AylarDefinitionXML">
    	<root>
        	<ay name="01">OCAK</ay>
            <ay name="02">SUBAT</ay>
            <ay name="03">MART</ay>
            <ay name="04">NISAN</ay>
            <ay name="05">MAYIS</ay>
            <ay name="06">HAZIRAN</ay>
            <ay name="07">TEMMUZ</ay>
            <ay name="08">AGUSTOS</ay>
            <ay name="09">EYLUL</ay>
            <ay name="10">EKIM</ay>
            <ay name="11">KASIM</ay>
            <ay name="12">ARALIK</ay>
        </root>
    </xsl:variable>
    
    <xsl:variable name="AylarDefinition"  select="exslt:node-set($AylarDefinitionXML)/root[1]/*"/>
    
    <xsl:variable name="AylarXML">
    	<root>
            <xsl:variable name="ROWS" select="/REPORT/BIRIM[1]/ROWSET/ROW"/>
            <xsl:for-each select="/REPORT/BIRIM[1]/ROWSET">
                <xsl:for-each select="ROW[not(TARIH_AY=preceding-sibling::ROW/TARIH_AY)]">
                    <xsl:variable name="TARIH_AY" select="TARIH_AY"/>
                    <xsl:variable name="ROW" select="$ROWS[TARIH_AY = $TARIH_AY]"/>
                    <xsl:element name="AY">
                    	<xsl:attribute name="name"><xsl:value-of select="$TARIH_AY"/></xsl:attribute>
                    	<xsl:attribute name="title"><xsl:value-of select="$AylarDefinition[@name=$TARIH_AY]"/></xsl:attribute>
                        <xsl:for-each select="$ROW[not(TARIH_TEXT=preceding::TARIH_TEXT)]">
                            <TARIH_TEXT><xsl:value-of select="TARIH_TEXT"/></TARIH_TEXT>
                        </xsl:for-each>
                    </xsl:element>
                </xsl:for-each>
            </xsl:for-each>
        </root>
    </xsl:variable>
    
    <xsl:variable name="Aylar"  select="exslt:node-set($AylarXML)/root[1]/*"/>
            
    <xsl:key name="dates" match="/REPORT/BIRIM/ROWSET[1]/ROW" use="TARIH_AY" />
    
    <xsl:template match="/REPORT">
    <xsl:text disable-output-escaping='yes'>&lt;!DOCTYPE html&gt;</xsl:text>
    <html>
    	<head>
	    <link href="excel.css" type="text/css" rel="stylesheet"/>
    	</head>
    <body>
        <h2>RAMSTORE KAZAKISTAN - Gunluk Satis</h2>
        
        <l>Magazalar</l><v><xsl:value-of select="@birim_no"/></v><br />
        <l>Doviz</l>
        <v>
        	<xsl:choose>
            	<xsl:when test="@doviz_kod = '02'">KZT</xsl:when>
                <xsl:when test="@doviz_kod = '01'">USD</xsl:when>
                <xsl:otherwise><xsl:value-of select="@doviz_kod"/></xsl:otherwise>
            </xsl:choose>
        </v><br />
        <l>Donem</l><v><xsl:value-of select="@donem_from"/> - <xsl:value-of select="@donem_to"/></v><br /><br />
        
        <!--xsl:for-each select="$Aylar">
         	<xsl:value-of select="@name"/>   <xsl:value-of select="@title"/>   <br />

         </xsl:for-each-->
        
        <div class="magazas">
	    	<xsl:apply-templates select="BIRIM[1]/ROWSET" mode="dates"/>
            <xsl:apply-templates select="." mode="toplam"/>
	    	<xsl:apply-templates select="BIRIM/ROWSET"/>
        </div>
        
        <xsl:call-template name="magaza-toplam"/>
        <xsl:call-template name="toptan-satis"/>
        <xsl:call-template name="sirket-toplam"/>
        <xsl:call-template name="sirket-toplam-kumule"/>
        
        <div style="height:200px; clear:both;"></div>
        
        
        <xsl:call-template name="footer"/>
        
        <!--script src="https://ajax.googleapis.com/ajax/libs/jquery/1.12.0/jquery.min.js"></script>
        <script>
		
		
		$(function() {
			var top = $('.tarih').position().top;
			$(window).scroll(function(){
			  $('.tarih').css('top',(top-$(window).scrollTop())+'px');
			});
		});
		</script-->
    </body>
    </html>
    </xsl:template>
    
    <xsl:template match="ROWSET" mode="dates">
    	<div class="tarih">
    	<table>
        	<thead>
    		<tr class="head">
    			<th>&#160;</th>
    		</tr>
    		<tr class="head2">
    			<th>TARIH</th>
    		</tr>
            </thead>
            <xsl:for-each select="$Aylar">
            <tbody>
	            <xsl:for-each select="TARIH_TEXT">
                <tr>
                    <td><xsl:value-of select="."/></td>
                </tr>
            	</xsl:for-each>
            </tbody>
            <tbody class="toplam">
                <tr>
                    <th><xsl:value-of select="@title"/></th>
                </tr>
            </tbody>
    		</xsl:for-each>
            <tfoot class="toplam">
    		<tr>
    			<th>TOPLAM</th>
    		</tr>
            </tfoot>
    	</table>
        </div>
    </xsl:template>
    
    <xsl:template match="REPORT" mode="toplam">
    	<div>
    	<table class="magaza">
        	<thead>
    		<tr class="head">
    			<th colspan="8">MAGAZA TOPLAM</th>
    		</tr>
    		<tr class="head2">
    			<th>NET SATIŞ GEÇEN YIL</th>
    			<th>NET SATIŞ BÜTÇE</th>
    			<th>NET SATIS FİİLİ</th>
    			<th>Prog.Göre Artış</th>
                <th>G.Yıla Göre Artış</th>
                <th>MUSTERI SAYISI</th>
                <th>MUSTERI SAYISI GEÇEN YIL</th>
                <th>G.Yıla Göre Artış</th>
    		</tr>
            </thead>
            
            <xsl:variable name="ROWS" select="BIRIM/ROWSET/ROW"/>
            <xsl:for-each select="$Aylar">
                    
                    	<xsl:variable name="TARIH_AY" select="@name"/>
                    	<xsl:variable name="ROW" select="$ROWS[TARIH_AY = $TARIH_AY]"/>
                        
                     <tbody>
                        <xsl:for-each select="TARIH_TEXT">
                            <xsl:variable name="TARIH_TEXT" select="."/>
                            <xsl:variable name="ROW_BY_DATE" select="$ROW[TARIH_TEXT = $TARIH_TEXT]"/>    
                            
                            <!--tr><td><xsl:value-of select="TARIH_TEXT"/></td></tr-->
                                                     
                            <xsl:call-template name="toplam">
                                <xsl:with-param name="ROW" select="$ROW_BY_DATE"/>
                            </xsl:call-template>
                        </xsl:for-each>
                        
                    </tbody>
	                <tbody class="toplam">
                        
                        <!--xsl:apply-templates select="$ROW"/-->
                        <xsl:call-template name="toplam">
                            <xsl:with-param name="ROW" select="$ROW"/>
                        </xsl:call-template>
                    </tbody>
            </xsl:for-each>
            
    		
            <tfoot class="toplam">
    			<xsl:call-template name="toplam">
                	<xsl:with-param name="ROW" select="BIRIM/ROWSET/ROW"/>
                </xsl:call-template>
            </tfoot>
    	</table>
        </div>
    </xsl:template>
    
    <xsl:template match="ROWSET">
    	<div>
    	<table class="magaza">
        	<thead>
    		<tr class="head">
    			<th colspan="8"><xsl:value-of select="../@BIRIM_NO"/> - <xsl:value-of select="../@BIRIM_ADI"/></th>
    		</tr>
    		<tr class="head2">
    			<th>NET SATIŞ GEÇEN YIL</th>
    			<th>NET SATIŞ BÜTÇE</th>
    			<th>NET SATIS FİİLİ</th>
    			<th>Prog.Göre Artış</th>
                <th>G.Yıla Göre Artış</th>
                <th>MUSTERI SAYISI</th>
                <th>MUSTERI SAYISI GEÇEN YIL</th>
                <th>G.Yıla Göre Artış</th>
    		</tr>
            </thead>
            
            <xsl:variable name="ROWS" select="ROW"/>
            <xsl:for-each select="$Aylar">
            	
            	<xsl:variable name="TARIH_AY" select="@name"/>
                <xsl:variable name="ROW" select="$ROWS[TARIH_AY = $TARIH_AY]"/>
                
                <tbody>
                    <xsl:apply-templates select="$ROW"/>
                </tbody>
                <tbody class="toplam">
                    <xsl:call-template name="toplam">
                        <xsl:with-param name="ROW" select="$ROW"/>
                    </xsl:call-template>
                </tbody>
    		</xsl:for-each>
            
    		
            <tfoot class="toplam">
    			<xsl:call-template name="toplam">
                	<xsl:with-param name="ROW" select="ROW"/>
                </xsl:call-template>
            </tfoot>
    	</table>
        </div>
    </xsl:template>
     
    <xsl:template name="toplam" match="ROW">
   		<xsl:param name="ROW" select="."/>
        <xsl:variable name="NET_SATIS" select="sum($ROW/NET_SATIS)"/>
        <xsl:variable name="BUTCE" select="sum($ROW/BUTCE)"/>
        <xsl:variable name="NET_SATIS_FIILI" select="sum($ROW/NET_SATIS_FIILI)"/>
        
        <xsl:variable name="PROG_GORE_ARTIS" select="($NET_SATIS_FIILI * 100 div $BUTCE) - 100"/>
        <xsl:variable name="GY_GORE_ARTIS" select="($NET_SATIS_FIILI * 100 div $NET_SATIS) - 100"/>
        
        <xsl:variable name="MUSTERI_SAYISI" select="sum($ROW/MUSTERI_SAYISI)"/>
        <xsl:variable name="GY_MUSTERI_SAYISI" select="sum($ROW/GY_MUSTERI_SAYISI)"/>
        <xsl:variable name="MUSTERI_SAYISI_ARTIS" select="($MUSTERI_SAYISI * 100 div $GY_MUSTERI_SAYISI) - 100"/>
        
    	<tr>
            <td><xsl:value-of select="format-number($NET_SATIS,'#&#160;###', 'euro')"/></td>
            <td><xsl:value-of select="format-number($BUTCE,'#&#160;###', 'euro')"/></td>
            <td><xsl:value-of select="format-number($NET_SATIS_FIILI,'#&#160;###', 'euro')"/></td>
            <td><xsl:call-template name="neg"><xsl:with-param name="value" select="$PROG_GORE_ARTIS"/></xsl:call-template>
            	<xsl:value-of select="format-number($PROG_GORE_ARTIS,'####.##')"/>%</td>
            <td><xsl:call-template name="neg"><xsl:with-param name="value" select="$GY_GORE_ARTIS"/></xsl:call-template>
            	<xsl:value-of select="format-number($GY_GORE_ARTIS,'##0.00')"/>%</td>
            <xsl:if test="$ROW/MUSTERI_SAYISI">
                <td><xsl:value-of select="format-number($MUSTERI_SAYISI,'#&#160;###', 'euro')"/></td>
                <td><xsl:value-of select="format-number($GY_MUSTERI_SAYISI,'#&#160;###', 'euro')"/></td>
                <td><xsl:call-template name="neg"><xsl:with-param name="value" select="$MUSTERI_SAYISI_ARTIS"/></xsl:call-template>
                    <xsl:value-of select="format-number($MUSTERI_SAYISI_ARTIS,'##0.00')"/>%</td>
            </xsl:if>
        </tr>
     </xsl:template>
    
    
    <xsl:template name="magaza-toplam">
   		 <table class="stats">
         	<thead>
    		<tr>
            	<th>MAGAZA TOPLAM</th>
    			<th>NET SATIŞ GEÇEN YIL</th>
    			<th>NET SATIŞ BÜTÇE</th>
    			<th>NET SATIS FİİLİ</th>
    			<th>Prog.Göre Artış</th>
                <th>G.Yıla Göre Artış</th>
                <th>MUSTERI SAYISI</th>
                <th>MUSTERI SAYISI GEÇEN YIL</th>
                <th>G.Yıla Göre Artış</th>
    		</tr>
            </thead>
            <xsl:variable name="ROWS" select="BIRIM/ROWSET/ROW"/>
    		<xsl:for-each select="$Aylar">
                <xsl:variable name="TARIH_AY" select="@name"/>
                <xsl:variable name="ROW" select="$ROWS[TARIH_AY = $TARIH_AY]"/>
                <tbody>
                    <tr><th><xsl:value-of select="@title"/></th></tr>
                    <xsl:call-template name="toplam">
                        <xsl:with-param name="ROW" select="$ROW"/>
                    </xsl:call-template>
                </tbody>
            </xsl:for-each>
    	</table>
    </xsl:template>
    
    <xsl:template name="toptan-satis">
        <table class="stats">
        	<thead>
                <tr>
                    <th>TOPTAN SATIS</th>
                    <th>NET SATIŞ GEÇEN YIL</th>
                    <th>NET SATIŞ BÜTÇE</th>
                    <th>NET SATIS FİİLİ</th>
                    <th>Prog.Göre Artış</th>
                    <th>G.Yıla Göre Artış</th>
                </tr>
            </thead>
            <xsl:variable name="ROWS" select="TOPTAN_SATIS/ROWSET/ROW"/>
    		<xsl:for-each select="$Aylar">
                <xsl:variable name="TARIH_AY" select="@name"/>
                <xsl:variable name="ROW" select="$ROWS[TARIH_AY = $TARIH_AY]"/>
                <tbody>
                    <tr><th><xsl:value-of select="@title"/></th></tr>
                    <xsl:call-template name="toplam">
                        <xsl:with-param name="ROW" select="$ROW"/>
                    </xsl:call-template>
                </tbody>
            </xsl:for-each>
    	</table>
    </xsl:template>
    
    <xsl:template name="sirket-toplam">
        <table class="stats">
	        <thead>
    		<tr class="head">
            	<th>SIRKET TOPLAM</th>
    			<th>NET SATIŞ GEÇEN YIL</th>
    			<th>NET SATIŞ BÜTÇE</th>
    			<th>NET SATIS FİİLİ</th>
    			<th>Prog.Göre Artış</th>
                <th>G.Yıla Göre Artış</th>
                <th>MUSTERI SAYISI GEÇEN YIL</th>
                <th>MUSTERI SAYISI</th>
                <th>G.Yıla Göre Artış</th>
    		</tr>
            </thead>
            <xsl:variable name="ROWS" select="//ROWSET/ROW"/>
    		<xsl:for-each select="$Aylar">
                <xsl:variable name="TARIH_AY" select="@name"/>
                <xsl:variable name="ROW" select="$ROWS[TARIH_AY = $TARIH_AY]"/>
                <tbody>
                    <tr><th><xsl:value-of select="@title"/></th></tr>
                    <xsl:call-template name="toplam">
                        <xsl:with-param name="ROW" select="$ROW"/>
                    </xsl:call-template>
                </tbody>
            </xsl:for-each>
    	</table>    
    </xsl:template>
    
    <xsl:template name="sirket-toplam-kumule">
    	<table class="stats">
        	<thead>
    		<tr class="head">
            	<th>SIRKET TOPLAM</th>
    			<th>NET SATIŞ GEÇEN YIL</th>
    			<th>NET SATIŞ BÜTÇE</th>
    			<th>NET SATIS FİİLİ</th>
    			<th>Prog.Göre Artış</th>
                <th>G.Yıla Göre Artış</th>
                <th>MUSTERI SAYISI</th>
                <th>MUSTERI SAYISI GEÇEN YIL</th>
                <th>G.Yıla Göre Artış</th>
    		</tr>
            </thead>
            <tbody>
            <xsl:variable name="ROWS" select="//ROWSET/ROW"/>
            	<tr><th>KUMULE</th></tr>
    			<xsl:call-template name="toplam">
                    <xsl:with-param name="ROW" select="$ROWS"/>
                </xsl:call-template>
            </tbody>
    	</table>  
    </xsl:template>
    
    <xsl:template name="footer">
    	<form class="footer" method="GET">
	    	
	    	<input type="radio" name="doviz_kod" id="KZT" value="02" checked="checked" />
	    	<label for="KZT">KZT</label>
	    	
	    	<input type="radio" name="doviz_kod" id="USD" value="01" />
	    	<label for="USD">USD</label>
	    	
	    	<br/>
	    	
	    	<xsl:variable name="ROWS" select="BIRIM"/>
	    	
	    	<xsl:for-each select="BIRIMLER/ROWSET/ROW">
	    		<input type="checkbox" name="birimler" id="{BIRIM_ADI}" value="{BIRIM_NO}">
	    			<xsl:if test="BIRIM_NO[. = $ROWS/@BIRIM_NO]"><xsl:attribute name="checked">checked</xsl:attribute></xsl:if>
	    		</input>
	    		<label for="{BIRIM_ADI}"><xsl:value-of select="BIRIM_ADI"/></label>
	        </xsl:for-each>
	        
	        <br/>
	        
	        <input type="month" name="donem"/>
	        
	        <br/>
	        <input type="submit" value="Yenile"/>
        </form>
    </xsl:template>
    
    <xsl:template name="neg">
    	<xsl:param name="value"/>
        <xsl:if test="$value &lt; 0"><xsl:attribute name="class">neg</xsl:attribute></xsl:if>
    </xsl:template>
    
    
    <msxsl:script language="JScript" implements-prefix="exslt">
     this['node-set'] =  function (x) {
      return x;
      }
    </msxsl:script>
	
</xsl:stylesheet>