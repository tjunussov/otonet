<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:exslt="http://exslt.org/common" exclude-result-prefixes="exslt msxsl">
	
	<xsl:output method="html" />

	<xsl:strip-space elements="*" />
    
    <xsl:decimal-format name="euro" decimal-separator="," grouping-separator="&#160;"/>
    
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
            <xsl:variable name="FIRST_BIRIM_ROWS" select="/REPORT/BIRIM[1]/ROWSET/ROW"/>
            <xsl:variable name="ROWS" select="/REPORT/BIRIM/ROWSET/ROW"/>
            
            <xsl:for-each select="/REPORT/BIRIM[1]/ROWSET">
                <xsl:for-each select="ROW[not(TARIH_AY=preceding-sibling::ROW/TARIH_AY)]">
                    <xsl:variable name="TARIH_AY" select="TARIH_AY"/>
                    <xsl:variable name="FIRST_BIRIM_ROW" select="$FIRST_BIRIM_ROWS[TARIH_AY = $TARIH_AY]"/>
                    
                    <xsl:element name="MONTH">
                    	<xsl:attribute name="name"><xsl:value-of select="$TARIH_AY"/></xsl:attribute>
                    	<xsl:attribute name="title"><xsl:value-of select="$AylarDefinition[@name=$TARIH_AY]"/></xsl:attribute>
                        
                        <xsl:for-each select="$FIRST_BIRIM_ROW[not(TARIH_TEXT=preceding::TARIH_TEXT)]">
                            <xsl:element name="DATES">
	                            <xsl:attribute name="tarih"><xsl:value-of select="TARIH_TEXT"/></xsl:attribute>
                            	
                                    <xsl:variable name="TARIH_TEXT" select="TARIH_TEXT"/>
                                    <xsl:variable name="ROW" select="$ROWS[TARIH_TEXT = $TARIH_TEXT]"/>
    
                                        <xsl:for-each select="$ROW">
                                            <xsl:element name="STORE">
                                            	<xsl:attribute name="BIRIM_NO"><xsl:value-of select="../../@BIRIM_NO"/></xsl:attribute>
                                                <xsl:attribute name="BIRIM_ADI"><xsl:value-of select="../../@BIRIM_ADI"/></xsl:attribute>
                                                <xsl:copy-of select="*"/>
                                            </xsl:element>
                                        </xsl:for-each>
                        	</xsl:element><!--DATES-->
                        </xsl:for-each>
                        
                    </xsl:element><!--MONTH-->
                    
                </xsl:for-each>
            </xsl:for-each>
        </root>
    </xsl:variable>
    
    <xsl:variable name="Aylar"  select="exslt:node-set($AylarXML)/root[1]/*"/>
    <xsl:variable name="AylarHeader"  select="$Aylar[1]/DATES[1]/STORE"/>
    
    
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
        
        
        <div class="magazas">
        	<xsl:call-template name="magazas"/>
        </div>
        
        <xsl:call-template name="magaza-toplam"/>
        <xsl:call-template name="toptan-satis"/>
        <xsl:call-template name="sirket-toplam"/>
        <xsl:call-template name="sirket-toplam-kumule"/>
        
        <br/><br/><br/><br/>
        
        
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
    
    
    <xsl:template name="magazas">
    	
    	<table class="magaza">
        	<thead>
    		<tr class="head">
    			<th class="tarih">&#160;</th>
                <th colspan="8">MAGAZA TOPLAM</th>
            <xsl:for-each select="$AylarHeader">
                <th colspan="8"><xsl:value-of select="@BIRIM_NO"/> - <xsl:value-of select="@BIRIM_ADI"/></th>
            </xsl:for-each>
    		</tr>
    		<tr class="head2">
    		   <th class="tarih">TARIH</th>
            	<xsl:call-template name="birim_titles"/>
            <xsl:for-each select="$AylarHeader">
               <xsl:call-template name="birim_titles"/>
            </xsl:for-each>
    		</tr>
            </thead>
            <xsl:for-each select="$Aylar">
            <tbody>
	            <xsl:for-each select="DATES">
                <tr>
                    <td class="tarih"><xsl:value-of select="@tarih"/></td>
                    <xsl:call-template name="birim_values">
                    	<xsl:with-param name="ROW" select="STORE"/>
                    </xsl:call-template>
                <xsl:for-each select="STORE">
                    <xsl:call-template name="birim_values"/>
                </xsl:for-each>
                </tr>
            	</xsl:for-each>
            </tbody>
            <tbody class="toplam">
                <tr>
                    <th class="tarih"><xsl:value-of select="@title"/></th>
                    <xsl:call-template name="birim_values">
                    	<xsl:with-param name="ROW" select="DATES/STORE"/>
                    </xsl:call-template>
                
                <xsl:variable name="DATES" select="DATES"/>
                <xsl:for-each select="DATES[1]/STORE">
                	<xsl:variable name="BIRIM_NO" select="@BIRIM_NO"/>
                    <xsl:call-template name="birim_values">
                    	<xsl:with-param name="ROW" select="$DATES/STORE[@BIRIM_NO = $BIRIM_NO]"/>
                    </xsl:call-template>
	            </xsl:for-each>
                </tr>
            </tbody>
    		</xsl:for-each>
            <tfoot class="toplam">
    		<tr>
    			<th class="tarih">TOPLAM</th>
                <xsl:call-template name="birim_values">
                    <xsl:with-param name="ROW" select="$Aylar/DATES/STORE"/>
                </xsl:call-template>
                
                <xsl:for-each select="$AylarHeader">
                	<xsl:variable name="BIRIM_NO" select="@BIRIM_NO"/>
                    <xsl:call-template name="birim_values">
                    	<xsl:with-param name="ROW" select="$Aylar/DATES/STORE[@BIRIM_NO = $BIRIM_NO]"/>
                    </xsl:call-template>
	            </xsl:for-each>
                
    		</tr>
            </tfoot>
    	</table>
        
    </xsl:template>
    
    <xsl:template name="birim_values">
   		<xsl:param name="ROW" select="."/>
        <xsl:variable name="NET_SATIS" select="sum($ROW/NET_SATIS)"/>
        <xsl:variable name="BUTCE" select="sum($ROW/BUTCE)"/>
        <xsl:variable name="NET_SATIS_FIILI" select="sum($ROW/NET_SATIS_FIILI)"/>
        
        <xsl:variable name="PROG_GORE_ARTIS" select="($NET_SATIS_FIILI * 100 div $BUTCE) - 100"/>
        <xsl:variable name="GY_GORE_ARTIS" select="($NET_SATIS_FIILI * 100 div $NET_SATIS) - 100"/>
        
        <xsl:variable name="MUSTERI_SAYISI" select="sum($ROW/MUSTERI_SAYISI)"/>
        <xsl:variable name="GY_MUSTERI_SAYISI" select="sum($ROW/GY_MUSTERI_SAYISI)"/>
        <xsl:variable name="MUSTERI_SAYISI_ARTIS" select="($MUSTERI_SAYISI * 100 div $GY_MUSTERI_SAYISI) - 100"/>
        
        <td class="col-start"><xsl:value-of select="format-number($NET_SATIS,'#&#160;###', 'euro')"/></td>
        <td><xsl:value-of select="format-number($BUTCE,'#&#160;###', 'euro')"/></td>
        <td><xsl:value-of select="format-number($NET_SATIS_FIILI,'#&#160;###', 'euro')"/></td>
        <td><xsl:call-template name="neg"><xsl:with-param name="value" select="$PROG_GORE_ARTIS"/></xsl:call-template>
            <xsl:value-of select="format-number($PROG_GORE_ARTIS,'####.##')"/>%</td>
        <td><xsl:call-template name="neg"><xsl:with-param name="value" select="$GY_GORE_ARTIS"/></xsl:call-template>
            <xsl:value-of select="format-number($GY_GORE_ARTIS,'##0.00')"/>%</td>
        <xsl:if test="$ROW/MUSTERI_SAYISI">
            <td><xsl:value-of select="format-number($MUSTERI_SAYISI,'#&#160;###', 'euro')"/></td>
            <td><xsl:value-of select="format-number($GY_MUSTERI_SAYISI,'#&#160;###', 'euro')"/></td>
            <td class="col-end"><xsl:call-template name="neg"><xsl:with-param name="value" select="$MUSTERI_SAYISI_ARTIS"/></xsl:call-template>
                <xsl:value-of select="format-number($MUSTERI_SAYISI_ARTIS,'##0.00')"/>%</td>
        </xsl:if>
        
     </xsl:template>
     
     
     <xsl:template name="birim_titles">
        <th class="col-start">NET SATIŞ GEÇEN YIL</th>
        <th>NET SATIŞ BÜTÇE</th>
        <th>NET SATIS FİİLİ</th>
        <th>Prog.Göre Artış</th>
        <th>G.Yıla Göre Artış</th>
        <th>MUSTERI SAYISI</th>
        <th>MUSTERI SAYISI GEÇEN YIL</th>
        <th class="col-end">G.Yıla Göre Artış</th>
     </xsl:template>
     
     
    
    <xsl:template name="magaza-toplam">
   		 
         <xsl:variable name="MtXML">
    		<tr>
            	<th>MAGAZA TOPLAM</th>
    			<xsl:call-template name="birim_titles"/>
            </tr>
            <xsl:variable name="ROWS" select="BIRIM/ROWSET/ROW"/>
    		<xsl:for-each select="$Aylar">
                <xsl:variable name="TARIH_AY" select="@name"/>
                <xsl:variable name="ROW" select="$ROWS[TARIH_AY = $TARIH_AY]"/>
                <tr>
                    <th><xsl:value-of select="@title"/></th>
                    <xsl:call-template name="birim_values">
                        <xsl:with-param name="ROW" select="$ROW"/>
                    </xsl:call-template>
                </tr>
            </xsl:for-each>
    	</xsl:variable>
        
        <table class="stats">
        	<xsl:call-template name="transpose">
            	<xsl:with-param name="value" select="exslt:node-set($MtXML)/*"/>
            </xsl:call-template>
        </table>
         
    </xsl:template>
    
    
    
    <xsl:template name="toptan-satis">

        <xsl:variable name="MtXML">
    		<tr>
            	<th>TOPTAN SATIS</th>
    			<th>NET SATIŞ GEÇEN YIL</th>
                <th>NET SATIŞ BÜTÇE</th>
                <th>NET SATIS FİİLİ</th>
                <th>Prog.Göre Artış</th>
                <th>G.Yıla Göre Artış</th>
            </tr>
            <xsl:variable name="ROWS" select="TOPTAN_SATIS/ROWSET/ROW"/>
    		<xsl:for-each select="$Aylar">
                <xsl:variable name="TARIH_AY" select="@name"/>
                <xsl:variable name="ROW" select="$ROWS[TARIH_AY = $TARIH_AY]"/>
                    <tr>
                    	<th><xsl:value-of select="@title"/></th>
                        <xsl:call-template name="birim_values">
                            <xsl:with-param name="ROW" select="$ROW"/>
                        </xsl:call-template>
                    </tr>
            </xsl:for-each>
    	</xsl:variable>
        
        <table class="stats">
        	<xsl:call-template name="transpose">
            	<xsl:with-param name="value" select="exslt:node-set($MtXML)/*"/>
            </xsl:call-template>
        </table>
        
    </xsl:template>
    
    <xsl:template name="sirket-toplam">
    
    	<xsl:variable name="MtXML">
    		<tr>
            	<th>SIRKET TOPLAM</th>
    			<xsl:call-template name="birim_titles"/>
            </tr>
            <xsl:variable name="ROWS" select="//ROWSET/ROW"/>
    		<xsl:for-each select="$Aylar">
                <xsl:variable name="TARIH_AY" select="@name"/>
                <xsl:variable name="ROW" select="$ROWS[TARIH_AY = $TARIH_AY]"/>
                <tr>
                	<th><xsl:value-of select="@title"/></th>
                    <xsl:call-template name="birim_values">
                        <xsl:with-param name="ROW" select="$ROW"/>
                    </xsl:call-template>
                </tr>
            </xsl:for-each>
    	</xsl:variable>
        
        <table class="stats">
        	<xsl:call-template name="transpose">
            	<xsl:with-param name="value" select="exslt:node-set($MtXML)/*"/>
            </xsl:call-template>
        </table>
        
    </xsl:template>
    
    <xsl:template name="sirket-toplam-kumule">
    
    	<xsl:variable name="MtXML">
    		<tr>
            	<th>SIRKET TOPLAM</th>
    			<xsl:call-template name="birim_titles"/>
            </tr>
            <xsl:variable name="ROWS" select="//ROWSET/ROW"/>
            <tr>
                <th>KUMULE</th>
                <xsl:call-template name="birim_values">
                    <xsl:with-param name="ROW" select="$ROWS"/>
                </xsl:call-template>
            </tr>
    	</xsl:variable>
        
        <table class="stats">
        	<xsl:call-template name="transpose">
            	<xsl:with-param name="value" select="exslt:node-set($MtXML)/*"/>
            </xsl:call-template>
        </table>
        
    </xsl:template>
    
    <xsl:template name="footer">
    	<form class="footer" method="GET">
	    	
	    	<input type="radio" name="doviz_kod" id="KZT" value="02" checked="checked" />
	    	<label for="KZT">KZT</label>
	    	
	    	<input type="radio" name="doviz_kod" id="USD" value="01" />
	    	<label for="USD">USD</label>
	    	
	    	&#160;
	    	
	    	<xsl:variable name="ROWS" select="BIRIM"/>
	    	
	    	<xsl:for-each select="BIRIMLER/ROWSET/ROW">
	    		<input type="checkbox" name="birimler" id="{BIRIM_ADI}" value="{BIRIM_NO}">
	    			<xsl:if test="BIRIM_NO[. = $ROWS/@BIRIM_NO]"><xsl:attribute name="checked">checked</xsl:attribute></xsl:if>
	    		</input>
	    		<label for="{BIRIM_ADI}"><xsl:value-of select="BIRIM_ADI"/></label>
	        </xsl:for-each>
	        &#160;
	        <input type="submit" value="Обновить"/>
        </form>
    </xsl:template>
    
    <xsl:template name="transpose">
   		 <xsl:param name="value"/>
         
         <xsl:for-each select="$value[1]/child::node()">
         <xsl:variable name="pos" select="position()"/>
         <tr>
			<xsl:for-each select="$value">
            	<xsl:copy-of select="*[position()=$pos]"/>
			</xsl:for-each>
         </tr>
         </xsl:for-each>
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