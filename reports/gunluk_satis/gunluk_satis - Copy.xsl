<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:date="http://exslt.org/dates-and-times" xmlns:exslt="http://exslt.org/common" xmlns:c="c" exclude-result-prefixes="exslt c date">
	
	<xsl:output method="html" omit-xml-declaration="yes"/>
	<xsl:strip-space elements="*" />
    
    <xsl:decimal-format name="euro" decimal-separator="," grouping-separator="&#160;"/>
    
    <xsl:param name="CalendarXML">
    </xsl:param>
            
    <xsl:key name="dates" match="/ALL/ROWSET[1]/ROW" use="TARIH_AY" />
    
    <xsl:template match="/ALL"><link href="excel.css" type="text/css" rel="stylesheet"/>
	    <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.12.0/jquery.min.js"></script>
    
        <h2>RAMSTORE KAZAKISTAN - Gunluk Satis</h2>
        <div class="magazas">
	    	<xsl:apply-templates select="ROWSET[1]" mode="dates"/>
	    	<xsl:apply-templates/>
        </div>
        
        <xsl:call-template name="magaza-toplam"/>
        <xsl:call-template name="toptan-satis"/>
        <xsl:call-template name="sirket-toplam"/>
        <xsl:call-template name="sirket-toplam-kumule"/>
        
        <script>
		
		
		$(function() {
			var top = $('.tarih').position().top;
			$(window).scroll(function(){
			  $('.tarih').css('top',(top-$(window).scrollTop())+'px');
			});
		});
		</script>
        
    </xsl:template>
    
    <xsl:template match="ROWSET" mode="dates">
    	<table class="tarih">
        	<thead>
    		<tr class="head">
    			<th>&#160;</th>
    		</tr>
    		<tr class="head2">
    			<th>TARIH</th>
    		</tr>
            </thead>
            <xsl:variable name="ROWS" select="ROW"/>
            <xsl:for-each select="ROW[not(TARIH_AY=preceding-sibling::ROW/TARIH_AY)]">
            <xsl:variable name="TARIH_AY" select="TARIH_AY"/>            
            <xsl:variable name="ROW" select="$ROWS[TARIH_AY = $TARIH_AY]"/>
            <tbody>
	            <xsl:for-each select="$ROW">
                <tr>
                    <td><xsl:value-of select="TARIH_TEXT"/></td>
                </tr>
            	</xsl:for-each>
                <tr class="toplam">
                    <th><xsl:value-of select="TARIH_AY"/></th>
                </tr>
            </tbody>
    		</xsl:for-each>
            <tfoot>
    		<tr class="toplam">
    			<th>TOPLAM</th>
    		</tr>
            </tfoot>
    	</table>
    </xsl:template>
    
    <xsl:template match="ROWSET">
    	<table class="magaza">
        	<thead>
    		<tr class="head">
    			<th colspan="8"><xsl:value-of select="ROW/BIRIM_NO"/> - <xsl:value-of select="ROW/BIRIM_ADI"/></th>
    		</tr>
    		<tr class="head2">
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
            
            <xsl:variable name="ROWS" select="ROW"/>
            <xsl:for-each select="ROW[not(TARIH_AY=preceding-sibling::ROW/TARIH_AY)]">
            	
            	<xsl:variable name="TARIH_AY" select="TARIH_AY"/>            
                <xsl:variable name="ROW" select="$ROWS[TARIH_AY = $TARIH_AY]"/>
                
                <tbody>
                    <xsl:apply-templates select="$ROW"/>
                    <tr class="toplam">
                        <th><xsl:value-of select="format-number(sum($ROW/NET_SATIS),'#&#160;###', 'euro')"/></th>
                        <th><xsl:value-of select="format-number(sum($ROW/BUTCE),'#&#160;###', 'euro')"/></th>
                        <th><xsl:value-of select="format-number(sum($ROW/NET_SATIS_FIILI),'#&#160;###', 'euro')"/></th>
                        <th><xsl:value-of select="format-number(sum($ROW/NET_SATIS_FIILI),'#&#160;###', 'euro')"/></th>
                        <th><xsl:value-of select="format-number(sum($ROW/NET_SATIS_FIILI),'#&#160;###', 'euro')"/></th>
                        <th><xsl:value-of select="format-number(sum($ROW/MUSTERI_SAYISI),'#&#160;###', 'euro')"/></th>
                        <th><xsl:value-of select="format-number(sum($ROW/MUSTERI_SAYISI),'#&#160;###', 'euro')"/></th>
                        <th><xsl:value-of select="format-number(sum($ROW/MUSTERI_SAYISI),'#&#160;###', 'euro')"/></th>
                    </tr>
                </tbody>
    		</xsl:for-each>
            
    		
            <tfoot>
    		<tr class="toplam">
                <th><xsl:value-of select="format-number(sum(ROW/NET_SATIS),'#&#160;###', 'euro')"/></th>
                <th><xsl:value-of select="format-number(sum(ROW/BUTCE),'#&#160;###', 'euro')"/></th>
    			<th><xsl:value-of select="format-number(sum(ROW/NET_SATIS_FIILI),'#&#160;###', 'euro')"/></th>
                <th><xsl:value-of select="format-number(sum(ROW/NET_SATIS_FIILI),'#&#160;###', 'euro')"/></th>
                <th><xsl:value-of select="format-number(sum(ROW/NET_SATIS_FIILI),'#&#160;###', 'euro')"/></th>
    			<th><xsl:value-of select="format-number(sum(ROW/MUSTERI_SAYISI),'#&#160;###', 'euro')"/></th>
                <th><xsl:value-of select="format-number(sum(ROW/MUSTERI_SAYISI),'#&#160;###', 'euro')"/></th>
                <th><xsl:value-of select="format-number(sum(ROW/MUSTERI_SAYISI),'#&#160;###', 'euro')"/></th>
    		</tr>
            </tfoot>
    	</table>
    </xsl:template>
    
    <xsl:template match="ROW">
    	<tr>
    		<td><xsl:value-of select="format-number(NET_SATIS,'#&#160;###', 'euro')"/></td>
            <td><xsl:value-of select="format-number(BUTCE,'#&#160;###', 'euro')"/></td>
    		<td><xsl:value-of select="format-number(NET_SATIS_FIILI,'#&#160;###', 'euro')"/></td>
    		<td><xsl:value-of select="format-number((NET_SATIS_FIILI * 100 div BUTCE) - 100,'####.##')"/>%</td>
            <td><xsl:value-of select="format-number((NET_SATIS_FIILI * 100 div NET_SATIS) - 100,'##0.00')"/>%</td>
    		<td><xsl:value-of select="format-number(MUSTERI_SAYISI,'#&#160;###', 'euro')"/></td>
            <td><xsl:value-of select="format-number(MUSTERI_SAYISI,'#&#160;###', 'euro')"/></td>
            <td><xsl:value-of select="format-number((MUSTERI_SAYISI * 100 div MUSTERI_SAYISI) - 100,'##0.00')"/>%</td>
    	</tr>
    </xsl:template>
    
    
    <xsl:template name="magaza-toplam">
   		 <xtable class="stats">
    		<tc class="head">
            	<ta>MAGAZA TOPLAM</ta>
    			<ta>NET SATIŞ GEÇEN YIL</ta>
    			<ta>NET SATIŞ BÜTÇE</ta>
    			<ta>NET SATIS FİİLİ</ta>
    			<ta>Prog.Göre Artış</ta>
                <ta>G.Yıla Göre Artış</ta>
                <ta>MUSTERI SAYISI GEÇEN YIL</ta>
                <ta>MUSTERI SAYISI</ta>
                <ta>G.Yıla Göre Artış</ta>
    		</tc>
    		<xsl:for-each select="ROWSET">
            <tc>
            	<ta><xsl:value-of select="ROW/TARIH_AY"/></ta>
                <tn><xsl:value-of select="format-number(sum(ROW/NET_SATIS),'#&#160;###', 'euro')"/></tn>
                <tn><xsl:value-of select="format-number(sum(ROW/BUTCE),'#&#160;###', 'euro')"/></tn>
    			<tn><xsl:value-of select="format-number(sum(ROW/NET_SATIS_FIILI),'#&#160;###', 'euro')"/></tn>
                <tn><xsl:value-of select="format-number(sum(ROW/NET_SATIS_FIILI),'#&#160;###', 'euro')"/></tn>
                <tn><xsl:value-of select="format-number(sum(ROW/NET_SATIS_FIILI),'#&#160;###', 'euro')"/></tn>
    			<tn><xsl:value-of select="format-number(sum(ROW/MUSTERI_SAYISI),'#&#160;###', 'euro')"/></tn>
                <tn><xsl:value-of select="format-number(sum(ROW/MUSTERI_SAYISI),'#&#160;###', 'euro')"/></tn>
                <tn><xsl:value-of select="format-number(sum(ROW/MUSTERI_SAYISI),'#&#160;###', 'euro')"/></tn>
    		</tc>
    		</xsl:for-each>
    	</xtable>
    </xsl:template>
    
    <xsl:template name="toptan-satis">
        <xtable class="stats">
    		<tc class="head">
            	<ta>TOPTAN SATIS</ta>
    			<ta>NET SATIŞ GEÇEN YIL</ta>
    			<ta>NET SATIŞ BÜTÇE</ta>
    			<ta>NET SATIS FİİLİ</ta>
    			<ta>Prog.Göre Artış</ta>
                <ta>G.Yıla Göre Artış</ta>
                <ta>MUSTERI SAYISI GEÇEN YIL</ta>
                <ta>MUSTERI SAYISI</ta>
                <ta>G.Yıla Göre Artış</ta>
    		</tc>
    	</xtable>
    </xsl:template>
    
    <xsl:template name="sirket-toplam">
        <xtable class="stats">
    		<tc class="head">
            	<ta>SIRKET TOPLAM</ta>
    			<ta>NET SATIŞ GEÇEN YIL</ta>
    			<ta>NET SATIŞ BÜTÇE</ta>
    			<ta>NET SATIS FİİLİ</ta>
    			<ta>Prog.Göre Artış</ta>
                <ta>G.Yıla Göre Artış</ta>
                <ta>MUSTERI SAYISI GEÇEN YIL</ta>
                <ta>MUSTERI SAYISI</ta>
                <ta>G.Yıla Göre Artış</ta>
    		</tc>
    	</xtable>    
    </xsl:template>
    
    <xsl:template name="sirket-toplam-kumule">
    	<xtable class="stats">
    		<tc class="head">
            	<ta>SIRKET TOPLAM</ta>
    			<ta>NET SATIŞ GEÇEN YIL</ta>
    			<ta>NET SATIŞ BÜTÇE</ta>
    			<ta>NET SATIS FİİLİ</ta>
    			<ta>Prog.Göre Artış</ta>
                <ta>G.Yıla Göre Artış</ta>
                <ta>MUSTERI SAYISI GEÇEN YIL</ta>
                <ta>MUSTERI SAYISI</ta>
                <ta>G.Yıla Göre Artış</ta>
    		</tc>
            <tc>
            	<ta>KUMULE</ta>
    			<tn>#&#160;</tn>
    			<tn>#&#160;</tn>
    			<tn>#&#160;</tn>
    			<tn>#&#160;</tn>
                <tn>#&#160;</tn>
                <tn>#&#160;</tn>
                <tn>#&#160;</tn>
                <tn>#&#160;</tn>
    		</tc>
    	</xtable>  
    </xsl:template>
	
</xsl:stylesheet>