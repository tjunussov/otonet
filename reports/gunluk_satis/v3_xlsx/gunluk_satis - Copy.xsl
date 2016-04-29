<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" 
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
    xmlns:msxsl="urn:schemas-microsoft-com:xslt" 
    xmlns:exslt="http://exslt.org/common" 
    xmlns="urn:schemas-microsoft-com:office:spreadsheet"
    xmlns:r="ramstore"
	xmlns:o="urn:schemas-microsoft-com:office:office" 
	xmlns:x="urn:schemas-microsoft-com:office:excel"
	xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
    exclude-result-prefixes="exslt msxsl r">
	
    <xsl:decimal-format name="euro" decimal-separator="," grouping-separator="&#160;"/>
    
    <xsl:variable name="AylarDefinitionXML">
    	<r:root>
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
        </r:root>
    </xsl:variable>
    
    <xsl:variable name="AylarDefinition"  select="exslt:node-set($AylarDefinitionXML)/r:root[1]/*"/>
    
    <xsl:variable name="AylarXML">
    	<r:root>
            <xsl:variable name="FIRST_BIRIM_ROWS" select="/REPORT/BIRIM[1]/ROWSET/ROW"/>
            <xsl:variable name="ROWS" select="/REPORT/BIRIM/ROWSET/ROW"/>
            
            <xsl:for-each select="/REPORT/BIRIM[1]/ROWSET">
                <xsl:for-each select="ROW[not(TARIH_AY=preceding-sibling::ROW/TARIH_AY)]">
                    <xsl:variable name="TARIH_AY" select="TARIH_AY"/>
                    <xsl:variable name="FIRST_BIRIM_ROW" select="$FIRST_BIRIM_ROWS[TARIH_AY = $TARIH_AY]"/>
                    
                    <xsl:element name="MONTH" namespace="ramstore">
                    	<xsl:attribute name="name"><xsl:value-of select="$TARIH_AY"/></xsl:attribute>
                    	<xsl:attribute name="title"><xsl:value-of select="$AylarDefinition[@name=$TARIH_AY]"/></xsl:attribute>
                        
                        <xsl:for-each select="$FIRST_BIRIM_ROW[not(TARIH_TEXT=preceding::TARIH_TEXT)]">
                            <xsl:element name="DATES" namespace="ramstore">
	                            <xsl:attribute name="tarih"><xsl:value-of select="TARIH_TEXT"/></xsl:attribute>
                            	
                                    <xsl:variable name="TARIH_TEXT" select="TARIH_TEXT"/>
                                    <xsl:variable name="ROW" select="$ROWS[TARIH_TEXT = $TARIH_TEXT]"/>
    
                                        <xsl:for-each select="$ROW">
                                            <xsl:element name="STORE" namespace="ramstore">
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
        </r:root>
    </xsl:variable>
    
    <xsl:variable name="Aylar"  select="exslt:node-set($AylarXML)/r:root[1]/*"/>
    <xsl:variable name="AylarHeader"  select="$Aylar[1]/r:DATES[1]/r:STORE"/>
    
    
    <xsl:template match="/REPORT">
    
    	<xsl:processing-instruction name="mso-application">   
		<xsl:text>progid="Excel.Sheet"</xsl:text>  
		</xsl:processing-instruction>
        
        <Workbook>
        	<DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
              <Author>root</Author>
              <LastAuthor>Timur Junussov</LastAuthor>
              <Created>2016-02-17T10:35:33Z</Created>
              <LastSaved>2016-03-05T11:10:27Z</LastSaved>
              <Version>14.00</Version>
             </DocumentProperties>
             <OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">
              <AllowPNG/>
             </OfficeDocumentSettings>
             <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
              <WindowHeight>1185</WindowHeight>
              <WindowWidth>2100</WindowWidth>
              <WindowTopX>360</WindowTopX>
              <WindowTopY>0</WindowTopY>
              <ActiveSheet>1</ActiveSheet>
              <ProtectStructure>False</ProtectStructure>
              <ProtectWindows>False</ProtectWindows>
             </ExcelWorkbook>
			<Styles>
				<Style ss:ID="Default" ss:Name="Normal">
					<Alignment ss:Vertical="Bottom" />
					<Borders />
					<Font />
					<Interior />
					<NumberFormat />
					<Protection />
				</Style>
				<Style ss:ID="s21">
					<Font ss:Size="22" ss:Bold="1" />
				</Style>
				<Style ss:ID="s22">
					<Font ss:Size="14" ss:Bold="1" />
				</Style>
				<Style ss:ID="s23">
					<Font ss:Size="12" ss:Bold="1" />
				</Style>
				<Style ss:ID="s24">
					<Font ss:Size="10" ss:Bold="1" />
				</Style>
			</Styles>
            
            <Worksheet ss:Name="{//Report/@caption}">
                <xsl:attribute name="ss:Name">
                    <xsl:choose>
                        <xsl:when test="@doviz_kod = '02'">KZT</xsl:when>
                        <xsl:when test="@doviz_kod = '01'">USD</xsl:when>
                        <xsl:otherwise><xsl:value-of select="@doviz_kod"/></xsl:otherwise>
                    </xsl:choose>
                </xsl:attribute>
                
                <Table>
                
                	<Column ss:AutoFitWidth="0" ss:Width="85" />
					<Column ss:AutoFitWidth="0" ss:Width="115" />
					<Column ss:AutoFitWidth="0" ss:Width="115" />
					<Column ss:AutoFitWidth="0" ss:Width="160" />
					<Column ss:AutoFitWidth="0" ss:Width="115" />
					<Column ss:AutoFitWidth="0" ss:Width="85" />
					<Column ss:AutoFitWidth="0" ss:Width="85" />
					<Column ss:AutoFitWidth="0" ss:Width="160" />
                    
                    <Row ss:Height="22.5">
                    	<Cell ss:StyleID="s62"><Data ss:Type="String">RAMSTORE KAZAKISTAN - Gunluk Satis</Data></Cell>
                    </Row>
        
        <!--l>Magazalar</l><v><xsl:value-of select="@birim_no"/></v><br />
        <l>Doviz</l>
        <v>
        	<xsl:choose>
            	<xsl:when test="@doviz_kod = '02'">KZT</xsl:when>
                <xsl:when test="@doviz_kod = '01'">USD</xsl:when>
                <xsl:otherwise><xsl:value-of select="@doviz_kod"/></xsl:otherwise>
            </xsl:choose>
        </v><br />
        <l>Donem</l><v><xsl:value-of select="@donem_from"/> - <xsl:value-of select="@donem_to"/></v><br /><br /-->
        
                    
                    <xsl:call-template name="magazas"/>
                    
                    <!--xsl:call-template name="magaza-toplam"/>
                    <xsl:call-template name="toptan-satis"/>
                    <xsl:call-template name="sirket-toplam"/>
                    <xsl:call-template name="sirket-toplam-kumule"/>
                    
                    <xsl:call-template name="footer"/-->
                    
                  </Table>
        
    		
        </Worksheet>
      </Workbook>
    </xsl:template>
    
    
    <xsl:template name="magazas">
    	
    	<Row ss:Height="22.5">
        	<Cell ss:StyleID="s62"></Cell>
            <Cell ss:StyleID="s62"><Data ss:Type="String">MAGAZA TOPLAM</Data></Cell>
            <xsl:for-each select="$AylarHeader">
                <Cell ss:StyleID="s62"><Data ss:Type="String"><xsl:value-of select="@BIRIM_NO"/> - <xsl:value-of select="@BIRIM_ADI"/></Data></Cell>
            </xsl:for-each>
    	</Row>
    	<Row ss:Height="22.5">
    		   <Cell ss:StyleID="s62" class="tarih">TARIH</Cell>
            	<xsl:call-template name="birim_titles"/>
            <xsl:for-each select="$AylarHeader">
               <xsl:call-template name="birim_titles"/>
            </xsl:for-each>
    	</Row>
            
            <xsl:for-each select="$Aylar">
            
	            <xsl:for-each select="DATES">
                <Row ss:Height="22.5">
                    <Cell ss:StyleID="s62" class="tarih"><Data ss:Type="String"><xsl:value-of select="@tarih"/></Data></Cell>
                    <xsl:call-template name="birim_values">
                    	<xsl:with-param name="ROW" select="r:STORE"/>
                    </xsl:call-template>
                <xsl:for-each select="STORE">
                    <xsl:call-template name="birim_values"/>
                </xsl:for-each>
                </Row>
            	</xsl:for-each>
            
            <!--tbody class="toplam"-->
                <Row ss:Height="22.5">
                    <Cell ss:StyleID="s62" class="tarih"><xsl:value-of select="@title"/></Cell>
                    <xsl:call-template name="birim_values">
                    	<xsl:with-param name="ROW" select="r:DATES/r:STORE"/>
                    </xsl:call-template>
                
                <xsl:variable name="DATES" select="r:DATES"/>
                <xsl:for-each select="r:DATES[1]/r:STORE">
                	<xsl:variable name="BIRIM_NO" select="@BIRIM_NO"/>
                    <xsl:call-template name="birim_values">
                    	<xsl:with-param name="ROW" select="$DATES/r:STORE[@BIRIM_NO = $BIRIM_NO]"/>
                    </xsl:call-template>
	            </xsl:for-each>
                </Row>
            
    		</xsl:for-each>
            <!--tfoot class="toplam"-->
    		<Row ss:Height="22.5">
    			<Cell ss:StyleID="s62" class="tarih">TOPLAM</Cell>
                <xsl:call-template name="birim_values">
                    <xsl:with-param name="ROW" select="$Aylar/r:DATES/r:STORE"/>
                </xsl:call-template>
                
                <xsl:for-each select="$AylarHeader">
                	<xsl:variable name="BIRIM_NO" select="@BIRIM_NO"/>
                    <xsl:call-template name="birim_values">
                    	<xsl:with-param name="ROW" select="$Aylar/r:DATES/r:STORE[@BIRIM_NO = $BIRIM_NO]"/>
                    </xsl:call-template>
	            </xsl:for-each>
                
    		</Row>
            
    	
        
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
        
        <Cell ss:StyleID="s62" class="col-start"><xsl:value-of select="format-number($NET_SATIS,'#&#160;###', 'euro')"/></Cell>
        <Cell ss:StyleID="s62"><xsl:value-of select="format-number($BUTCE,'#&#160;###', 'euro')"/></Cell>
        <Cell ss:StyleID="s62"><xsl:value-of select="format-number($NET_SATIS_FIILI,'#&#160;###', 'euro')"/></Cell>
        <Cell ss:StyleID="s62"><xsl:call-template name="neg"><xsl:with-param name="value" select="$PROG_GORE_ARTIS"/></xsl:call-template>
            <xsl:value-of select="format-number($PROG_GORE_ARTIS,'####.##')"/>%</Cell>
        <Cell ss:StyleID="s62"><xsl:call-template name="neg"><xsl:with-param name="value" select="$GY_GORE_ARTIS"/></xsl:call-template>
            <xsl:value-of select="format-number($GY_GORE_ARTIS,'##0.00')"/>%</Cell>
        <xsl:if test="$ROW/MUSTERI_SAYISI">
            <Cell ss:StyleID="s62"><xsl:value-of select="format-number($MUSTERI_SAYISI,'#&#160;###', 'euro')"/></Cell>
            <Cell ss:StyleID="s62"><xsl:value-of select="format-number($GY_MUSTERI_SAYISI,'#&#160;###', 'euro')"/></Cell>
            <Cell ss:StyleID="s62" class="col-end"><xsl:call-template name="neg"><xsl:with-param name="value" select="$MUSTERI_SAYISI_ARTIS"/></xsl:call-template>
                <xsl:value-of select="format-number($MUSTERI_SAYISI_ARTIS,'##0.00')"/>%</Cell>
        </xsl:if>
        
     </xsl:template>
     
     
     <xsl:template name="birim_titles">
        <Cell ss:StyleID="s62" class="col-start">NET SATIŞ GEÇEN YIL</Cell>
        <Cell ss:StyleID="s62">NET SATIŞ BÜTÇE</Cell>
        <Cell ss:StyleID="s62">NET SATIS FİİLİ</Cell>
        <Cell ss:StyleID="s62">Prog.Göre Artış</Cell>
        <Cell ss:StyleID="s62">G.Yıla Göre Artış</Cell>
        <Cell ss:StyleID="s62">MUSTERI SAYISI</Cell>
        <Cell ss:StyleID="s62">MUSTERI SAYISI GEÇEN YIL</Cell>
        <Cell ss:StyleID="s62" class="col-end">G.Yıla Göre Artış</Cell>
     </xsl:template>
     
     
    
    <xsl:template name="magaza-toplam">
   		 
         <xsl:variable name="MtXML">
    		<Row ss:Height="22.5">
            	<Cell ss:StyleID="s62">MAGAZA TOPLAM</Cell>
    			<xsl:call-template name="birim_titles"/>
            </Row>
            <xsl:variable name="ROWS" select="BIRIM/ROWSET/ROW"/>
    		<xsl:for-each select="$Aylar">
                <xsl:variable name="TARIH_AY" select="@name"/>
                <xsl:variable name="ROW" select="$ROWS[TARIH_AY = $TARIH_AY]"/>
                <Row ss:Height="22.5">
                    <Cell ss:StyleID="s62"><xsl:value-of select="@title"/></Cell>
                    <xsl:call-template name="birim_values">
                        <xsl:with-param name="ROW" select="$ROW"/>
                    </xsl:call-template>
                </Row>
            </xsl:for-each>
    	</xsl:variable>
        
        <!--table class="stats"-->
        	<xsl:call-template name="transpose">
            	<xsl:with-param name="value" select="exslt:node-set($MtXML)/*"/>
            </xsl:call-template>
        
         
    </xsl:template>
    
    
    
    <xsl:template name="toptan-satis">

        <xsl:variable name="MtXML">
    		<Row ss:Height="22.5">
            	<Cell ss:StyleID="s62">TOPTAN SATIS</Cell>
    			<Cell ss:StyleID="s62">NET SATIŞ GEÇEN YIL</Cell>
                <Cell ss:StyleID="s62">NET SATIŞ BÜTÇE</Cell>
                <Cell ss:StyleID="s62">NET SATIS FİİLİ</Cell>
                <Cell ss:StyleID="s62">Prog.Göre Artış</Cell>
                <Cell ss:StyleID="s62">G.Yıla Göre Artış</Cell>
            </Row>
            <xsl:variable name="ROWS" select="TOPTAN_SATIS/ROWSET/ROW"/>
    		<xsl:for-each select="$Aylar">
                <xsl:variable name="TARIH_AY" select="@name"/>
                <xsl:variable name="ROW" select="$ROWS[TARIH_AY = $TARIH_AY]"/>
                    <Row ss:Height="22.5">
                    	<Cell ss:StyleID="s62"><xsl:value-of select="@title"/></Cell>
                        <xsl:call-template name="birim_values">
                            <xsl:with-param name="ROW" select="$ROW"/>
                        </xsl:call-template>
                    </Row>
            </xsl:for-each>
    	</xsl:variable>
        
        <!--table class="stats"-->
        	<xsl:call-template name="transpose">
            	<xsl:with-param name="value" select="exslt:node-set($MtXML)/*"/>
            </xsl:call-template>
        
        
    </xsl:template>
    
    <xsl:template name="sirket-toplam">
    
    	<xsl:variable name="MtXML">
    		<Row ss:Height="22.5">
            	<Cell ss:StyleID="s62">SIRKET TOPLAM</Cell>
    			<xsl:call-template name="birim_titles"/>
            </Row>
            <xsl:variable name="ROWS" select="//ROWSET/ROW"/>
    		<xsl:for-each select="$Aylar">
                <xsl:variable name="TARIH_AY" select="@name"/>
                <xsl:variable name="ROW" select="$ROWS[TARIH_AY = $TARIH_AY]"/>
                <Row ss:Height="22.5">
                	<Cell ss:StyleID="s62"><xsl:value-of select="@title"/></Cell>
                    <xsl:call-template name="birim_values">
                        <xsl:with-param name="ROW" select="$ROW"/>
                    </xsl:call-template>
                </Row>
            </xsl:for-each>
    	</xsl:variable>
        
        <!--table class="stats"-->
        	<xsl:call-template name="transpose">
            	<xsl:with-param name="value" select="exslt:node-set($MtXML)/*"/>
            </xsl:call-template>
        
        
    </xsl:template>
    
    <xsl:template name="sirket-toplam-kumule">
    
    	<xsl:variable name="MtXML">
    		<Row ss:Height="22.5">
            	<Cell ss:StyleID="s62">SIRKET TOPLAM</Cell>
    			<xsl:call-template name="birim_titles"/>
            </Row>
            <xsl:variable name="ROWS" select="//ROWSET/ROW"/>
            <Row ss:Height="22.5">
                <Cell ss:StyleID="s62">KUMULE</Cell>
                <xsl:call-template name="birim_values">
                    <xsl:with-param name="ROW" select="$ROWS"/>
                </xsl:call-template>
            </Row>
    	</xsl:variable>
        
        <!--table class="stats"-->
        	<xsl:call-template name="transpose">
            	<xsl:with-param name="value" select="exslt:node-set($MtXML)/*"/>
            </xsl:call-template>
        
        
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
         <Row ss:Height="22.5">
			<xsl:for-each select="$value">
            	<xsl:copy-of select="(node()|@*)[position()=$pos]"/>
			</xsl:for-each>
         </Row>
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