<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" 
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
    xmlns:msxsl="urn:schemas-microsoft-com:xslt" 
    xmlns:exslt="http://exslt.org/common" 
    xmlns:s="urn:schemas-microsoft-com:office:spreadsheet"
    xmlns:r="ramstore"
	xmlns:o="urn:schemas-microsoft-com:office:office" 
	xmlns:x="urn:schemas-microsoft-com:office:excel"
	xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
    exclude-result-prefixes="exslt msxsl s">
    
    <xsl:output method="xml" indent="no"/>
	
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
            
            
            <xsl:for-each select="/REPORT/PAGE">
            
            	<xsl:element name="PAGE" namespace="ramstore">
		            <xsl:attribute name="name"><xsl:value-of select="@name"/></xsl:attribute>
	                
		            <xsl:variable name="ROWS" select="BIRIM/ROWSET/ROW"/>
		            <xsl:variable name="TOPTAN_ROWS" select="TOPTAN_SATIS/ROWSET/ROW"/>
		            
		            <xsl:for-each select="$AylarDefinition">
		            	<xsl:variable name="TARIH_AY" select="@name"/>
		            	<xsl:variable name="TARIH_AY_NAME" select="."/>  
		            	<xsl:variable name="MONTH_ROWS" select="$ROWS[TARIH_AY=$TARIH_AY]"/>
					  	
					  	<xsl:for-each select="$MONTH_ROWS[1]"> <!-- DISTINCT MONTHS -->

						  	<xsl:element name="MONTH" namespace="ramstore">
		                    	<xsl:attribute name="name"><xsl:value-of select="$TARIH_AY"/></xsl:attribute>
		                    	<xsl:attribute name="title"><xsl:value-of select="$TARIH_AY_NAME"/></xsl:attribute>
		                    	
		                    	 <xsl:for-each select="$MONTH_ROWS/TARIH_TEXT">
		                    	 
		                    	 	<xsl:if test="generate-id(.) = generate-id($MONTH_ROWS/TARIH_TEXT[. = current()][1])"> <!-- DISTINCT DATES-->
		                    	 
			                    	 	<xsl:element name="DATES" namespace="ramstore">
				                    		<xsl:attribute name="tarih"><xsl:value-of select="current()"/></xsl:attribute>
				                    		
					                            <xsl:apply-templates select="$MONTH_ROWS[TARIH_TEXT = current()]" mode="inner"/>
					                        
			                    	 	</xsl:element><!--DATES-->
			                    	 
			                    	 </xsl:if>
			                    	 
		                    	 </xsl:for-each>
		                    	 
		                    	 <xsl:element name="TOPTAN_SATIS" namespace="ramstore">
				                    <xsl:copy-of select="$TOPTAN_ROWS[TARIH_AY=$TARIH_AY]"/>
				                </xsl:element>
						  
						 	</xsl:element><!--MONTH-->
						</xsl:for-each>
					  
					</xsl:for-each>
		            
		            <xsl:element name="BIRIMLER" namespace="ramstore">
		            	<xsl:for-each select="BIRIM">
                            <xsl:copy select="."> <!-- this copies element name -->
							   <xsl:copy-of select="@*"/> <!-- this copies all its attributes -->
							</xsl:copy>
                        </xsl:for-each>
	                </xsl:element>
		            
		            
	            
	            </xsl:element><!--PAGE-->
	            
            </xsl:for-each>
        </r:root>
    </xsl:variable>
    
    <xsl:template match="ROW" mode="inner">
      <xsl:copy>
        <xsl:attribute name="BIRIM_NO"><xsl:value-of select="../../@BIRIM_NO"/></xsl:attribute>
	    <xsl:attribute name="BIRIM_ADI"><xsl:value-of select="../../@BIRIM_ADI"/></xsl:attribute>
        <xsl:apply-templates select="@*|node()" mode="inner"/>
      </xsl:copy>
    </xsl:template>
    
    <!--Identity template copies content forward -->
    <xsl:template match="@*|node()" mode="inner">
        <xsl:copy>
            <xsl:apply-templates select="@*|node()"/>
        </xsl:copy>
    </xsl:template>
    
    <!--xsl:variable name="Aylar"  select="exslt:node-set($AylarXML)/r:root[1]/r:PAGE"/-->
    <xsl:variable name="Pages"  select="exslt:node-set($AylarXML)/r:root[1]/r:PAGE"/>
    
    
    <xsl:template match="/REPORT">
    
    	<xsl:processing-instruction name="mso-application">   
		<xsl:text>progid="Excel.Sheet"</xsl:text>  
		</xsl:processing-instruction>
        
        <!--Workbook>
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
             </ExcelWorkbook-->
		<s:Workbook xmlns:s="urn:schemas-microsoft-com:office:spreadsheet"
			 xmlns:o="urn:schemas-microsoft-com:office:office"
			 xmlns:x="urn:schemas-microsoft-com:office:excel"
			 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
			 xmlns:html="http://www.w3.org/TR/REC-html40">

            <xsl:call-template name="styles"/>
            <!--xsl:call-template name="worksheet"/-->

            <xsl:apply-templates select="PAGE"/>

      </s:Workbook>
    </xsl:template>

    <xsl:template match="PAGE" name="worksheet">
    	<s:Worksheet ss:Name="{@name}">

    		<xsl:variable name="name" select="@name"/>
    		<xsl:variable name="DATA" select="$Pages[@name=$name]"/>

            <s:Table>

               <s:Column ss:AutoFitWidth="0" ss:Width="140"/>
			    
                <s:Row>
                	<s:Cell ss:StyleID="magazaHeader"><s:Data ss:Type="String">RAMSTORE KAZAKISTAN</s:Data></s:Cell>
                	<s:Cell></s:Cell>
                	<s:Cell ss:StyleID="magazaHeader"><s:Data ss:Type="String">DONEM</s:Data></s:Cell>
                	<s:Cell><s:Data ss:Type="String"><xsl:value-of select="../@donem_from"/></s:Data></s:Cell>
                	<s:Cell><s:Data ss:Type="String"><xsl:value-of select="../@donem_to"/></s:Data></s:Cell>
                </s:Row>
                
                <xsl:call-template name="magazas">
                	<xsl:with-param name="DATA" select="$DATA/r:MONTH"/>
                </xsl:call-template>
                
                <xsl:call-template name="magaza-toplam">
                	<xsl:with-param name="DATA" select="$DATA/r:MONTH"/>
                </xsl:call-template>
                
                <xsl:call-template name="toptan-satis">
                	<xsl:with-param name="DATA" select="$DATA/r:MONTH"/>
                </xsl:call-template>
                
                <xsl:call-template name="sirket-toplam">
                	<xsl:with-param name="DATA" select="$DATA/r:MONTH"/>
                </xsl:call-template>
                
                <xsl:call-template name="sirket-toplam-kumule">
                	<xsl:with-param name="DATA" select="$DATA/r:MONTH"/>
                </xsl:call-template>
                
              </s:Table>
              
              <xsl:call-template name="footer"/>

        </s:Worksheet>
    </xsl:template>
    
    
    <xsl:template name="magazas">
    	<xsl:param name="DATA"/>
    	<xsl:param name="DATAHEAD" select="$DATA/../r:BIRIMLER/BIRIM"/>
    	
    	
    	
    	<s:Row  ss:Index="3" ss:AutoFitHeight="0" ss:Height="17.4375">
        	<s:Cell ss:StyleID="magazaHeader"><s:Data ss:Type="String"></s:Data></s:Cell>
            <s:Cell ss:MergeAcross="7" ss:StyleID="magazaHeader"><s:Data ss:Type="String">SIRKET TOPLAM</s:Data></s:Cell>
            <xsl:for-each select="$DATAHEAD">
                <s:Cell ss:MergeAcross="7" ss:StyleID="magazaHeader"><s:Data ss:Type="String"><xsl:if test="@BIRIM_NO != '118'"><xsl:value-of select="@BIRIM_NO"/> - </xsl:if><xsl:value-of select="@BIRIM_ADI"/></s:Data></s:Cell>
            </xsl:for-each>
    	</s:Row>
    	<s:Row ss:Height="35">
    		   <s:Cell ss:StyleID="magazaHeader" class="tarih"><s:Data ss:Type="String">TARIH</s:Data></s:Cell>
            	<xsl:call-template name="birim_titles"/>
            <xsl:for-each select="$DATAHEAD">
               <xsl:call-template name="birim_titles"/>
            </xsl:for-each>
    	</s:Row>
            
            <xsl:for-each select="$DATA">
            
	            <xsl:for-each select="r:DATES">
                <s:Row>
                    <s:Cell ss:StyleID="s64" class="tarih"><s:Data ss:Type="String"><xsl:value-of select="@tarih"/></s:Data></s:Cell>
                    <xsl:call-template name="birim_values">
                    	<xsl:with-param name="ROW" select="ROW"/>
                    </xsl:call-template>
                <xsl:for-each select="ROW">
                    <xsl:call-template name="birim_values"/>
                </xsl:for-each>
                </s:Row>
            	</xsl:for-each>
            
            <!--tbody class="toplam"-->
                <s:Row>
                    <s:Cell ss:StyleID="magazaHeader" class="tarih"><s:Data ss:Type="String"><xsl:value-of select="@title"/></s:Data></s:Cell>
                    <xsl:call-template name="birim_values">
                    	<xsl:with-param name="ROW" select="r:DATES/ROW"/>
                    	<xsl:with-param name="style" select="'s66'"/>
                    	<xsl:with-param name="stylePerc" select="'s67'"/>
                    </xsl:call-template>
                
                <xsl:variable name="DATES" select="r:DATES"/>
                <xsl:for-each select="r:DATES[1]/ROW">
                	<xsl:variable name="BIRIM_NO" select="@BIRIM_NO"/>
                    <xsl:call-template name="birim_values">
                    	<xsl:with-param name="ROW" select="$DATES/ROW[@BIRIM_NO = $BIRIM_NO]"/>
                    	<xsl:with-param name="style" select="'s66'"/>
                    	<xsl:with-param name="stylePerc" select="'s67'"/>
                    </xsl:call-template>
	            </xsl:for-each>
                </s:Row>
            
    		</xsl:for-each>
    		
            <!--tfoot class="toplam"-->
            
            <s:Row></s:Row>
    		<s:Row>
    			<s:Cell ss:StyleID="magazaHeader" class="tarih"><s:Data ss:Type="String">TOPLAM</s:Data></s:Cell>
                <xsl:call-template name="birim_values">
                    <xsl:with-param name="ROW" select="$DATA/r:DATES/ROW"/>
                    <xsl:with-param name="style" select="'s66'"/>
                    <xsl:with-param name="stylePerc" select="'s67'"/>
                </xsl:call-template>
                
                <xsl:for-each select="$DATAHEAD">
                	<xsl:variable name="BIRIM_NO" select="@BIRIM_NO"/>
                    <xsl:call-template name="birim_values">
                    	<xsl:with-param name="ROW" select="$DATA/r:DATES/ROW[@BIRIM_NO = $BIRIM_NO]"/>
                    	<xsl:with-param name="style" select="'s66'"/>
                    	<xsl:with-param name="stylePerc" select="'s67'"/>
                    </xsl:call-template>
	            </xsl:for-each>
                
    		</s:Row>
            
    	
        
    </xsl:template>
    
    <xsl:template name="birim_values">
   		<xsl:param name="ROW" select="."/>
   		<xsl:param name="style" select="'s63'"/>
   		<xsl:param name="stylePerc" select="'s65'"/>
   		<xsl:param name="formula" select="'h'"/>
   		
        <xsl:variable name="NET_SATIS" select="sum($ROW/NET_SATIS)"/>
        <xsl:variable name="BUTCE" select="sum($ROW/BUTCE)"/>
        <xsl:variable name="NET_SATIS_FIILI" select="sum($ROW/NET_SATIS_FIILI)"/>
        
        <!--xsl:variable name="PROG_GORE_ARTIS" select="($NET_SATIS_FIILI * 100 div $BUTCE) - 100"/>
        <xsl:variable name="GY_GORE_ARTIS" select="($NET_SATIS_FIILI * 100 div $NET_SATIS) - 100"/-->
        
        <xsl:variable name="MUSTERI_SAYISI" select="sum($ROW/MUSTERI_SAYISI)"/>
        <xsl:variable name="GY_MUSTERI_SAYISI" select="sum($ROW/GY_MUSTERI_SAYISI)"/>
        <xsl:variable name="MUSTERI_SAYISI_ARTIS" select="($MUSTERI_SAYISI * 100 div $GY_MUSTERI_SAYISI) - 100"/>
        
        <s:Cell ss:StyleID="{$style}"><s:Data ss:Type="Number"><xsl:value-of select="format-number($NET_SATIS,'#')"/></s:Data></s:Cell>
        <s:Cell ss:StyleID="{$style}"><s:Data ss:Type="Number"><xsl:value-of select="format-number($BUTCE,'#')"/></s:Data></s:Cell>
        <s:Cell ss:StyleID="{$style}"><s:Data ss:Type="Number"><xsl:value-of select="format-number($NET_SATIS_FIILI,'#')"/></s:Data></s:Cell>
        
        <xsl:choose>
        	<xsl:when test="$formula = 'h'">
        		<s:Cell ss:StyleID="{$stylePerc}" ss:Formula="=+RC[-1]/RC[-2]-1"><s:Data ss:Type="Number"></s:Data></s:Cell>
        		<s:Cell ss:StyleID="{$stylePerc}" ss:Formula="=+RC[-2]/RC[-4]-1"><s:Data ss:Type="Number"></s:Data></s:Cell>
        	</xsl:when>
            <xsl:otherwise>
            	<s:Cell ss:StyleID="{$stylePerc}" ss:Formula="=+R[-1]C/R[-2]C-1"><s:Data ss:Type="Number"></s:Data></s:Cell>
        		<s:Cell ss:StyleID="{$stylePerc}" ss:Formula="=+R[-2]C/R[-4]C-1"><s:Data ss:Type="Number"></s:Data></s:Cell>
            </xsl:otherwise>
        </xsl:choose>
        
        <xsl:if test="$ROW/MUSTERI_SAYISI">
            <s:Cell ss:StyleID="{$style}"><s:Data ss:Type="Number"><xsl:value-of select="format-number($MUSTERI_SAYISI,'#')"/></s:Data></s:Cell>
            <s:Cell ss:StyleID="{$style}"><s:Data ss:Type="Number"><xsl:value-of select="format-number($GY_MUSTERI_SAYISI,'#')"/></s:Data></s:Cell>
            
            
            <xsl:choose>
	        	<xsl:when test="$formula = 'h'">
	        		<s:Cell ss:StyleID="{$stylePerc}" ss:Formula="=+RC[-2]/RC[-1]-1"><s:Data ss:Type="Number"></s:Data></s:Cell>
	        	</xsl:when>
	            <xsl:otherwise>
	        		<s:Cell ss:StyleID="{$stylePerc}" ss:Formula="=+R[-2]C/R[-1]C-1"><s:Data ss:Type="Number"></s:Data></s:Cell>
	            </xsl:otherwise>
	        </xsl:choose>
        </xsl:if>
     </xsl:template>
     
     
     <xsl:template name="birim_titles">
     	<xsl:param name="style" select="'magazaHeader'"/>
        <s:Cell ss:StyleID="{$style}"><s:Data ss:Type="String">NET SATIŞ GEÇEN YIL</s:Data></s:Cell>
        <s:Cell ss:StyleID="{$style}"><s:Data ss:Type="String">NET SATIŞ BÜTÇE</s:Data></s:Cell>
        <s:Cell ss:StyleID="{$style}"><s:Data ss:Type="String">NET SATIS FİİLİ</s:Data></s:Cell>
        <s:Cell ss:StyleID="{$style}"><s:Data ss:Type="String">Prog.Göre Artış</s:Data></s:Cell>
        <s:Cell ss:StyleID="{$style}"><s:Data ss:Type="String">G.Yıla Göre Artış</s:Data></s:Cell>
        <s:Cell ss:StyleID="{$style}"><s:Data ss:Type="String">MUSTERI SAYISI</s:Data></s:Cell>
        <s:Cell ss:StyleID="{$style}"><s:Data ss:Type="String">MUSTERI SAYISI GEÇEN YIL</s:Data></s:Cell>
        <s:Cell ss:StyleID="{$style}"><s:Data ss:Type="String">G.Yıla Göre Artış</s:Data></s:Cell>
     </xsl:template>
     
     
    
    <xsl:template name="magaza-toplam">
    
    	<xsl:param name="DATA"/>
    
    	 <s:Row></s:Row>
    	 <s:Row></s:Row>
   		 
         <xsl:variable name="MtXML">
         <r:root>
    		<s:Row>
            	<s:Cell ss:StyleID="magazaHeader"><s:Data ss:Type="String">MAGAZA KASA SATIS TOPLAM</s:Data></s:Cell>
    			<xsl:call-template name="birim_titles">
    				<xsl:with-param name="style" select="'s68'"/>
    			</xsl:call-template>
            </s:Row>

    		<xsl:for-each select="$DATA">
                <s:Row>
                    <s:Cell ss:StyleID="magazaHeader"><s:Data ss:Type="String"><xsl:value-of select="@title"/></s:Data></s:Cell>
                    <xsl:call-template name="birim_values">
                        <xsl:with-param name="ROW" select="r:DATES/ROW"/>
                        <xsl:with-param name="formula" select="'v'"/>
                    </xsl:call-template>
                </s:Row>
            </xsl:for-each>
        </r:root>
    	</xsl:variable>

        <!--table class="stats"-->
        	<xsl:call-template name="transpose">
            	<xsl:with-param name="value" select="exslt:node-set($MtXML)/r:root/*"/>
            </xsl:call-template>
        
         
    </xsl:template>
    
    
    
    <xsl:template name="toptan-satis">
    
    	<xsl:param name="DATA"/>
    
    	<s:Row></s:Row>
    	<s:Row></s:Row>

        <xsl:variable name="MtXML">
    		<s:Row>
            	<s:Cell ss:StyleID="magazaHeader"><s:Data ss:Type="String">TOPLU SATIS</s:Data></s:Cell>
    			<s:Cell ss:StyleID="s68"><s:Data ss:Type="String">NET SATIŞ GEÇEN YIL</s:Data></s:Cell>
                <s:Cell ss:StyleID="s68"><s:Data ss:Type="String">NET SATIŞ BÜTÇE</s:Data></s:Cell>
                <s:Cell ss:StyleID="s68"><s:Data ss:Type="String">NET SATIS FİİLİ</s:Data></s:Cell>
                <s:Cell ss:StyleID="s68"><s:Data ss:Type="String">Prog.Göre Artış</s:Data></s:Cell>
                <s:Cell ss:StyleID="s68"><s:Data ss:Type="String">G.Yıla Göre Artış</s:Data></s:Cell>
            </s:Row>
            
    		<xsl:for-each select="$DATA/r:TOPTAN_SATIS/ROW">
                    <s:Row>
                    	<s:Cell ss:StyleID="magazaHeader"><s:Data ss:Type="String"><xsl:value-of select="$AylarDefinition[@name=current()/TARIH_AY]"/></s:Data></s:Cell>
                        <xsl:call-template name="birim_values">
                            <xsl:with-param name="ROW" select="current()"/>
                            <xsl:with-param name="formula" select="'v'"/>
                        </xsl:call-template>
                    </s:Row>
            </xsl:for-each>
    	</xsl:variable>
        
        <!--table class="stats"-->
        	<xsl:call-template name="transpose">
            	<xsl:with-param name="value" select="exslt:node-set($MtXML)/*"/>
            </xsl:call-template>
        
        
    </xsl:template>
    
    <xsl:template name="sirket-toplam">
    
    	<xsl:param name="DATA"/>
    
    	<s:Row></s:Row>
    	<s:Row></s:Row>
    
    	<xsl:variable name="MtXML">
    		<s:Row>
            	<s:Cell ss:StyleID="magazaHeader"><s:Data ss:Type="String">SIRKET TOPLAM</s:Data></s:Cell>
    			<xsl:call-template name="birim_titles">
    				<xsl:with-param name="style" select="'s68'"/>
    			</xsl:call-template>
            </s:Row>
            <!--xsl:variable name="ROWS" select="$DATA//ROW"/-->
    		<xsl:for-each select="$DATA">
                <!--xsl:variable name="TARIH_AY" select="@name"/>
                <xsl:variable name="ROW" select="$ROWS[TARIH_AY = $TARIH_AY]"/-->
                <s:Row>
                	<s:Cell ss:StyleID="magazaHeader"><s:Data ss:Type="String"><xsl:value-of select="@title"/></s:Data></s:Cell>
                    <xsl:call-template name="birim_values">
                        <xsl:with-param name="ROW" select="descendant::ROW"/>
                        <xsl:with-param name="formula" select="'v'"/>
                    </xsl:call-template>
                </s:Row>
            </xsl:for-each>
    	</xsl:variable>
        
        	<xsl:call-template name="transpose">
            	<xsl:with-param name="value" select="exslt:node-set($MtXML)/*"/>
            </xsl:call-template>
        
        
    </xsl:template>
    
    <xsl:template name="sirket-toplam-kumule">
    
    	<xsl:param name="DATA"/>
    
    	<s:Row></s:Row>
    	<s:Row></s:Row>
    
    	<xsl:variable name="MtXML">
    		<s:Row>
            	<s:Cell ss:StyleID="magazaHeader"><s:Data ss:Type="String">SIRKET TOPLAM</s:Data></s:Cell>
    			<xsl:call-template name="birim_titles">
    				<xsl:with-param name="style" select="'s68'"/>
    			</xsl:call-template>
            </s:Row>
            <!--xsl:variable name="ROWS" select="//ROWSET/ROW"/-->
            <s:Row>
                <s:Cell ss:StyleID="magazaHeader"><s:Data ss:Type="String">KUMULE</s:Data></s:Cell>
                <xsl:call-template name="birim_values">
                    <xsl:with-param name="ROW" select="descendant::ROW"/>
                    <xsl:with-param name="formula" select="'v'"/>
                </xsl:call-template>
            </s:Row>
    	</xsl:variable>
        
        <!--table class="stats"-->
        	<xsl:call-template name="transpose">
            	<xsl:with-param name="value" select="exslt:node-set($MtXML)/*"/>
            </xsl:call-template>
        
        
    </xsl:template>
    
    <xsl:template name="transpose">
   		 <xsl:param name="value"/>
         
         <xsl:for-each select="$value[1]/child::node()">
         <xsl:variable name="pos" select="position()"/>
         <s:Row>
			<xsl:for-each select="$value">
            	<xsl:copy-of select="(node()|@*)[position()=$pos]"/>
			</xsl:for-each>
         </s:Row>
         </xsl:for-each>
    </xsl:template>
    
    <xsl:template name="styles">    
    	<s:Styles>
		  <s:Style ss:ID="Default" ss:Name="Normal">
		   <s:Alignment ss:Vertical="Bottom"/>
		   <s:Borders/>
		   <s:Font ss:FontName="Arial"/>
		   <s:Interior/>
		   <s:NumberFormat/>
		   <s:Protection/>
		  </s:Style>
		  <s:Style ss:ID="magazaHeader">
		   <s:Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
		   <s:Borders>
		    <s:Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
		    <s:Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
		   </s:Borders>
		   <s:Font ss:FontName="Arial" ss:Size="8" ss:Bold="1"/>
		   <s:Interior ss:Color="#FFCC00" ss:Pattern="Solid"/>
		  </s:Style>
		  <s:Style ss:ID="s63">
		   <s:Alignment ss:Horizontal="Right" ss:Vertical="Bottom"/>
		   <s:Borders>
		    <s:Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
		    <s:Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
		    <s:Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
		    <s:Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
		   </s:Borders>
		   <s:Font ss:FontName="Arial" ss:Size="8"/>
		   <s:NumberFormat ss:Format="###\ ###\ ###\ ###\ ##0"/>
		  </s:Style>
		  <s:Style ss:ID="s64">
		   <s:Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
		   <s:Borders>
		    <s:Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
		    <s:Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
		    <s:Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
		    <s:Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
		   </s:Borders>
		   <s:Font ss:FontName="Arial" ss:Size="8"/>
		   <s:NumberFormat ss:Format="Standard"/>
		  </s:Style>
		  <s:Style ss:ID="s65">
		   <s:Alignment ss:Horizontal="Right" ss:Vertical="Bottom"/>
		   <s:Borders>
		    <s:Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
		    <s:Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
		    <s:Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
		    <s:Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
		   </s:Borders>
		   <s:Font ss:FontName="Arial" ss:Size="8"/>
		   <s:NumberFormat ss:Format="Percent"/>
		  </s:Style>
		  <s:Style ss:ID="s66">
		   <s:Alignment ss:Horizontal="Right" ss:Vertical="Bottom"/>
		   <s:Borders>
		    <s:Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
		    <s:Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
		    <s:Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
		    <s:Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
		   </s:Borders>
		   <s:Font ss:FontName="Arial" ss:Size="8" ss:Bold="1"/>
		   <s:Interior ss:Color="#FFCC00" ss:Pattern="Solid"/>
		   <s:NumberFormat ss:Format="###\ ###\ ###\ ###\ ##0"/>
		  </s:Style>
		  <s:Style ss:ID="s67">
		   <s:Alignment ss:Horizontal="Right" ss:Vertical="Bottom"/>
		   <s:Borders>
		    <s:Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
		    <s:Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
		    <s:Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
		    <s:Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
		   </s:Borders>
		   <s:Font ss:FontName="Arial" ss:Size="8" ss:Bold="1"/>
		   <s:Interior ss:Color="#FFCC00" ss:Pattern="Solid"/>
		   <s:NumberFormat ss:Format="Percent"/>
		  </s:Style>
		  <s:Style ss:ID="s68">
		   <s:Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
		   <s:Borders>
		    <s:Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
		    <s:Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
		    <s:Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
		    <s:Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
		   </s:Borders>
		   <s:Font ss:FontName="Arial" ss:Size="8" ss:Bold="1"/>
		   <s:Interior ss:Color="#FFCC00" ss:Pattern="Solid"/>
		  </s:Style>
		</s:Styles>
    </xsl:template>
    
    <xsl:template name="footer">    
	      <s:WorksheetOptions xmlns:s="urn:schemas-microsoft-com:office:excel">
		   <s:FreezePanes/>
		   <s:FrozenNoSplit/>
		   <s:SplitHorizontal>4</s:SplitHorizontal>
		   <s:TopRowBottomPane>35</s:TopRowBottomPane>
		   <s:SplitVertical>1</s:SplitVertical>
		   <s:LeftColumnRightPane>1</s:LeftColumnRightPane>
		   <s:ActivePane>0</s:ActivePane>
		   <s:Panes>
		    <s:Pane>
		     <s:Number>3</s:Number>
		    </s:Pane>
		    <s:Pane>
		     <s:Number>1</s:Number>
		     <s:ActiveCol>0</s:ActiveCol>
		    </s:Pane>
		    <s:Pane>
		     <s:Number>2</s:Number>
		     <s:ActiveRow>0</s:ActiveRow>
		    </s:Pane>
		    <s:Pane>
		     <Number>0</Number>
		     <s:ActiveRow>57</s:ActiveRow>
		     <s:ActiveCol>7</s:ActiveCol>
		    </s:Pane>
		   </s:Panes>
		  </s:WorksheetOptions>
		  <s:ConditionalFormatting xmlns:s="urn:schemas-microsoft-com:office:excel">
		   <s:Range>R5C2:R400C400</s:Range>
		   <s:Condition>
		    <s:Qualifier>Less</s:Qualifier>
		    <s:Value1>0</s:Value1>
		    <s:Format s:Style='color:red'/>
		   </s:Condition>
		  </s:ConditionalFormatting>
	 </xsl:template>
    
    
    <msxsl:script language="JScript" implements-prefix="exslt">
     this['node-set'] =  function (x) {
      return x;
      }
    </msxsl:script>
	
</xsl:stylesheet>