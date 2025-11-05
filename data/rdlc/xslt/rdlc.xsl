<xsl:stylesheet version="1.0"  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"  xmlns:msxsl="urn:schemas-microsoft-com:xslt"  xmlns:xs="http://www.w3.org/2001/XMLSchema"  xmlns:msdata="urn:schemas-microsoft-com:xml-msdata"  xmlns:rd="http://schemas.microsoft.com/SQLServer/reporting/reportdesigner"  xmlns="http://schemas.microsoft.com/sqlserver/reporting/2008/01/reportdefinition"  >
  <xsl:variable name="mvarName" select="/xs:schema/@Name"/>
  <xsl:variable name="mvarTitle" select="/xs:schema/xs:element/xs:complexType/xs:choice/xs:element/xs:complexType/xs:sequence/@title"/>
  <xsl:variable name="mvarPageHeight" select="/xs:schema/xs:element/xs:complexType/xs:choice/xs:element/xs:complexType/xs:sequence/@pageheight"/>
  <xsl:variable name="mvarPageWidth" select="/xs:schema/xs:element/xs:complexType/xs:choice/xs:element/xs:complexType/xs:sequence/@pagewidth"/>
  <xsl:variable name="mvarLineWidth" select="/xs:schema/xs:element/xs:complexType/xs:choice/xs:element/xs:complexType/xs:sequence/@linewidth"/>
  <xsl:variable name="mvarFontSize" select="/xs:schema/xs:element/xs:complexType/xs:choice/xs:element/xs:complexType/xs:sequence/@fontsize"/>
  <xsl:variable name="mvarUserID" select="/xs:schema/xs:element/xs:complexType/xs:choice/xs:element/xs:complexType/xs:sequence/@userid"/>
  <xsl:variable name="mvarCompany" select="/xs:schema/xs:element/xs:complexType/xs:choice/xs:element/xs:complexType/xs:sequence/@company"/>
  <xsl:variable name="mvarBorderStyle" select="/xs:schema/xs:element/xs:complexType/xs:choice/xs:element/xs:complexType/xs:sequence/@border"/>
  <xsl:variable name="mvarShowHeader" select="/xs:schema/xs:element/xs:complexType/xs:choice/xs:element/xs:complexType/xs:sequence/@showheader"/>
  <xsl:variable name="mvarShowFooter" select="/xs:schema/xs:element/xs:complexType/xs:choice/xs:element/xs:complexType/xs:sequence/@showfooter"/>
  <xsl:variable name="mvarFontWeight">Medium</xsl:variable>
  <xsl:variable name="mvarFontWeightBold">Bold</xsl:variable>
  <xsl:template match="/">
    <xsl:apply-templates select="/xs:schema/xs:element/xs:complexType/xs:choice/xs:element/xs:complexType/xs:sequence"></xsl:apply-templates>
  </xsl:template>
  <xsl:template match="xs:sequence">
    <Report xmlns:rd="http://schemas.microsoft.com/SQLServer/reporting/reportdesigner" xmlns="http://schemas.microsoft.com/sqlserver/reporting/2008/01/reportdefinition">
      <DataSources>
        <DataSource Name="DummyDataSource">
          <ConnectionProperties>
            <DataProvider>SQL</DataProvider>
            <ConnectString>/* Local Connection */</ConnectString>
          </ConnectionProperties>
          <rd:DataSourceID>84635ff8-d177-4a25-9aa5-5a921652c79c</rd:DataSourceID>
        </DataSource>
      </DataSources>
  <xsl:call-template name="BuildDataSet"></xsl:call-template>
      <Body>
        <ReportItems>
          <Tablix Name="table1">
            <TablixBody>
              <TablixColumns>
                <xsl:apply-templates select="xs:element" mode="TablixColumn"></xsl:apply-templates>
              </TablixColumns>
              <TablixRows>
                <TablixRow>
                  <Height>0.2in</Height>
                  <TablixCells>
                    <xsl:apply-templates select="xs:element" mode="HeaderTableCell"></xsl:apply-templates>
                  </TablixCells>
                </TablixRow>
                <TablixRow>
                  <Height>0.2in</Height>
                  <TablixCells>
                    <xsl:apply-templates select="xs:element" mode="DetailTableCell"></xsl:apply-templates>
                  </TablixCells>
                </TablixRow>
              </TablixRows>
            </TablixBody>
            <TablixColumnHierarchy>
              <TablixMembers>
                <xsl:apply-templates select="xs:element" mode="TablixMember"></xsl:apply-templates>
              </TablixMembers>
            </TablixColumnHierarchy>
            <TablixRowHierarchy>
              <TablixMembers>
                <TablixMember>
                  <KeepWithGroup>After</KeepWithGroup>
                  <RepeatOnNewPage>true</RepeatOnNewPage>                  
                </TablixMember>
                <TablixMember>
                  <Group Name="RowPageBreakGroup">
                    <GroupExpressions>
                      <GroupExpression></GroupExpression>
                    </GroupExpressions>
                    <PageBreak>
                      <BreakLocation>None</BreakLocation>
                    </PageBreak>
                  </Group>
                  <TablixMembers>
                    <TablixMember>
                      <Group Name="table1_Details_Group">
                        <DataElementName>Detail</DataElementName>
                      </Group>
                      <TablixMembers>
                        <TablixMember/>
                      </TablixMembers>
                      <DataElementName>Detail_Collection</DataElementName>
                      <DataElementOutput>Output</DataElementOutput>                  
                    </TablixMember>
                  </TablixMembers>
                </TablixMember>
              </TablixMembers>
            </TablixRowHierarchy>   
	    <RepeatColumnHeaders>true</RepeatColumnHeaders>
            <RepeatRowHeaders>true</RepeatRowHeaders>
            <DataSetName>
              <xsl:value-of select="$mvarName" />
            </DataSetName>
            <Top>0cm</Top>
            <Height>0.5in</Height>
            <Width>3.74016in</Width>
            <Style>
              <Border>
                <Color>#333333</Color>
                <Style>None</Style>
                <Width>0.2mm</Width>
              </Border>
              <BottomBorder>                
                <Color>#333333</Color>
                <xsl:choose>
		  <xsl:when test="$mvarBorderStyle!=''">
                    <Style><xsl:value-of select="$mvarBorderStyle"/></Style>
	          </xsl:when>
		  <xsl:otherwise>
		    <Style>Solid</Style>
	          </xsl:otherwise>
	        </xsl:choose>
              </BottomBorder>
              <LeftBorder>
                <Color>#333333</Color>
                <xsl:choose>
		  <xsl:when test="$mvarBorderStyle!=''">
                    <Style><xsl:value-of select="$mvarBorderStyle"/></Style>
	          </xsl:when>
		  <xsl:otherwise>
		    <Style>Solid</Style>
	          </xsl:otherwise>
	        </xsl:choose>
              </LeftBorder>
              <RightBorder>
                <Color>#333333</Color>
                <xsl:choose>
		  <xsl:when test="$mvarBorderStyle!=''">
                    <Style><xsl:value-of select="$mvarBorderStyle"/></Style>
	          </xsl:when>
		  <xsl:otherwise>
		    <Style>Solid</Style>
	          </xsl:otherwise>
	        </xsl:choose>
              </RightBorder>
            </Style>
          </Tablix>
        </ReportItems>
        <Height>0.52778in</Height>
        <Style />
      </Body>
      <Width>6.85417in</Width>
      <Page>
	<xsl:if test="$mvarShowHeader='true'">
	<PageHeader>
          <Height>3.8cm</Height>
          <PrintOnFirstPage>true</PrintOnFirstPage>
          <PrintOnLastPage>true</PrintOnLastPage>
          <ReportItems>
            <Image Name="image1">
              <Source>Embedded</Source>
              <Value>rpt_logo</Value>
              <MIMEType>image/jpeg</MIMEType>
              <Sizing>FitProportional</Sizing>
              <Height>1.77941cm</Height>
	      <Width>2.4cm</Width>
              <Style />
            </Image>            
            <Textbox Name="textbox61">
              <CanGrow>true</CanGrow>
              <KeepTogether>true</KeepTogether>
              <Paragraphs>
                <Paragraph>
                  <TextRuns>
                    <TextRun>
                      <Value>="Print Date : "+Format(now(),"dd/MM/yyyy")</Value>
                      <Style>
                        <FontFamily>Arial Narrow</FontFamily>
                        <FontSize>8pt</FontSize>
                        <FontWeight>Medium</FontWeight>
                      </Style>
                    </TextRun>
                  </TextRuns>
                  <Style>
                    <TextAlign>Left</TextAlign>
                  </Style>
                </Paragraph>
              </Paragraphs>              
	            <xsl:choose>
	              <xsl:when test="$mvarCompany!=''">
		              <Top>1.61cm</Top>
		            </xsl:when>
		            <xsl:otherwise>
                  <Top>2.7875cm</Top>
		            </xsl:otherwise>
	            </xsl:choose> 
              <Left>0cm</Left> 
              <Height>0.5cm</Height>
              <Width>1.93449in</Width>
              <Style>
                <PaddingLeft>2pt</PaddingLeft>
                <PaddingRight>2pt</PaddingRight>
                <PaddingTop>0pt</PaddingTop>
                <PaddingBottom>0pt</PaddingBottom>
              </Style>
            </Textbox>
            <Textbox Name="textbox62">
              <CanGrow>true</CanGrow>
              <KeepTogether>true</KeepTogether>
              <Paragraphs>
                <Paragraph>
                  <TextRuns>
                    <TextRun>
                      <Value>=iif(Globals!TotalPages=1, "", "Page : " &amp; Globals!PageNumber &amp; " of " &amp; Globals!TotalPages)</Value>
                      <Style>
                        <FontFamily>Arial Narrow</FontFamily>
                        <FontSize>8pt</FontSize>
                        <FontWeight>Medium</FontWeight>
                      </Style>
                    </TextRun>
                  </TextRuns>
                  <Style>
                    <TextAlign>Left</TextAlign>
                  </Style>
                </Paragraph>
              </Paragraphs>
	            <xsl:choose>
	              <xsl:when test="$mvarCompany!=''">
                      <Top>2.01cm</Top>
		            </xsl:when>
		            <xsl:otherwise>
                      <Top>3.1975cm</Top>
		            </xsl:otherwise>
	            </xsl:choose>               
              <Left>0cm</Left>
              <Height>0.5cm</Height>
              <Width>1.93449in</Width>
              <Style>
                <PaddingLeft>2pt</PaddingLeft>
                <PaddingRight>2pt</PaddingRight>
                <PaddingTop>0pt</PaddingTop>
                <PaddingBottom>0pt</PaddingBottom>
              </Style>
            </Textbox>
	    <xsl:if test="$mvarCompany!=''">
              <Textbox Name="textbox82">
                <CanGrow>true</CanGrow>
                <KeepTogether>true</KeepTogether>
                <Paragraphs>
                  <Paragraph>
                    <TextRuns>
                      <TextRun>
		        <Value><xsl:value-of select="$mvarCompany" /></Value>
                        <Style>
                          <FontFamily>Arial Narrow</FontFamily>
                          <FontSize>15.5pt</FontSize>
                          <FontWeight>Bold</FontWeight>
                        </Style>
                      </TextRun>
                    </TextRuns>
                    <Style />
                  </Paragraph>
                </Paragraphs>
                <rd:DefaultName>textbox82</rd:DefaultName>
                <Top>0.02in</Top>
                <Left>2.65111cm</Left>
                <Height>0.28in</Height>
                <Width>4in</Width>     
                <Style>
                  <PaddingLeft>2pt</PaddingLeft>
                  <PaddingRight>2pt</PaddingRight>
                </Style>
              </Textbox>     
	    </xsl:if>
    	    <Textbox Name="textbox83">
              <CanGrow>true</CanGrow>
              <KeepTogether>true</KeepTogether>
              <Paragraphs>
                <Paragraph>
                  <TextRuns>
                    <TextRun>
                      <Value><xsl:value-of select="$mvarTitle" /></Value>
                      <Style>
                        <FontFamily>Arial Narrow</FontFamily>
                        <FontSize>13pt</FontSize>
                        <FontWeight>Bold</FontWeight>
                      </Style>
                    </TextRun>
                  </TextRuns>
                  <Style />
                </Paragraph>
              </Paragraphs>
              <rd:DefaultName>textbox83</rd:DefaultName>
	            <xsl:choose>
	              <xsl:when test="$mvarCompany!=''">
		              <Top>0.34in</Top>
		            </xsl:when>
		            <xsl:otherwise>
                  <Top>0.77in</Top>
		            </xsl:otherwise>
	            </xsl:choose>                 
              <Left>0cm</Left>
              <Height>0.22in</Height>
              <Width>4in</Width>
              <Style>
                <PaddingLeft>2pt</PaddingLeft>
                <PaddingRight>2pt</PaddingRight>         
              </Style>
            </Textbox>     
          </ReportItems>
          <Style>
            <Border>
              <Style>None</Style>
            </Border>            
            <BottomBorder>
              <Color>#A0A0A0</Color>
              <Style>None</Style>
              <Width>0.3mm</Width>
            </BottomBorder>            
          </Style>
        </PageHeader>
        </xsl:if>
	<xsl:if test="$mvarShowFooter='true'">
        <PageFooter>
          <Height>1.15cm</Height>
          <PrintOnFirstPage>true</PrintOnFirstPage>
          <PrintOnLastPage>true</PrintOnLastPage>
          <ReportItems>
            <Line Name="FooterLine">
              <Top>0.45cm</Top>
              <Height>0cm</Height>
              <Width>
                <xsl:value-of select="$mvarLineWidth"/>
              </Width>
              <Style>
                <Border>
                  <Color>#a0a0a0</Color>
                  <Style>Solid</Style>
                  <Width>0.3mm</Width>
                </Border>
              </Style>
            </Line>
            <Textbox Name="FooterTextBox">
              <CanGrow>true</CanGrow>
              <KeepTogether>true</KeepTogether>
              <Paragraphs>
                <Paragraph>
                  <TextRuns>
                    <TextRun>
                      <Value>
                        ="<xsl:value-of select="$mvarTitle" /> generated by : <xsl:value-of select="$mvarUserID" /> on "+Format(now(),"dd/MM/yyyy")
                      </Value>
                      <Style>
                        <FontFamily>Arial Narrow</FontFamily>
                        <FontSize>8pt</FontSize>
                        <FontWeight>Medium</FontWeight>
                      </Style>
                    </TextRun>
                  </TextRuns>
                  <Style>
		    <TextAlign>Center</TextAlign>
		  </Style>
                </Paragraph>
              </Paragraphs>
              <rd:DefaultName>FooterTextBox</rd:DefaultName>
              <Top>0.5cm</Top>
              <Left>0cm</Left>
              <Height>0.6cm</Height>
              <Width><xsl:value-of select="$mvarLineWidth"/></Width>
              <Style>
                <Border>
                  <Style>None</Style>
                </Border>
                <PaddingLeft>2pt</PaddingLeft>
                <PaddingRight>2pt</PaddingRight>
                <PaddingTop>2pt</PaddingTop>
                <PaddingBottom>2pt</PaddingBottom>
              </Style>
            </Textbox>
          </ReportItems>
          <Style>
            <Border>
              <Style>None</Style>
            </Border>
          </Style>
        </PageFooter>
	</xsl:if>
        <PageHeight><xsl:value-of select="$mvarPageHeight"/></PageHeight>
        <PageWidth><xsl:value-of select="$mvarPageWidth"/></PageWidth>
        <LeftMargin>0.3cm</LeftMargin>
        <RightMargin>0.3cm</RightMargin>
        <TopMargin>0.6cm</TopMargin>
        <BottomMargin>0.4cm</BottomMargin>
        <ColumnSpacing>0.1cm</ColumnSpacing>
        <Style />
      </Page>
      <EmbeddedImages>
        <EmbeddedImage Name="rpt_logo">
          <MIMEType>image/jpeg</MIMEType>
         <ImageData>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</ImageData>
        </EmbeddedImage>
      </EmbeddedImages>
      <rd:ReportID>b1ba4c3f-0785-48ea-9cde-36d93e5436ab</rd:ReportID>
      <rd:ReportUnitType>Cm</rd:ReportUnitType>
    </Report>
  </xsl:template>
  <xsl:template match="xs:element" mode="TablixColumn">
    <TablixColumn>
      <Width>
        <xsl:value-of select="@width"/>
      </Width>
    </TablixColumn>
  </xsl:template>
  <xsl:template match="xs:element" mode="TablixMember">
    <TablixMember />    
  </xsl:template>  
  <xsl:template match="xs:element" mode="HeaderTableCell">  
  <xsl:variable name="varFieldName"><xsl:value-of select="@name" /></xsl:variable>
  <xsl:choose>
    <xsl:when test="@cellMode='MergePrevious'">
      <TablixCell />
    </xsl:when>
    <xsl:otherwise>   
      <TablixCell>
        <CellContents>
	  <xsl:choose>
	     <xsl:when test="@cellMode='MergeDown'">
                  <Rectangle Name="rect_{$varFieldName}">
		     <KeepTogether>true</KeepTogether>
		     <ReportItems>
		        <Textbox Name="header_{$varFieldName}_above">
			    <CanGrow>true</CanGrow>
			    <KeepTogether>true</KeepTogether>
			    <Paragraphs>
			      <Paragraph>
				<TextRuns>
				  <TextRun>                  
				    <Value><xsl:value-of select="@label2"/></Value>
				    <Style>
				      <FontFamily>Arial Narrow</FontFamily>
				      <FontSize><xsl:value-of select="$mvarFontSize"/></FontSize>       
				      <FontWeight><xsl:value-of select="$mvarFontWeightBold"/></FontWeight>       
				      <Color>#000000</Color>
				    </Style>
				  </TextRun>
				</TextRuns>
				<Style>
				  <xsl:choose>
				    <xsl:when test="@hd_align='Left'">
				      <TextAlign>Left</TextAlign>
				    </xsl:when>
				    <xsl:when test="@hd_align='Right'">
				      <TextAlign>Right</TextAlign>
				    </xsl:when>
				    <xsl:when test="@hd_align='Center'">
				      <TextAlign>Center</TextAlign>
				    </xsl:when>
				    <xsl:otherwise>
				      <TextAlign>Left</TextAlign>
				    </xsl:otherwise>
				  </xsl:choose>
				</Style>
			      </Paragraph>
			    </Paragraphs>
			    <rd:DefaultName>header_<xsl:value-of select="$varFieldName"/></rd:DefaultName>    			    
			    <Height><xsl:value-of select="@hd_above_height"/></Height>			   
			    <Width><xsl:value-of select="@hd_width"/></Width>
			    <Style>
			      <Border>
				<Color>#777777</Color>
			        <Style>None</Style>
				<Width>0.16mm</Width>
			      </Border>
			      <RightBorder>
				<Color>#777777</Color>
			        <Style>Solid</Style>
				<Width>0.16mm</Width>
			      </RightBorder>
			      <BottomBorder>
				<Color>#777777</Color>
			        <Style>Solid</Style>
				<Width>0.16mm</Width>
			      </BottomBorder>
			      <BackgroundColor>#B6BCC4</BackgroundColor>
			      <VerticalAlign>Middle</VerticalAlign>
			      <PaddingLeft>2pt</PaddingLeft>
			      <PaddingRight>2pt</PaddingRight>
			      <PaddingTop>2pt</PaddingTop>
			      <PaddingBottom>2pt</PaddingBottom>
			    </Style>
			  </Textbox>
			  <xsl:choose>
			    <xsl:when test="MergeColumns/Columns">
			    <xsl:for-each select="MergeColumns/Columns">
    			        <xsl:variable name="innerFieldName"><xsl:value-of select="@name" /></xsl:variable>
				<Textbox Name="header_{$innerFieldName}_below">
				  <CanGrow>false</CanGrow>
				  <KeepTogether>true</KeepTogether>
				  <Paragraphs>
				    <Paragraph>
				      <TextRuns>
					<TextRun>                  
					  <Value><xsl:value-of select="@label"/></Value>
					  <Style>
					    <FontFamily>Arial Narrow</FontFamily>
					    <FontSize><xsl:value-of select="$mvarFontSize"/></FontSize>       
					    <FontWeight><xsl:value-of select="$mvarFontWeightBold"/></FontWeight>       
					    <Color>#000000</Color>
					  </Style>
					</TextRun>
				      </TextRuns>
				      <Style>
					<xsl:choose>
					  <xsl:when test="@align='Left'">
					    <TextAlign>Left</TextAlign>
					  </xsl:when>
					  <xsl:when test="@align='Right'">
					    <TextAlign>Right</TextAlign>
					  </xsl:when>
					  <xsl:when test="@align='Center'">
					    <TextAlign>Center</TextAlign>
					  </xsl:when>
					  <xsl:otherwise>
					    <TextAlign>Left</TextAlign>
					  </xsl:otherwise>
					</xsl:choose>
				      </Style>
				    </Paragraph>
				  </Paragraphs>
				  <rd:DefaultName>header_<xsl:value-of select="$innerFieldName"/></rd:DefaultName>    
				  <Top>0.508cm</Top>
				  <Left><xsl:value-of select="@left"/></Left>
				  <Width><xsl:value-of select="@width"/></Width>
				  <xsl:choose>
				      <xsl:when test="@height!=''">
					<Height><xsl:value-of select="@height"/></Height>
				      </xsl:when>			   
				      <xsl:otherwise>
					<Height>0.508cm</Height>
				      </xsl:otherwise>
				    </xsl:choose>
				  <Style>	
				    <BottomBorder>
				      <Color>#777777</Color>
			              <Style>None</Style>
				      <Width>0.16mm</Width>
			            </BottomBorder>
				    <RightBorder>
				      <Color>#777777</Color>
			              <Style>Solid</Style>
				      <Width>0.16mm</Width>
			            </RightBorder>
				    <LeftBorder>
				      <Color>#777777</Color>
			              <Style>Solid</Style>
				      <Width>0.12mm</Width>
			            </LeftBorder>
				    <TopBorder>
				      <Color>#777777</Color>
			              <Style>Solid</Style>
				      <Width>0.16mm</Width>
			            </TopBorder>
				    <BackgroundColor>#B6BCC4</BackgroundColor>
				    <VerticalAlign>Middle</VerticalAlign>
				    <PaddingLeft>2pt</PaddingLeft>
				    <PaddingRight>2pt</PaddingRight>
				    <PaddingTop>2pt</PaddingTop>
				    <PaddingBottom>2pt</PaddingBottom>
				  </Style>
				</Textbox>
			    </xsl:for-each>
			    </xsl:when>
			    <xsl:otherwise>
				<Textbox Name="header_{$varFieldName}_below">
				  <CanGrow>true</CanGrow>
				  <KeepTogether>true</KeepTogether>
				  <Paragraphs>
				    <Paragraph>
				      <TextRuns>
					<TextRun>                  
					  <Value><xsl:value-of select="@label"/></Value>
					  <Style>
					    <FontFamily>Arial Narrow</FontFamily>
					    <FontSize><xsl:value-of select="$mvarFontSize"/></FontSize>       
					    <FontWeight><xsl:value-of select="$mvarFontWeightBold"/></FontWeight>       
					    <Color>#000000</Color>
					  </Style>
					</TextRun>
				      </TextRuns>
				      <Style>
					<xsl:choose>
				          <xsl:when test="@hd_align='Left'">
				            <TextAlign>Left</TextAlign>
				          </xsl:when>
				          <xsl:when test="@hd_align='Right'">
				            <TextAlign>Right</TextAlign>
				          </xsl:when>
				          <xsl:when test="@hd_align='Center'">
				            <TextAlign>Center</TextAlign>
				          </xsl:when>
				          <xsl:otherwise>
				            <TextAlign>Left</TextAlign>
				          </xsl:otherwise>
				        </xsl:choose>
				      </Style>
				    </Paragraph>
				  </Paragraphs>
				  <rd:DefaultName>header_<xsl:value-of select="$varFieldName"/></rd:DefaultName>    
				  <Top>0.508cm</Top>
				  <Height><xsl:value-of select="@hd_below_height"/></Height>
				  <Style>						      
				    <BackgroundColor>#B6BCC4</BackgroundColor>
				    <VerticalAlign>Middle</VerticalAlign>
				    <PaddingLeft>2pt</PaddingLeft>
				    <PaddingRight>2pt</PaddingRight>
				    <PaddingTop>2pt</PaddingTop>
				    <PaddingBottom>2pt</PaddingBottom>
				  </Style>
				</Textbox>
			    </xsl:otherwise>
			  </xsl:choose>			  
		     </ReportItems>
		     <Height>1.016cm</Height>
		     <Style>
		      <BackgroundColor>#B6BCC4</BackgroundColor>
		      <Border>
			<Color>#777777</Color>
			<Style>Solid</Style>
			<Width>0.16mm</Width>
		      </Border>
		    </Style>
		  </Rectangle>
	     </xsl:when>
             <xsl:otherwise>
		  <Textbox Name="header_{$varFieldName}">
		    <CanGrow>true</CanGrow>
		    <KeepTogether>true</KeepTogether>
		    <Paragraphs>
		      <Paragraph>
			<TextRuns>
			  <TextRun>                  
			    <Value><xsl:value-of select="@label"/></Value>
			    <Style>
			      <FontFamily>Arial Narrow</FontFamily>
			      <FontSize><xsl:value-of select="$mvarFontSize"/></FontSize>       
			      <FontWeight><xsl:value-of select="$mvarFontWeightBold"/></FontWeight>       
			      <Color>#000000</Color>
			    </Style>
			  </TextRun>
			</TextRuns>
			<Style>
			  <xsl:choose>
			    <xsl:when test="@hd_align='Left'">
			      <TextAlign>Left</TextAlign>
			    </xsl:when>
			    <xsl:when test="@hd_align='Right'">
			      <TextAlign>Right</TextAlign>
			    </xsl:when>
 			    <xsl:when test="@hd_align='Center'">
			      <TextAlign>Center</TextAlign>
			    </xsl:when>
			    <xsl:otherwise>
			      <TextAlign>Left</TextAlign>
			    </xsl:otherwise>
			  </xsl:choose>
			</Style>
		      </Paragraph>
		    </Paragraphs>
		    <rd:DefaultName>header_<xsl:value-of select="$varFieldName"/></rd:DefaultName>                 
		    <Style>		      
                      <Border>
		        <Color>#777777</Color>
		        <Style>None</Style>
		       <Width>0.16mm</Width>
		      </Border>
		      <BottomBorder>
		        <xsl:choose>
			  <xsl:when test="$mvarBorderStyle!=''">
			    <xsl:choose>
      			      <xsl:when test="$mvarBorderStyle='None'">
				  <Style>Solid</Style>
			      </xsl:when>
			      <xsl:otherwise>
        			  <Style><xsl:value-of select="$mvarBorderStyle"/></Style>
			      </xsl:otherwise>
			    </xsl:choose>
			  </xsl:when>
			  <xsl:otherwise>
			    <Style>Solid</Style>
			  </xsl:otherwise>
		         </xsl:choose>
		      </BottomBorder>
		      <TopBorder>
		        <xsl:choose>
			  <xsl:when test="$mvarBorderStyle!=''">
			    <Style><xsl:value-of select="$mvarBorderStyle"/></Style>
			  </xsl:when>
			  <xsl:otherwise>
			    <Style>Solid</Style>
			  </xsl:otherwise>
		         </xsl:choose>
		      </TopBorder>
		      <LeftBorder>
		        <xsl:choose>
			  <xsl:when test="$mvarBorderStyle!=''">
			    <Style><xsl:value-of select="$mvarBorderStyle"/></Style>
			  </xsl:when>
			  <xsl:otherwise>
			    <Style>Solid</Style>
			  </xsl:otherwise>
		         </xsl:choose>
		      </LeftBorder>
		      <RightBorder>
		        <xsl:choose>
			  <xsl:when test="$mvarBorderStyle!=''">
			    <Style><xsl:value-of select="$mvarBorderStyle"/></Style>
			  </xsl:when>
			  <xsl:otherwise>
			    <Style>Solid</Style>
			  </xsl:otherwise>
		         </xsl:choose>
		      </RightBorder>
		      <BackgroundColor>#B6BCC4</BackgroundColor>
		      <VerticalAlign>Middle</VerticalAlign>
		      <PaddingLeft>2pt</PaddingLeft>
		      <PaddingRight>2pt</PaddingRight>
		      <PaddingTop>2pt</PaddingTop>
		      <PaddingBottom>2pt</PaddingBottom>
		    </Style>
		  </Textbox>
	      </xsl:otherwise>
	   </xsl:choose>
           <xsl:if test="@colspan &gt; 1">
	     <ColSpan><xsl:value-of select="@colspan"/></ColSpan>	   
	   </xsl:if>			   
        </CellContents>
      </TablixCell> 
    </xsl:otherwise>
  </xsl:choose>      
  </xsl:template>
  <xsl:template match="xs:element" mode="DetailTableCell">
  <xsl:variable name="varFieldName">
  <xsl:value-of select="@name" /></xsl:variable>
    <TablixCell>
      <CellContents>
        <Textbox Name="{$varFieldName}">
          <CanGrow>true</CanGrow>
          <KeepTogether>true</KeepTogether>
          <Paragraphs>
            <Paragraph>
              <TextRuns>
                <TextRun>
                  <Value>=Fields!<xsl:value-of select="$varFieldName"/>.Value</Value>
                  <Style>
                    <FontFamily>Arial Narrow</FontFamily>                  
                    <FontSize>
                      <xsl:value-of select="$mvarFontSize"/>
                    </FontSize>
                    <FontWeight>
                      <xsl:value-of select="$mvarFontWeight"/>
                    </FontWeight>
		                <xsl:choose>
                      <xsl:when test="@type='xs:dateTime'">
                        <xsl:choose>
                          <xsl:when test="@dateformat!=''">
                            <Format><xsl:value-of select="@dateformat"/></Format>
                          </xsl:when>
                        </xsl:choose>                                               
                      </xsl:when>
                    </xsl:choose>      
                  </Style>
                </TextRun>
              </TextRuns>
              <Style>
                <xsl:choose>
                  <xsl:when test="@align='Left'">
                    <TextAlign>Left</TextAlign>
                  </xsl:when>
                  <xsl:when test="@align='Right'">
                    <TextAlign>Right</TextAlign>
                  </xsl:when>
                  <xsl:when test="@align='Center'">
                    <TextAlign>Center</TextAlign>
                  </xsl:when>
                  <xsl:otherwise>
                    <TextAlign>Left</TextAlign>
                  </xsl:otherwise>
                </xsl:choose>
              </Style>
            </Paragraph>
          </Paragraphs>
          <rd:DefaultName><xsl:value-of select="$varFieldName"/></rd:DefaultName>        
          <Style>
            <Border>
              <Color>#333333</Color>
              <Style>None</Style>
              <Width>0.16mm</Width>
            </Border>
            <TopBorder>
              <xsl:choose>
		<xsl:when test="$mvarBorderStyle!=''">
                  <Style><xsl:value-of select="$mvarBorderStyle"/></Style>
	        </xsl:when>
		<xsl:otherwise>
		  <Style>Solid</Style>
	        </xsl:otherwise>
	       </xsl:choose>
            </TopBorder>
            <BottomBorder>
              <xsl:choose>
		<xsl:when test="$mvarBorderStyle!=''">
                  <Style><xsl:value-of select="$mvarBorderStyle"/></Style>
	        </xsl:when>
		<xsl:otherwise>
		  <Style>Solid</Style>
	        </xsl:otherwise>
	       </xsl:choose>
            </BottomBorder>
            <LeftBorder>
              <xsl:choose>
		<xsl:when test="$mvarBorderStyle!=''">
                  <Style><xsl:value-of select="$mvarBorderStyle"/></Style>
	        </xsl:when>
		<xsl:otherwise>
		  <Style>Solid</Style>
	        </xsl:otherwise>
	       </xsl:choose>
            </LeftBorder>
            <RightBorder>
              <xsl:choose>
		<xsl:when test="$mvarBorderStyle!=''">
                  <Style><xsl:value-of select="$mvarBorderStyle"/></Style>
	        </xsl:when>
		<xsl:otherwise>
		  <Style>Solid</Style>
	        </xsl:otherwise>
	       </xsl:choose>
            </RightBorder>
            <PaddingLeft>2pt</PaddingLeft>
            <PaddingRight>2pt</PaddingRight>
            <PaddingTop>2pt</PaddingTop>
            <PaddingBottom>2pt</PaddingBottom>
          </Style>
        </Textbox>
      </CellContents>
    </TablixCell> 
   </xsl:template>
    <xsl:template name="BuildDataSet">
      <DataSets>
        <DataSet Name="{$mvarName}">        
          <Fields>
            <xsl:apply-templates select="xs:element" mode="Field"></xsl:apply-templates>
          </Fields>
          <Query>
            <CommandText>/* Local Query */</CommandText>
            <DataSourceName>DummyDataSource</DataSourceName>
          </Query>
        </DataSet>
      </DataSets>
    </xsl:template>
    <xsl:template match="xs:element" mode="Field">
      <xsl:variable name="varFieldName">
        <xsl:value-of select="@name" />
      </xsl:variable>
      <xsl:variable name="varDataType">
        <xsl:choose>
          <xsl:when test="@type='xs:int'">System.Int32</xsl:when>
          <xsl:when test="@type='xs:string'">System.String</xsl:when>
          <xsl:when test="@type='xs:dateTime'">System.DateTime</xsl:when>
          <xsl:when test="@type='xs:boolean'">System.Boolean</xsl:when>
        </xsl:choose>
      </xsl:variable>
      <Field Name="{$varFieldName}">
        <rd:TypeName>
          <xsl:value-of select="$varDataType"/>
        </rd:TypeName>
        <DataField>
          <xsl:value-of select="$varFieldName"/>
        </DataField>
      </Field>
    </xsl:template>
    <xsl:template name="replace-string">
      <xsl:param name="text"/>
      <xsl:param name="from"/>
      <xsl:param name="to"/>
      <xsl:choose>
        <xsl:when test="contains($text, $from)">
          <xsl:variable name="before" select="substring-before($text, $from)"/>
          <xsl:variable name="after" select="substring-after($text, $from)"/>
          <xsl:variable name="prefix" select="concat($before, $to)"/>
          <xsl:value-of select="$before"/>
          <xsl:value-of select="$to"/>
          <xsl:call-template name="replace-string">
            <xsl:with-param name="text" select="$after"/>
            <xsl:with-param name="from" select="$from"/>
            <xsl:with-param name="to" select="$to"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$text"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:template>
</xsl:stylesheet>