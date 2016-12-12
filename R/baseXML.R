

genBaseContent_Type <- function(){
  
  c(
    '<Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"/>',
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    '<Default Extension="xml" ContentType="application/xml"/>',
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
    '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>',
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
    '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
  )
  
}


genBaseShapeVML <- function(clientData, id){
  
  if(grepl("visible", clientData, ignore.case = TRUE)){
    visible <- "visible"
  }else{
    visible <- "hidden"
  }
  
  paste0(sprintf('<v:shape id="_x0000_s%s" type="#_x0000_t202" style=\'position:absolute;', id),
         sprintf('margin-left:107.25pt;margin-top:172.5pt;width:147pt;height:96pt;z-index:1;
          visibility:%s;mso-wrap-style:tight\' fillcolor="#ffffe1" o:insetmode="auto">',  visible),
            '<v:fill color2="#ffffe1"/>
            <v:shadow color="black" obscured="t"/>
            <v:path o:connecttype="none"/>
            <v:textbox style=\'mso-direction-alt:auto\'>
            <div style=\'text-align:left\'/>
            </v:textbox>', clientData, '</v:shape>')
}





genClientData <- function(col, row, visible, height, width){

  txt <- sprintf('<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>%s, 15, %s, 10, %s, 147, %s, 18</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>%s</x:Row><x:Column>%s</x:Column>',
          col, row-2L, col + width - 1L, row + height - 1L, row-1L, col-1L)
  
  if(visible)
    txt <- paste0(txt, '<x:Visible/>')
  
  txt <- paste0(txt, '</x:ClientData>')
  
  return(txt)
  
}


# genBaseRels <- function(){
#   
#   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
#    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
#    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
#   
# }
# 
# 
# genBaseApp <- function(){
#   list('<Application>Microsoft Excel</Application>')
# }


genBaseCore <- function(creator){  
  sprintf('<dcterms:created xsi:type="dcterms:W3CDTF">2014-03-07T16:08:25Z</dcterms:created><dc:creator>%s</dc:creator>', creator)
} 


genBaseWorkbook.xml.rels <- function(){
  
  c('<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>',
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
    '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>')
}


genBaseWorkbook <- function(){
  
  list(workbookPr = '<workbookPr date1904="false"/>',
       bookViews = 	'<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="13125" windowHeight="6105"/></bookViews>',
       sheets = NULL,
       externalReferences = NULL,
       definedNames = NULL,
       calcPr = NULL,
       pivotCaches = NULL,
       extLst = NULL
  )
  
}




genBaseSheetRels <- function(sheetInd){
  
  c(
    sprintf('<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing%s.xml"/>', sheetInd),
    sprintf('<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings" Target="../printerSettings/printerSettings%s.bin"/>', sheetInd),
    sprintf('<Relationship Id="rIdvml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing%s.vml"/>', sheetInd)
  )
  
}

genBaseStyleSheet <- function(dxfs = NULL, tableStyles = NULL, extLst = NULL){
  
  list(
    
    numFmts = NULL,
    
    fonts = c('<font><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>'),
    
    fills = c('<fill><patternFill patternType="none"/></fill>',
              '<fill><patternFill patternType="gray125"/></fill>'),
    
    borders = c('<border><left/><right/><top/><bottom/><diagonal/></border>'),
    
    cellStyleXfs = c('<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'),
    
    cellXfs = c('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'),
    
    cellStyles = c('<cellStyle name="Normal" xfId="0" builtinId="0"/>'),
    
    dxfs = dxfs,
    
    tableStyles = tableStyles,
    
    indexedColors = NULL,
    
    extLst = extLst
    
  )
  
}


genBasePic <- function(imageNo){
  
  sprintf('<xdr:pic xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
      <xdr:nvPicPr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
        <xdr:cNvPr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" id="%s" name="Picture %s"/>
        <xdr:cNvPicPr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
          <a:picLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
        </xdr:cNvPicPr>
      </xdr:nvPicPr>
      <xdr:blipFill xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
        <a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId%s">
        </a:blip>
        <a:stretch xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:fillRect/>
        </a:stretch>
      </xdr:blipFill>
      <xdr:spPr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
        <a:prstGeom xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" prst="rect">
          <a:avLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>
        </a:prstGeom>
      </xdr:spPr>
    </xdr:pic>', imageNo, imageNo, imageNo)
  
}


genBaseTable <- function(id, ref, colNames){
  
  
  table <- sprintf('<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="%s" name="Table%s" displayName="Table%s" ref="%s" totalsRowShown="0">', id, id, id, ref)
  table <- c(table, sprintf('<autoFilter ref="%s"/>', ref))
  
  tableCols <- NULL
  for(i in 1:length(colNames))
    tableCols <- c(tableCols, sprintf('<tableColumn id="%s" name="%s"/>', i, colNames[[i]]))
  
  tableCols <- c(sprintf('<tableColumns count="%s">', length(colNames)), tableCols, '</tableColumns>')
  tableStyle <- '<tableStyleInfo name="TableStyleLight9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>'
  
  
  table <- paste(c(table, tableCols, tableStyle, '</table>'))
  writeLines(table)
  
  
}




genBaseChartSheet <- function(tabSelected = FALSE, 
                              tabColour = NULL, 
                              zoom = 100){
  
  if(!is.null(tabColour))
    tabColour <- sprintf('<sheetPr>%s</sheetPr>', tabColour)
  
  if(zoom < 10){
    zoom <- 10
  }else if(zoom > 400){
    zoom <- 400
  }
  
  ## list of all possible children
  tmp <- list(list(sheetPr = tabColour,
                   sheetViews = sprintf('<sheetViews><sheetView workbookViewId="0" zoomScale="%s" tabSelected="%s"/></sheetViews>', as.integer(zoom), as.integer(tabSelected)),
                   pageMargins = '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>',
                   # pageSetup = '<pageSetup paperSize="9" orientation="portrait" horizontalDpi="300" verticalDpi="300" r:id="rId2"/>',  ## will always be 2
                   drawing = '<drawing r:id=\"rId1\"/>' ## will always be 1
  ))
  
  return(tmp)
}






genBaseTheme <- function(){
  
  '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements><a:clrScheme name="Office">
  <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
  <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
  <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
  <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
  <a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2>
  <a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4>
  <a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6>
  <a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink>
  </a:clrScheme><a:fontScheme name="Office">
  <a:majorFont>
  <a:latin typeface="Cambria"/>
  <a:ea typeface=""/>
  <a:cs typeface=""/>
  <a:font script="Arab" typeface="Times New Roman"/>
  <a:font script="Hebr" typeface="Times New Roman"/>
  <a:font script="Thai" typeface="Tahoma"/>
  <a:font script="Ethi" typeface="Nyala"/>
  <a:font script="Beng" typeface="Vrinda"/>
  <a:font script="Gujr" typeface="Shruti"/>
  <a:font script="Khmr" typeface="MoolBoran"/>
  <a:font script="Knda" typeface="Tunga"/>
  <a:font script="Guru" typeface="Raavi"/>
  <a:font script="Cans" typeface="Euphemia"/>
  <a:font script="Cher" typeface="Plantagenet Cherokee"/>
  <a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
  <a:font script="Tibt" typeface="Microsoft Himalaya"/>
  <a:font script="Thaa" typeface="MV Boli"/>
  <a:font script="Deva" typeface="Mangal"/>
  <a:font script="Telu" typeface="Gautami"/>
  <a:font script="Taml" typeface="Latha"/>
  <a:font script="Syrc" typeface="Estrangelo Edessa"/>
  <a:font script="Orya" typeface="Kalinga"/>
  <a:font script="Mlym" typeface="Kartika"/>
  <a:font script="Laoo" typeface="DokChampa"/>
  <a:font script="Sinh" typeface="Iskoola Pota"/>
  <a:font script="Mong" typeface="Mongolian Baiti"/>
  <a:font script="Viet" typeface="Times New Roman"/>
  <a:font script="Uigh" typeface="Microsoft Uighur"/>
  <a:font script="Geor" typeface="Sylfaen"/>
  </a:majorFont>
  <a:minorFont>
  <a:latin typeface="Calibri"/>
  <a:ea typeface=""/>
  <a:cs typeface=""/>
  <a:font script="Arab" typeface="Arial"/>
  <a:font script="Hebr" typeface="Arial"/>
  <a:font script="Thai" typeface="Tahoma"/>
  <a:font script="Ethi" typeface="Nyala"/>
  <a:font script="Beng" typeface="Vrinda"/>
  <a:font script="Gujr" typeface="Shruti"/>
  <a:font script="Khmr" typeface="DaunPenh"/>
  <a:font script="Knda" typeface="Tunga"/>
  <a:font script="Guru" typeface="Raavi"/>
  <a:font script="Cans" typeface="Euphemia"/>
  <a:font script="Cher" typeface="Plantagenet Cherokee"/>
  <a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
  <a:font script="Tibt" typeface="Microsoft Himalaya"/>
  <a:font script="Thaa" typeface="MV Boli"/>
  <a:font script="Deva" typeface="Mangal"/>
  <a:font script="Telu" typeface="Gautami"/>
  <a:font script="Taml" typeface="Latha"/>
  <a:font script="Syrc" typeface="Estrangelo Edessa"/>
  <a:font script="Orya" typeface="Kalinga"/>
  <a:font script="Mlym" typeface="Kartika"/>
  <a:font script="Laoo" typeface="DokChampa"/>
  <a:font script="Sinh" typeface="Iskoola Pota"/>
  <a:font script="Mong" typeface="Mongolian Baiti"/>
  <a:font script="Viet" typeface="Arial"/>
  <a:font script="Uigh" typeface="Microsoft Uighur"/>
  <a:font script="Geor" typeface="Sylfaen"/>
  </a:minorFont>
  </a:fontScheme>
  <a:fmtScheme name="Office">
  <a:fillStyleLst>
  <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
  <a:gradFill rotWithShape="1">
  <a:gsLst>
  <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
  <a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
  <a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
  </a:gsLst>
  <a:lin ang="16200000" scaled="1"/>
  </a:gradFill>
  <a:gradFill rotWithShape="1">
  <a:gsLst>
  <a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs>
  <a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs>
  <a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs>
  </a:gsLst>
  <a:lin ang="16200000" scaled="0"/>
  </a:gradFill>
  </a:fillStyleLst>
  <a:lnStyleLst>
  <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
  <a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr>
  </a:solidFill>
  <a:prstDash val="solid"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">
  <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
  <a:prstDash val="solid"/></a:ln>
  <a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
  </a:lnStyleLst>
  <a:effectStyleLst>
  <a:effectStyle>
  <a:effectLst>
  <a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0">
  <a:srgbClr val="000000">
  <a:alpha val="38000"/>
  </a:srgbClr>
  </a:outerShdw>
  </a:effectLst>
  </a:effectStyle>
  <a:effectStyle>
  <a:effectLst>
  <a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
  <a:srgbClr val="000000">
  <a:alpha val="35000"/>
  </a:srgbClr>
  </a:outerShdw>
  </a:effectLst>
  </a:effectStyle>
  <a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/>
  </a:srgbClr>
  </a:outerShdw>
  </a:effectLst>
  <a:scene3d>
  <a:camera prst="orthographicFront">
  <a:rot lat="0" lon="0" rev="0"/>
  </a:camera>
  <a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig>
  </a:scene3d>
  <a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d>
  </a:effectStyle>
  </a:effectStyleLst>
  <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1">
  <a:gsLst>
  <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
  <a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/> <a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
  <a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs>
  </a:gsLst>
  <a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst>
  <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
  <a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs>
  </a:gsLst>
  <a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>
  </a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>'  
  
  
}




genPrinterSettings <- function(){
  "5c 00 5c 00 41 00 55 00 43 00 41 00 4c 00 50 00 52 00 4f 00 44 00 46 00 50 00 5c 00 4c 00 31 00 34 00 78 00 65 00 72 00 6f 00 78 00 31 00 20 00 2d 00 20 00 58 00 65 00 72 00 6f 00 00 00 00 00 01 04 00 52 dc 00 5c 05 13 ff 81 07 02 00 09 00 9a 0b 34 08 64 00 01 00 0f 00 2c 01 02 00 02 00 2c 01 03 00 01 00 41 00 34 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 01 00 00 00 01 00 00 00 01 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 52 c0 21 46 00 58 00 20 00 41 00 70 00 65 00 6f 00 73 00 50 00 6f 00 72 00 74 00 2d 00 49 00 49 00 49 00 20 00 43 00 34 00 34 00 30 00 30 00 20 00 50 00 43 00 4c 00 20 00 36 00 00 00 00 00 00 00 00 00 4e 08 a0 13 40 09 08 00 0b 01 64 00 01 00 07 00 01 00 00 00 00 00 00 00 00 00 07 00 01 00 08 01 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 08 08 08 00 08 08 08 00 08 08 08 00 08 08 08 00 00 01 03 00 02 02 00 01 02 02 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 01 00 00 00 00 00 00 00 02 02 48 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 bc 02 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 05 00 00 00 00 00 00 08 00 01 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 01 00 00 00 00 00 00 01 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 0b 96 00 00 00 c8 00 01 01 01 01 01 01 01 01 01 01 01 01 09 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 bc 02 00 00 00 00 00 00 00 00 02 00 41 00 72 00 69 00 61 00 6c 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 01 00 01 00 01 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 12 70 5f 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 01 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00"
}


genExtLst <- function(guid, sqref, posColour, negColour, values){
  
  if(is.null(values)){
    
    xml <- sprintf('<ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:conditionalFormattings><x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="dataBar" id="{%s}">
                      <x14:dataBar minLength="0" maxLength="100" border="1" negativeBarBorderColorSameAsPositive="0">
                      <x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:borderColor rgb="%s"/><x14:negativeFillColor rgb="%s"/><x14:negativeBorderColor rgb="%s"/><x14:axisColor rgb="FF000000"/>
                      </x14:dataBar></x14:cfRule><xm:sqref>%s</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext>', guid, posColour, negColour, negColour, sqref)
  }else{
    
    xml <- sprintf('<ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:conditionalFormattings><x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
                      <x14:cfRule type="dataBar" id="{%s}">
                      <x14:dataBar minLength="0" maxLength="100" border="1" negativeBarBorderColorSameAsPositive="0">
                      <x14:cfvo type="num"><xm:f>%s</xm:f></x14:cfvo><x14:cfvo type="num"><xm:f>%s</xm:f></x14:cfvo>
                      <x14:borderColor rgb="%s"/><x14:negativeFillColor rgb="%s"/><x14:negativeBorderColor rgb="%s"/><x14:axisColor rgb="FF000000"/>
                      </x14:dataBar></x14:cfRule><xm:sqref>%s</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext>', guid, values[[1]], values[[2]], posColour, negColour, negColour, sqref)
    
    
    
  }
  
  return(xml)
  
}



contentTypePivotXML <- function(i){
  
  c(sprintf('<Override PartName="/xl/pivotCache/pivotCacheDefinition%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>', i),
    sprintf('<Override PartName="/xl/pivotCache/pivotCacheRecords%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>', i),
    sprintf('<Override PartName="/xl/pivotTables/pivotTable%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>', i))
  
}

contentTypeSlicerCacheXML <- function(i){
  
  c(sprintf('<Override PartName="/xl/slicerCaches/slicerCache%s.xml" ContentType="application/vnd.ms-excel.slicerCache+xml"/>', i),
    sprintf('<Override PartName="/xl/slicers/slicer%s.xml" ContentType="application/vnd.ms-excel.slicer+xml"/>', i)
  )
  
}


genBaseSlicerXML <- function(){
  '<ext uri="{A8765BA9-456A-4dab-B4F3-ACF838C121DE}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
    <x14:slicerList>
    <x14:slicer r:id="rId0"/>
      </x14:slicerList>
      </ext>'
}


genSlicerCachesExtLst <- function(i){
  
  paste0(
    '<extLst>
    <ext uri=\"{BBE1A952-AA13-448e-AADC-164F8A28A991}\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">
    <x14:slicerCaches>',
    
    paste(sprintf('<x14:slicerCache r:id="rId%s"/>', i), collapse = ""),
    
    '</x14:slicerCaches></ext></extLst>')
  
  
}