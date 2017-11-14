
#' @include class_definitions.R


WorkSheet$methods(initialize = function(showGridLines = TRUE, 
                                        tabSelected = FALSE, 
                                        tabColour = NULL, 
                                        zoom = 100, 
                                        
                                        oddHeader,
                                        oddFooter,
                                        evenHeader, 
                                        evenFooter, 
                                        firstHeader, 
                                        firstFooter,
                                        
                                        paperSize, 
                                        orientation,
                                        hdpi = 300,
                                        vdpi = 300){
  
  if(!is.null(tabColour)){
    tabColour <- sprintf('<sheetPr><tabColor rgb="%s"/></sheetPr>', tabColour)
  }else{
    tabColour <- character(0)
  }
  
  if(zoom < 10){
    zoom <- 10
  }else if(zoom > 400){
    zoom <- 400
  }
  
  naToNULLList <- function(x){
    lapply(x, function(x) {
      if(is.na(x))
        return(NULL)
      x})
  }
  
  hf <- list(oddHeader = naToNULLList(oddHeader),
             oddFooter = naToNULLList(oddFooter),
             evenHeader = naToNULLList(evenHeader),
             evenFooter = naToNULLList(evenFooter), 
             firstHeader = naToNULLList(firstHeader),
             firstFooter = naToNULLList(firstFooter))
  
  if(all(sapply(hf, length) == 0))
    hf <- list()
  
  ## list of all possible children
  sheetPr <<- tabColour
  dimension <<- '<dimension ref="A1"/>'
  sheetViews <<- sprintf('<sheetViews><sheetView workbookViewId="0" zoomScale="%s" showGridLines="%s" tabSelected="%s"/></sheetViews>', as.integer(zoom), as.integer(showGridLines), as.integer(tabSelected))
  sheetFormatPr <<- '<sheetFormatPr defaultRowHeight="15.0"/>'
  cols <<- character(0)
  
  autoFilter <<- character(0)
  mergeCells <<- character(0)
  conditionalFormatting <<- character(0)
  dataValidations <<- NULL
  hyperlinks <<- list()
  pageMargins <<- '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
  pageSetup <<- sprintf('<pageSetup paperSize="%s" orientation="%s" horizontalDpi="%s" verticalDpi="%s" r:id="rId2"/>', paperSize, orientation, hdpi, vdpi)  ## will always be 2
  headerFooter <<- hf
  rowBreaks <<- character(0)
  colBreaks <<- character(0)
  drawing <<- '<drawing r:id=\"rId1\"/>' ## will always be 1
  legacyDrawing <<- character(0)
  legacyDrawingHF <<- character(0)
  oleObjects <<- character(0)
  tableParts <<- character(0)
  extLst <<- character(0)
  
  freezePane <<- character(0)
  
  sheet_data <<- Sheet_Data$new()
  
})









WorkSheet$methods(get_prior_sheet_data = function(){
  
  xml <- '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'
  
  if(length(sheetPr) > 0){
    tmp <- sheetPr
    if(!any(grepl("<sheetPr>", tmp, fixed = TRUE)))
      tmp <- paste0("<sheetPr>", paste(tmp, collapse = ""), "</sheetPr>")
    
    xml <- paste(xml, tmp, collapse = "")
  }
  
  if(length(dimension) > 0)
    xml <- paste(xml, dimension, collapse = "")
  
  ## sheetViews handled here
  if(length(freezePane) > 0){
    xml <- paste(xml, gsub("/></sheetViews>", paste0(">", freezePane, "</sheetView></sheetViews>"), sheetViews, fixed = TRUE), collapse = "")
  }else if(length(sheetViews) > 0){
    xml <- paste(xml, sheetViews, collapse = "")
  }
  
  if(length(sheetFormatPr) > 0)
    xml <- paste(xml, sheetFormatPr, collapse = "")
  
  if(length(cols) > 0)
    xml <- paste(xml, pxml(c("<cols>", cols, "</cols>")), collapse = "")
  
  
  return(xml)  
  
})




WorkSheet$methods(get_post_sheet_data = function(){
  
  
  xml <- ""
  
  if(length(autoFilter) > 0)
    xml <- paste0(xml, autoFilter, collapse = "")
  
  if(length(mergeCells) > 0)
    xml <- paste0(xml, paste0(sprintf('<mergeCells count="%s">', length(mergeCells)), pxml(mergeCells), '</mergeCells>'), collapse = "")
  
  
  
  if(length(conditionalFormatting) > 0){
    
    nms <- names(conditionalFormatting)
    xml <- paste0(xml
                  , paste(
                    sapply(unique(nms), function(x) {
                      paste0(sprintf('<conditionalFormatting sqref="%s">', x)
                             , pxml(conditionalFormatting[nms == x])
                             , '</conditionalFormatting>')
                    })
                    , collapse = "")
                  
                  , collapse = "")
    
  }
  
  
  if(length(dataValidations) > 0)
    xml <- paste0(xml, paste0(sprintf('<dataValidations count="%s">', length(dataValidations)), pxml(dataValidations), '</dataValidations>'))
  
  
  
  if(length(hyperlinks) > 0){
    h_inds <- paste0(1:length(hyperlinks), "h")
    xml <- paste(xml, paste("<hyperlinks>", paste(sapply(1:length(h_inds), function(i) hyperlinks[[i]]$to_xml(h_inds[i])), collapse = ""), "</hyperlinks>"), collapse = "")
  }
  
  
  if(length(pageMargins) > 0)
    xml <- paste0(xml, pageMargins, collapse = "")
  
  if(length(pageSetup) > 0)
    xml <- paste0(xml, pageSetup, collapse = "")
  
  if(length(headerFooter) > 0)
    xml <- paste0(xml, genHeaderFooterNode(headerFooter), collapse = "")
  
  
  ## rowBreaks and colBreaks
  if(length(rowBreaks) > 0){
    xml <- paste0(xml
                  , paste0(sprintf('<rowBreaks count="%s" manualBreakCount="%s">', length(rowBreaks), length(rowBreaks)), paste(rowBreaks, collapse = ""),'</rowBreaks>')
                  , collapse = "")
  }
  
  
  if(length(colBreaks) > 0){
    xml <- paste0(xml
                  , paste0(sprintf('<colBreaks count="%s" manualBreakCount="%s">', length(colBreaks), length(colBreaks)), paste(colBreaks, collapse = ""),'</colBreaks>')
                  , collapse = "")
  }
  
  
  
  
  
  
  
  
  
  if(length(drawing) > 0)
    xml <- paste0(xml, drawing, collapse = "")
  
  if(length(legacyDrawing) > 0)
    xml <- paste0(xml, legacyDrawing, collapse = "")
  
  if(length(legacyDrawingHF) > 0)
    xml <- paste0(xml, legacyDrawingHF, collapse = "")
  
  if(length(oleObjects) > 0)
    xml <- paste0(xml, oleObjects, collapse = "")
  
  if(length(tableParts) > 0){
    xml <- paste0(xml
                  , paste0(sprintf('<tableParts count="%s">', length(tableParts)), pxml(tableParts), '</tableParts>')
                  , collapse = "")
  }
  
  
  if(length(extLst) > 0) 
    xml <- paste0(xml
                  , sprintf('<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:dataValidations count="%s" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">'
                  , length(extLst))
                  , paste0(pxml(extLst), '</x14:dataValidations></ext></extLst>'), collapse = "") 
  
  xml <- paste0(xml, "</worksheet>") 
  
  return(xml)
  
  
})


WorkSheet$methods(order_sheetdata = function(){
  
  if(sheet_data$n_elements == 0)
    return(invisible(0))
  
  if(sheet_data$data_count > 1){
    
    ord <- order(sheet_data$rows, sheet_data$cols, method = "radix", na.last = TRUE)
    sheet_data$rows <<- sheet_data$rows[ord]
    sheet_data$cols <<- sheet_data$cols[ord]
    sheet_data$t <<- sheet_data$t[ord]
    sheet_data$v <<- sheet_data$v[ord]
    sheet_data$f <<- sheet_data$f[ord]
    
    sheet_data$style_id <<- sheet_data$style_id[ord]
    
    sheet_data$data_count <<- 1L

    dm1 <- paste0(int_2_cell_ref(cols = sheet_data$cols[1]), sheet_data$rows[1])
    dm2 <- paste0(int_2_cell_ref(cols = sheet_data$cols[sheet_data$n_elements]), sheet_data$rows[sheet_data$n_elements])
    
    if(length(dm1) == 1 & length(dm2) != 1){
      if(!is.na(dm1) & !is.na(dm2) & dm1 != "NA" & dm2 != "NA")
        dimension <<- sprintf("<dimension ref=\"%s:%s\"/>", dm1, dm2)
    }
    
    
  }
  
  
  invisible(0)
  
})

