

#' @include class_definitions.R


ChartSheet$methods(initialize = function(tabSelected = FALSE, 
                                         tabColour = character(0), 
                                         zoom = 100){
  
  if(length(tabColour) > 0){
    tabColour <- sprintf('<sheetPr>%s</sheetPr>', tabColour)
  }else{
    tabColour <- character(0)
  }
  if(zoom < 10){
    zoom <- 10
  }else if(zoom > 400){
    zoom <- 400
  }
  
  sheetPr <<- tabColour
  sheetViews <<- sprintf('<sheetViews><sheetView workbookViewId="0" zoomScale="%s" tabSelected="%s"/></sheetViews>', as.integer(zoom), as.integer(tabSelected))
  pageMargins <<- '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
  drawing <<- '<drawing r:id=\"rId1\"/>'
  hyperlinks <<- character(0)
  
  return(invisible(0))
  
})





ChartSheet$methods(get_prior_sheet_data = function(){
  
  xml <- '<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">>'
  
  if(length(sheetPr) > 0)
    xml <- paste(xml, sheetPr, collapse = "")
  
  if(length(sheetViews) > 0)
    xml <- paste(xml, sheetViews, collapse = "")
  
  if(length(pageMargins) > 0)
    xml <- paste(xml, pageMargins, collapse = "")
  
  if(length(drawing) > 0)
    xml <- paste(xml, drawing, collapse = "")
  
  xml <- paste(xml, "</chartsheet>")
  
  return(xml)  
  
})


