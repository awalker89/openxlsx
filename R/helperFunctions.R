


validateBorderColour <- function(borderColour){
  
  ## check if
  if(is.null(borderColour))
    borderColour = "black"
  
  validColours <- colours()
  
  if(any(borderColour %in% validColours))
    borderColour[borderColour %in% validColours] <- col2hex(borderColour[borderColour %in% validColours])
  
  if(any(!grepl("^#[A-Fa-f0-9]{6}$", borderColour)))
    stop("Invalid borderColour!", call.=FALSE)
  
  return(toupper(borderColour))
  
}

## color helper function: eg col2hex(colors())
col2hex <- function(my.col) {
  rgb(t(col2rgb(my.col)), maxColorValue = 255)
}



## border helper function
doBorders <- function(borders, wb, sheet, srow, scol, nrow,
                      ncol, borderColour) {
  
  if("surrounding" == borders ){
    surroundingBorders(wb = wb, sheet = sheet,
                       startRow = srow, startCol = scol,
                       nRow = nrow, nCol = ncol,
                       borderColour = borderColour)
    
  }else if("rows" == borders ){
    rowBorders(wb = wb, sheet = sheet,
               startRow = srow, startCol = scol,
               nRow = nrow, nCol = ncol,
               borderColour = borderColour)
    
  }else if("columns" == borders ){
    colBorders(wb = wb, sheet = sheet,
               startRow = srow, startCol = scol,
               nRow = nrow, nCol = ncol,
               borderColour = borderColour)
  }
  
}



replaceIllegalCharacters <- function(v){
  
  v <- gsub('&', "&amp;", v)
  v <- gsub('"', "&quot;", v)
  v <- gsub("'", "&apos;", v)
  v <- gsub('<', "&lt;", v)
  v <- gsub('>', "&gt;", v)
  
  return(v)
}


replaceXMLEntities <- function(v){
  
  v <- gsub("&amp;", "&", v)
  v <- gsub("&quot;", '"', v)
  v <- gsub("&apos;", "'", v)
  v <- gsub("&lt;", "<", v)
  v <- gsub("&gt;", ">", v)
  
  return(v)
}



