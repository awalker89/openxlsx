


validateColour <- function(colour){
  
  ## check if
  if(is.null(colour))
    colour = "black"
  
  validColours <- colours()
  
  if(any(colour %in% validColours))
    colour[colour %in% validColours] <- col2hex(colour[colour %in% validColours])
  
  if(any(!grepl("^#[A-Fa-f0-9]{6}$", colour)))
    stop("Invalid colour!", call.=FALSE)
  
  return(toupper(colour))
  
}

## color helper function: eg col2hex(colors())
col2hex <- function(my.col) {
  rgb(t(col2rgb(my.col)), maxColorValue = 255)
}



## border helper function
doBorders <- function(borders, wb, sheet, startRow, startCol, nrow,
                      ncol, borderColour) {
  
  if("surrounding" == borders ){
    surroundingBorders(wb = wb, sheet = sheet,
                       startRow = startRow, startCol = startCol,
                       nRow = nrow, nCol = ncol,
                       borderColour = borderColour)
    
  }else if("rows" == borders ){
    rowBorders(wb = wb, sheet = sheet,
               startRow = startRow, startCol = startCol,
               nRow = nrow, nCol = ncol,
               borderColour = borderColour)
    
  }else if("columns" == borders ){
    colBorders(wb = wb, sheet = sheet,
               startRow = startRow, startCol = startCol,
               nRow = nrow, nCol = ncol,
               borderColour = borderColour)
  }
  
}



replaceIllegalCharacters <- function(v){
  
  v <- iconv(as.character(v), to = "UTF-8")
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


pxml <- function(x){
  paste(unique(unlist(x)), collapse = "")
}


removeHeadTag <- function(x){
  
  x <- paste(x, collapse = "")
  
  if(any(grepl("<\\?", x)))
    x <- gsub("<\\?xml [^>]+", "", x)
  
  x <- gsub("^>", "", x)
  x
  
}


