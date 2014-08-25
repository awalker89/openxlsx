

## creates style object based on column classes
## Used in writeData and writeDataTable
classStyles <- function(wb, sheet, startRow, startCol, colNames, nRow, colClasses){
  
  sheet = wb$validateSheet(sheet)
  allColClasses <- unlist(colClasses, use.names = FALSE)
  rowInds <- 1:nRow + startRow + colNames - 1L
  startCol <- startCol - 1L
  
  newStylesElements <- NULL
  names(colClasses) <- NULL
  
  if("hyperlink" %in% allColClasses){
    
    ## style hyperlinks
    inds <- which(sapply(colClasses, function(x) "hyperlink" %in% x))
    coords <- expand.grid(rowInds, inds +startCol)   
    hyperlinkstyle <- createStyle(fontColour = "#0000FF", textDecoration = "underline")
    hyperlinkstyle$fontColour <- list("theme"="10")
    styleElements <- list(style = hyperlinkstyle,
                          cells = list(list(sheet =  names(wb$worksheets)[[sheet]],
                                            rows = coords[[1]],
                                            cols = coords[[2]])))
    
    newStylesElements <- append(newStylesElements, list(styleElements))
    
  }

  if("date" %in% allColClasses){

    ## style dates
    inds <- which(sapply(colClasses, function(x) "date" %in% x)) 
    coords <- expand.grid(rowInds, inds +startCol)   
    styleElements <- list(style = createStyle(numFmt = "date"),
                          cells = list(list(sheet =  names(wb$worksheets)[[sheet]],
                                            rows = coords[[1]],
                                            cols = coords[[2]])))
    
    newStylesElements <- append(newStylesElements, list(styleElements))
    
  }
  
  if(any(c("posixlt", "posixct", "posixt") %in% allColClasses)){
    
    ## style POSIX
    inds <- which(sapply(colClasses, function(x) any(c("posixct", "posixt", "posixlt") %in% x)))
    coords <- expand.grid(rowInds, inds +startCol)   
    
    styleElements <- list(style = createStyle(numFmt = "LONGDATE"),
                          cells = list(list(sheet =  names(wb$worksheets)[[sheet]],
                                            rows = coords[[1]],
                                            cols = coords[[2]])))
    
    newStylesElements <- append(newStylesElements, list(styleElements))
    
  }
  
  
  ## style currency as CURRENCY
  if("currency" %in% allColClasses){
    inds <- which(sapply(colClasses, function(x) "currency" %in% x))
    coords <- expand.grid(rowInds, inds +startCol)  
    
    styleElements <- list(style = createStyle(numFmt = "CURRENCY"),
                          cells = list(list(sheet =  names(wb$worksheets)[[sheet]],
                                            rows = coords[[1]],
                                            cols = coords[[2]])))
    
    newStylesElements <- append(newStylesElements, list(styleElements))
  }
  
  ## style accounting as ACCOUNTING
  if("accounting" %in% allColClasses){
    inds <- which(sapply(colClasses, function(x) "accounting" %in% x))
    coords <- expand.grid(rowInds, inds +startCol)  
    
    styleElements <- list(style = createStyle(numFmt = "ACCOUNTING"),
                          cells = list(list(sheet =  names(wb$worksheets)[[sheet]],
                                            rows = coords[[1]],
                                            cols = coords[[2]])))
    
    newStylesElements <- append(newStylesElements, list(styleElements))
    
  }
  
  ## style percentages
  if("percentage" %in% allColClasses){
    inds <- which(sapply(colClasses, function(x) "percentage" %in% x))
    coords <- expand.grid(rowInds, inds +startCol)  
    
    styleElements <- list(style = createStyle(numFmt = "percentage"),
                          cells = list(list(sheet =  names(wb$worksheets)[[sheet]],
                                            rows = coords[[1]],
                                            cols = coords[[2]])))
    
    newStylesElements <- append(newStylesElements, list(styleElements))
  }
  
  ## style big mark
  if("scientific" %in% allColClasses){
    inds <- which(sapply(colClasses, function(x) "scientific" %in% x))
    coords <- expand.grid(rowInds, inds +startCol)  
    
    styleElements <- list(style = createStyle(numFmt = "scientific"),
                          cells = list(list(sheet =  names(wb$worksheets)[[sheet]],
                                            rows = coords[[1]],
                                            cols = coords[[2]])))
    
    newStylesElements <- append(newStylesElements, list(styleElements))
  }
  
  ## style big mark
  if("3" %in% allColClasses){
    inds <- which(sapply(colClasses, function(x) "3" %in% tolower(x)))
    coords <- expand.grid(rowInds, inds +startCol)  
    
    styleElements <- list(style = createStyle(numFmt = "3"),
                          cells = list(list(sheet =  names(wb$worksheets)[[sheet]],
                                            rows = coords[[1]],
                                            cols = coords[[2]])))
    
    newStylesElements <- append(newStylesElements, list(styleElements))
  }
  
  if(!is.null(newStylesElements))
    wb$styleObjects <- append(wb$styleObjects, newStylesElements)
  
  
  invisible(1)
  
}














validateColour <- function(colour, errorMsg = "Invalid colour!"){
  
  ## check if
  if(is.null(colour))
    colour = "black"
  
  validColours <- colours()
  
  if(any(colour %in% validColours))
    colour[colour %in% validColours] <- col2hex(colour[colour %in% validColours])
  
  if(any(!grepl("^#[A-Fa-f0-9]{6}$", colour)))
    stop(errorMsg, call.=FALSE)
  
  colour <- gsub("^#", "FF", toupper(colour))
  
  return(colour)
  
}

## color helper function: eg col2hex(colors())
col2hex <- function(my.col) {
  rgb(t(col2rgb(my.col)), maxColorValue = 255)
}







replaceIllegalCharacters <- function(v){

  vEnc <- Encoding(v)
  if("UTF-8" %in% vEnc){
    fromEnc <- "UTF-8"
  }else{
    fromEnc <- ""
  }
  
  v <- iconv(as.character(v), from = fromEnc, to = "UTF-8")
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



validateBorderStyle <- function(borderStyle){
  
  
  valid <- c("none", "thin", "medium", "dashed", "dotted", "thick", "double", "hair", "mediumDashed", 
             "dashDot", "mediumDashDot", "dashDotDot", "mediumDashDotDot", "slantDashDot")
  
  ind <- match(tolower(borderStyle), tolower(valid))
  if(any(is.na(ind)))
    stop("Invalid borderStyle", call. = FALSE)
  
  return(valid[ind])

}





getAttrs <- function(xml, tag){
  
  x <- lapply(xml, function(x) .Call("openxlsx_getChildlessNode", x, tag, PACKAGE = "openxlsx"))
  x[sapply(x, length) == 0] <- ""
  a <- lapply(x, function(x) regmatches(x, regexpr('[a-zA-Z]+=".*?"', x)))
  
  names = lapply(a, function(xml) regmatches(xml, regexpr('[a-zA-Z]+(?=\\=".*?")', xml, perl = TRUE)))
  vals =  lapply(a, function(xml) regmatches(xml, regexpr('(?<=").*?(?=")', xml, perl = TRUE)))
  names(vals) <- names
  return(vals)
  
}


buildFontList <- function(fonts){
  
  sz <- getAttrs(fonts, "<sz ")
  colour <- getAttrs(fonts, "<color ")
  name <- getAttrs(fonts, "<name ")
  family <- getAttrs(fonts, "<family ")
  scheme <- getAttrs(fonts, "<scheme ")
  
  italic <- lapply(fonts, function(x) .Call("openxlsx_getChildlessNode", x, "<i", PACKAGE = "openxlsx"))
  bold <- lapply(fonts, function(x) .Call("openxlsx_getChildlessNode", x, "<b", PACKAGE = "openxlsx"))
  underline <- lapply(fonts, function(x) .Call("openxlsx_getChildlessNode", x, "<u", PACKAGE = "openxlsx"))
  
  ## Build font objects
  ft <- replicate(list(), n=length(fonts))
  for(i in 1:length(fonts)){
    
    f <- NULL
    nms <- NULL
    if(length(unlist(sz[i])) > 0){
      f <- c(f, sz[i])
      nms <- c(nms, "sz")     
    }
    
    if(length(unlist(colour[i])) > 0){
      f <- c(f, colour[i])
      nms <- c(nms, "color")  
    }
    
    if(length(unlist(name[i])) > 0){
      f <- c(f, name[i])
      nms <- c(nms, "name") 
    }
    
    if(length(unlist(family[i])) > 0){
      f <- c(f, family[i])
      nms <- c(nms, "family") 
    }
    
    if(length(unlist(scheme[i])) > 0){
      f <- c(f, scheme[i])
      nms <- c(nms, "scheme") 
    }
    
    if(length(italic[[i]]) > 0){
      f <- c(f, "italic")
      nms <- c(nms, "italic") 
    }
    
    if(length(bold[[i]]) > 0){
      f <- c(f, "bold")
      nms <- c(nms, "bold") 
    }
    
    if(length(underline[[i]]) > 0){
      f <- c(f, "underline")
      nms <- c(nms, "underline") 
    }
    
    f <- lapply(1:length(f), function(i) unlist(f[i]))
    names(f) <- nms
    
    ft[[i]] <- f
    
  }
  
  ft
  
}


nodeAttributes <- function(x){
  
  
  x <- paste0("<", unlist(strsplit(x, split = "<")))
  x <- x[grepl("<bgColor|<fgColor", x)]
  
  if(length(x) == 0)
    return("")
  
  attrs <- regmatches(x, gregexpr(' [a-zA-Z]+="[^"]*', x, perl =  TRUE))
  tags <- regmatches(x, gregexpr('<[a-zA-Z]+ ', x, perl =  TRUE))
  tags <- lapply(tags, gsub, pattern = "<| ", replacement = "")  
  attrs <- lapply(attrs, gsub, pattern = '"', replacement = "")      
  
  attrs <- lapply(attrs, strsplit, split = "=")
  for(i in 1:length(attrs)){
    nms <- lapply(attrs[[i]], "[[", 1)
    vals <- lapply(attrs[[i]], "[[", 2)
    a <- unlist(vals)
    names(a) <- unlist(nms)
    attrs[[i]] <- a
  }
  
  names(attrs) <- unlist(tags)
  
  attrs
}


buildBorder <- function(x){
  
  ## gets all borders that have children
  x <- unlist(lapply(c("<left", "<right", "<top", "<bottom"), function(tag) .Call("openxlsx_getNodes", x, tag, PACKAGE = "openxlsx")))
  if(length(x) == 0)
    return(NULL)
  
  sides <- c("TOP", "BOTTOM", "LEFT", "RIGHT")
  sideBorder <- character(length=length(x))
  for(i in 1:length(x)){
    tmp <- sides[sapply(sides, function(s) grepl(s, x[[i]], ignore.case = TRUE))]
    if(length(tmp) > 1) tmp <- tmp[[1]]
    if(length(tmp) == 1)
      sideBorder[[i]] <- tmp
  }
  
  sideBorder <- sideBorder[sideBorder != ""]
  x <- x[sideBorder != ""]
  if(length(sideBorder) == 0)
    return(NULL)
  
  
  ## style
  weight <- gsub('style=|"', "", regmatches(x, regexpr('style="[a-z]+"', x, perl = TRUE)))
  
  ## Colours
  cols <- replicate(n = length(sideBorder), list(rgb = "FF000000"))
  colNodes <- unlist(sapply(x, function(xml) .Call("openxlsx_getChildlessNode", xml, "<color", PACKAGE = "openxlsx"), USE.NAMES = FALSE))
  
  if(length(colNodes) > 0){
    attrs <- regmatches(colNodes, regexpr('(theme|indexed|rgb)=".+"', colNodes))
  }else{
    attrs <- NULL
  }
  
  if(length(attrs) != length(x)){
    return(
      list("borders" = paste(sideBorder, collapse = ""),
           "colour" = cols)
    ) 
  }
  
  attrs <- strsplit(attrs, split = "=")
  cols <- sapply(attrs, function(attr){
    y <- list(gsub('"', "", attr[[2]]))
    names(y) <- gsub(" ", "", attr[[1]])
    y
  })
  
  ## sideBorder & cols
  style <- list()
  
  if("LEFT" %in% sideBorder){
    style$borderLeft <- weight[which(sideBorder == "LEFT")]
    style$borderLeftColour <- cols[which(sideBorder == "LEFT")]
  }
  
  if("RIGHT" %in% sideBorder){
    style$borderRight <- weight[which(sideBorder == "RIGHT")]
    style$borderRightColour <- cols[which(sideBorder == "RIGHT")]
  }
  
  if("TOP" %in% sideBorder){
    style$borderTop <- weight[which(sideBorder == "TOP")]
    style$borderTopColour <- cols[which(sideBorder == "TOP")]
  }
  
  if("BOTTOM" %in% sideBorder){
    style$borderBottom <- weight[which(sideBorder == "BOTTOM")]
    style$borderBottomColour <- cols[which(sideBorder == "BOTTOM")]
  }
  
  return(style)
}



