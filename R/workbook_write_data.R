
#' @include class_definitions.R



Workbook$methods(writeData = function(df, sheet, startRow, startCol, colNames, colClasses, hlinkNames, keepNA, list_sep){
  
  sheet <- validateSheet(sheet)
  nCols <- ncol(df)
  nRows <- nrow(df)  
  df_nms <- names(df)
  
  allColClasses <- unlist(colClasses)  
  df <- as.list(df)
  
  ## pull out NaN values
  nans <- unlist(lapply(1:nCols, function(i) {
    tmp <- df[[i]]
    if(!"character" %in% class(tmp) & !"list" %in% class(tmp)){
      v <- which(is.nan(tmp) | is.infinite(tmp))
      if(length(v) == 0) 
        return(v)
      return(as.integer(nCols * (v - 1) + i)) ## row position
    }
  }))
  
  ## convert any Dates to integers and create date style object
  if(any(c("date", "posixct", "posixt") %in% allColClasses)){
    
    dInds <- which(sapply(colClasses, function(x) "date" %in% x))
    
    origin <- 25569L
    if(grepl('date1904="1"|date1904="true"', paste(unlist(workbook), collapse = ""), ignore.case = TRUE))
      origin <- 24107L
    
    for(i in dInds)
      df[[i]] <- as.integer(df[[i]]) + origin
    
    
    
    pInds <- which(sapply(colClasses, function(x) any(c("posixct", "posixt", "posixlt") %in% x)))
    if(length(pInds) > 0){
      t <- sapply(pInds,  function(i) format(df[[i]][[1]], "%z"))
      
      offSet <- suppressWarnings(ifelse(substr(t,1,1) == "+", 1L, -1L) * (as.integer(substr(t,2,3)) + as.integer(substr(t,4,5)) / 60) / 24)
      
      for(i in 1:length(pInds)){
        
        if(is.na(offSet[i]))
          offSet[i] <- 0
        
        df[[pInds[i]]] <- as.numeric(as.POSIXct(df[[pInds[i]]])) / 86400 + origin + offSet[i]
        
      }
      
      
    }
    
  }
  
  ## convert any Dates to integers and create date style object
  if(any(c("currency", "accounting", "percentage", "3", "comma") %in% allColClasses)){
    cInds <- which(sapply(colClasses, function(x) any(c("accounting", "currency", "percentage", "3", "comma") %in% tolower(x))))
    for(i in cInds)
      df[[i]] <- as.numeric(gsub("[^0-9\\.-]", "", df[[i]], perl = TRUE))
  }
  
  ## convert scientific
  if("scientific" %in% allColClasses){
    for(i in which(sapply(colClasses, function(x) "scientific" %in% x)))
      class(df[[i]]) <- "numeric"
  }
  
  ##
  if("list" %in% allColClasses){
    for(i in which(sapply(colClasses, function(x) "list" %in% x)))
      df[[i]] <- sapply(lapply(df[[i]], unlist), paste, collapse = list_sep)
  }
  
  if("formula" %in% allColClasses){
    for(i in which(sapply(colClasses, function(x) "formula" %in% x))){
      df[[i]] <- replaceIllegalCharacters(as.character(df[[i]]))
      class(df[[i]]) <- "openxlsx_formula"
    }
  }
  
  if("hyperlink" %in% allColClasses){
    for(i in which(sapply(colClasses, function(x) "hyperlink" %in% x)))
      class(df[[i]]) <- "hyperlink"
  }
  
  colClasses <- sapply(df, function(x) tolower(class(x))[[1]]) ## by here all cols must have a single class only
  
  ## convert logicals (Excel stores logicals as 0 & 1)
  if("logical" %in% allColClasses){
    for(i in which(sapply(colClasses, function(x) "logical" %in% x)))
      class(df[[i]]) <- "numeric"
  }
  
  ## convert all numerics to character (this way preserves digits)
  if("numeric" %in% allColClasses){
    for(i in which(sapply(colClasses, function(x) "numeric" %in% x)))
      class(df[[i]]) <- "character"
  }
  
  ## cell types
  t <- .Call("openxlsx_buildCellTypes", colClasses, nRows, PACKAGE = "openxlsx")
  for(i in which(sapply(colClasses, function(x) !"character" %in% x & !"numeric" %in% x)))
    df[[i]] <- as.character(df[[i]])
  
  
  ## cell values
  v <- as.character(t(as.matrix(
    data.frame(df, stringsAsFactors = FALSE, check.names = FALSE, fix.empty.names = FALSE)
  ))); rm(df)
  
  if(keepNA){
    t[is.na(v)] <- "e"
    v[is.na(v)] <- "#N/A"
  }else{
    t[is.na(v)] <- as.character(NA)  
    v[is.na(v)] <- as.character(NA)
  }
  
  ## If any NaN values
  if(length(nans) > 0){
    t[nans] <- "e"
    v[nans] <- "#NUM!"
  }
  
  
  #prepend column headers 
  if(colNames){
    t <- c(rep.int('s', nCols), t)
    v <- c(df_nms, v)
    nRows <- nRows + 1L
  }
  
  ## create references
  r <- .Call("openxlsx_convert_to_excel_ref_expand", startCol:(startCol+nCols-1L), LETTERS, as.character(startRow:(startRow+nRows-1L)))
  
  ##Append hyperlinks, convert h to s in cell type
  if("hyperlink" %in% colClasses){
    
    hInds <- which(t == "h")
    
    if(length(hInds) > 0){
      t[hInds] <- "s"
      
      exHlinks <- worksheets[[sheet]]$hyperlinks
      newHlinks <- r[hInds]
      targets <- replaceIllegalCharacters(v[hInds])
      
      if(!is.null(hlinkNames) & length(hlinkNames) == length(hInds))
        v[hInds] <- hlinkNames ## this is text to display instead of hyperlink
      
      ## create hyperlink objects
      newhl <- lapply(1:length(hInds), function(i){
        Hyperlink$new(ref = newHlinks[i], target = targets[i], location = NULL, display = NULL, is_external = TRUE)
      })
      
      worksheets[[sheet]]$hyperlinks <<- append(worksheets[[sheet]]$hyperlinks, newhl)
      
    }
  }
  
  ## convert all strings to references in sharedStrings and update values (v)
  strFlag <- which(t == "s")
  newStrs <- v[strFlag]
  if(length(newStrs) > 0){
    
    newStrs <- replaceIllegalCharacters(newStrs)
    newStrs <- paste0("<si><t xml:space=\"preserve\">", newStrs, "</t></si>")
    
    uNewStr <- unique(newStrs)
    
    .self$updateSharedStrings(uNewStr)  
    v[strFlag] <- match(newStrs, sharedStrings) - 1L
  }
  
  ## Create cell list of lists
  cells <- .Call("openxlsx_buildCellList", r , t , v , PACKAGE="openxlsx")
  names(cells) <- as.integer(names(r))
  rm(t); rm(v); 
  
  if(dataCount[[sheet]] > 0){
    worksheets[[sheet]]$sheetData <<- .Call("openxlsx_unique_cell_append", worksheets[[sheet]]$sheetData, r, cells, PACKAGE = "openxlsx")
    if(attr(worksheets[[sheet]]$sheetData, "overwrite"))
      warning("Overwriting existing cell data.", call. = FALSE)
    
    attr(worksheets[[sheet]]$sheetData, "overwrite") <<- NULL
  }else{
    worksheets[[sheet]]$sheetData <<- cells
  }
  
  dataCount[[sheet]] <<- dataCount[[sheet]] + 1L
  
})






