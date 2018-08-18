
#' @include class_definitions.R



Workbook$methods(writeData = function(df, sheet, startRow, startCol, colNames, colClasses, hlinkNames, keepNA, list_sep){
  
  sheet <- validateSheet(sheet)
  nCols <- ncol(df)
  nRows <- nrow(df)  
  df_nms <- names(df)
  
  allColClasses <- unlist(colClasses)  
  df <- as.list(df)
  
  ###################################################################### 
  ## standardise all column types
  
  
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
    if(length(pInds) > 0 & nRows > 0){
      t <- sapply(pInds,  function(i) {
        tzi <- format(df[[i]][[1]], "%z")
        if(is.na(tzi)){
          tz_tmp <- na.omit(df[[i]])
          tzi <- ifelse(length(tz_tmp) > 0, format(tz_tmp[1], "%z"), NA)
        }
        return(tzi)
      })
      
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
  
  ## End standardise all column types
  ###################################################################### 
  
  
  ## cell types
  t <- build_cell_types_integer(classes = colClasses, n_rows = nRows)
  
  for(i in which(sapply(colClasses, function(x) !"character" %in% x & !"numeric" %in% x)))
    df[[i]] <- as.character(df[[i]])
  
  ## cell values
  v <- as.character(t(as.matrix(
    data.frame(df, stringsAsFactors = FALSE, check.names = FALSE, fix.empty.names = FALSE)
  )));
  
  
  
  
  if (is.character(keepNA)) {
    t[is.na(v)] <- 4L
    v[is.na(v)] <- keepNA
  } else if (keepNA) {
    t[is.na(v)] <- 4L
    v[is.na(v)] <- "#N/A"
  }else{
    t[is.na(v)] <- as.integer(NA)  
    v[is.na(v)] <- as.character(NA)
  }
  
  ## If any NaN values
  if(length(nans) > 0){
    t[nans] <- 4L
    v[nans] <- "#NUM!"
  }
  
  
  #prepend column headers 
  if(colNames){
    t <- c(rep.int(1L, nCols), t)
    v <- c(df_nms, v)
    nRows <- nRows + 1L
  }
  
  
  
  ## Forumlas
  f_in <- rep.int(as.character(NA), length(t))
  any_functions <- FALSE
  if("openxlsx_formula" %in% colClasses){
    
    ## alter the elements of t where we have a formula to be "str"
    formula_cols <- which(sapply(colClasses, function(x) "openxlsx_formula" %in% x, USE.NAMES = FALSE), useNames = FALSE)
    formula_strs <- paste0("<f>", unlist(df[formula_cols], use.names = FALSE), "</f>")
    formula_inds <- unlist(lapply(formula_cols, function(i) i + (1:(nRows - colNames) - 1)*nCols + (colNames * nCols)), use.names = FALSE)
    f_in[formula_inds] <- formula_strs
    any_functions <- TRUE
    
    rm(formula_cols)
    rm(formula_strs)
    rm(formula_inds)
    
    
  }
  
  suppressWarnings(try(rm(df), silent = TRUE))
  
  ##Append hyperlinks, convert h to s in cell type
  hyperlink_cols <- which(sapply(colClasses, function(x) "hyperlink" %in% x, USE.NAMES = FALSE), useNames = FALSE)
  if(length(hyperlink_cols) > 0){
    
    hyperlink_inds <- sort(unlist(lapply(hyperlink_cols, function(i) i + (1:(nRows - colNames) - 1)*nCols + (colNames * nCols)), use.names = FALSE))
    na_hyperlink <- intersect(hyperlink_inds, which(is.na(t)))

    if(length(hyperlink_inds) > 0){
      t[t %in% 9] <- 1L ## set cell type to "s"
      
      hyperlink_refs <- convert_to_excel_ref_expand(cols = hyperlink_cols + startCol - 1, LETTERS = LETTERS, rows = as.character((startRow + colNames):(startRow+nRows - 1L)) )
      
      if(length(na_hyperlink) > 0){
        to_remove <- which(hyperlink_inds %in% na_hyperlink)
        hyperlink_refs <- hyperlink_refs[-to_remove]
        hyperlink_inds <- hyperlink_inds[-to_remove]
      }
        
      exHlinks <- worksheets[[sheet]]$hyperlinks
      targets <- replaceIllegalCharacters(v[hyperlink_inds])
      
      if(!is.null(hlinkNames) & length(hlinkNames) == length(hyperlink_inds))
        v[hyperlink_inds] <- hlinkNames ## this is text to display instead of hyperlink
      
      ## create hyperlink objects
      newhl <- lapply(1:length(hyperlink_inds), function(i){
        Hyperlink$new(ref = hyperlink_refs[i], target = targets[i], location = NULL, display = NULL, is_external = TRUE)
      })
      
      worksheets[[sheet]]$hyperlinks <<- append(worksheets[[sheet]]$hyperlinks, newhl)
      
    }
  }
  
  
  
  
  
  
  
  ## convert all strings to references in sharedStrings and update values (v)
  strFlag <- which(t == 1L)
  newStrs <- v[strFlag]
  if(length(newStrs) > 0){
    
    newStrs <- replaceIllegalCharacters(newStrs)
    newStrs <- paste0("<si><t xml:space=\"preserve\">", newStrs, "</t></si>")
    
    uNewStr <- unique(newStrs)
    
    .self$updateSharedStrings(uNewStr)  
    v[strFlag] <- match(newStrs, sharedStrings) - 1L
  }
  
  # ## Create cell list of lists
  worksheets[[sheet]]$sheet_data$write( rows_in = startRow:(startRow + nRows - 1L)
                                        , cols_in = startCol:(startCol + nCols - 1L)
                                        , t_in = t
                                        , v_in = v
                                        , f_in = f_in
                                        , any_functions = any_functions)
  
  
  
  invisible(0)
  
})






