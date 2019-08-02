



#' @export
read.xlsx.Workbook <- function(xlsxFile,
                               sheet = 1,
                               startRow = 1,
                               colNames = TRUE, 
                               rowNames = FALSE,
                               detectDates = FALSE, 
                               skipEmptyRows = TRUE, 
                               skipEmptyCols = TRUE, 
                               rows = NULL,
                               cols = NULL,
                               check.names = FALSE,
                               sep.names = ".",
                               namedRegion = NULL,
                               na.strings = "NA",
                               fillMergedCells = FALSE){
  
  
  ## Validate inputs and get files
  if(!is.logical(colNames))
    stop("colNames must be TRUE/FALSE.")
  
  if(!is.logical(rowNames))
    stop("rowNames must be TRUE/FALSE.")
  
  if(!is.logical(detectDates))
    stop("detectDates must be TRUE/FALSE.")
  
  if(!is.logical(skipEmptyRows))
    stop("skipEmptyRows must be TRUE/FALSE.")
  
  if(!is.logical(check.names))
    stop("check.names must be TRUE/FALSE.")
  
  if(!is.character(sep.names) | nchar(sep.names)!=1)
    stop("sep.names must be a character and only one.")
  
  if(length(sheet) != 1)
    stop("sheet must be of length 1.")
  
  ## Named region logic
  reading_named_region <- FALSE
  if(!is.null(namedRegion)){
    
    dn <- xlsxFile$workbook$definedNames
    if(length(dn) == 0){
      warning("Workbook has no named regions.")
      return(NULL)
    }
    
    dn_names <- replaceXMLEntities(regmatches(dn, regexpr('(?<=name=")[^"]+', dn, perl = TRUE)))
    ind <- tolower(dn_names) == tolower(namedRegion)
    if(!any(ind))
      stop(sprintf("Region '%s' not found!", namedRegion))
    
    ## pull out first node value
    dn <- dn[ind] 
    region <- regmatches(dn, regexpr('(?<=>)[^\\<]+', dn, perl = TRUE))
    sheet <- names(xlsxFile)[sapply(names(xlsxFile), function(x) grepl(x, dn))]
    if(length(sheet) > 1)
      sheet <- sheet[which.max(nchar(sheet))]
    
    region <- gsub("[^A-Z0-9:]", "", gsub(sheet, "", region, fixed = TRUE))
    
    if (grepl(":", region, fixed = TRUE)) {
      cols <- unlist(lapply(strsplit(region, split = ":", fixed = TRUE), convertFromExcelRef))
      rows <- unlist(lapply(strsplit(region, split = ":", fixed = TRUE), function(x) as.integer(gsub("[A-Z]", "", x))))
      
      cols <- seq(from = cols[1], to = cols[2], by = 1)
      rows <- seq(from = rows[1], to = rows[2], by = 1)
      
    } else {
      cols <- convertFromExcelRef(region)
      rows <- as.integer(gsub("[A-Z]", "", region, perl = TRUE))
    }
    startRow <- 1
    reading_named_region <- TRUE
    named_region_rows <- rows
    
  }
  
  
  if(is.null(rows)){
    rows <- NA
  }else if(length(rows) > 1){
    rows <- as.integer(sort(rows))
  }
  
  ## check startRow
  if(!is.null(startRow)){
    if(length(startRow) > 1)
      stop("startRow must have length 1.")
  }
  
  
  
  ## create temp dir and unzip
  nSheets <- length(xlsxFile$worksheets)
  if(nSheets == 0)
    stop("Workbook has no worksheets")
  
  ## get workbook names
  sheetNames <- xlsxFile$sheet_names
  
  if("character" %in% class(sheet)){
    sheetNames <- replaceXMLEntities(sheetNames)
    if(!sheet %in% sheetNames)
      stop(sprintf('Cannot find sheet named "%s"', sheet))
    sheet <- which(sheetNames == sheet)
  }else{
    sheet <- sheet
    if(sheet > nSheets)
      stop(sprintf("sheet %s does not exist.", sheet))
  }
  
  
  ## read in sharedStrings
  sharedStrings <- paste(unlist(xlsxFile$sharedStrings), collapse = "\n")
  if(length(sharedStrings) > 0){
    sharedStrings <- getSharedStringsFromFile(sharedStringsFile = sharedStrings, isFile = FALSE)
    if(!is.null(na.strings)){
      sharedStrings[sharedStrings %in% na.strings] <- NA
    }
  }
  
  ## read in worksheet and get cells with a value node, skip emptyStrs cells
  xlsxFile$worksheets[[sheet]]$order_sheetdata()
  sheet_data <- xlsxFile$worksheets[[sheet]]$sheet_data
  
  
  
  ######################################################
  ## What data to read
  
  
  keep <- rep.int(TRUE, length(sheet_data$rows))
  if(!is.na(rows[1]))
    keep <- keep & (sheet_data$rows %in% rows)
  
  if(!is.null(cols[1]))
    keep <- keep & (sheet_data$cols %in% cols)
  
  if(startRow > 1)
    keep <- keep & (sheet_data$rows >= startRow)
  
  ## error cells  
  keep <- keep & (sheet_data$t != 4 & !is.na(sheet_data$t) & !is.na(sheet_data$v)) ## "e" or missing
  if(any(is.na(sharedStrings)))
    keep[(sheet_data$t %in% 1 & (sheet_data$v %in% as.character(which(is.na(sharedStrings)) - 1L)))] <- FALSE
  
  ## End what data to read
  ######################################################
  
  
  
  rows <- sheet_data$rows[keep]
  cols <- sheet_data$cols[keep]
  v <- sheet_data$v[keep]
  t <- sheet_data$t[keep]
  
  if(length(v) == 0){
    warning("No data found on worksheet.", call. = FALSE)
    return(NULL)
  }
  
  
  if(is.null(rows)){
    warning("No data found on worksheet.", call. = FALSE)
    return(NULL)
  }else{
    
    if(skipEmptyRows){
      nRows <- length(unique(rows))
    }else if(reading_named_region){
      nRows <- max(named_region_rows) - min(named_region_rows) + 1;
    }else{
      nRows <- max(rows) - min(rows) + 1;
    }
    
  }
  
  ## get references for string cells
  string_refs <- which(t == 2 | t == 1) ## "b" or "s"
  if(length(string_refs) == 0)
    string_refs <- -1L
  
  
  ## get Refs for boolean 
  bool_refs <- which(t == 2) ## "b"
  if(length(bool_refs) == 0)
    bool_refs <- -1L
  
  if(bool_refs[1] != -1L){
    
    false_ind <- which(sharedStrings == "FALSE") - 1L
    if(length(false_ind) == 0){
      false_ind <- length(sharedStrings) 
      sharedStrings <- c(sharedStrings, "FALSE")
    }
    
    true_ind <- which(sharedStrings == "TRUE") - 1L
    if(length(true_ind) == 0){
      true_ind <- length(sharedStrings) 
      sharedStrings <- c(sharedStrings, "TRUE")
    }
    
    logical_vals <- v[bool_refs]
    logical_vals[logical_vals == "0"] <- false_ind[1]
    logical_vals[logical_vals == "1"] <- true_ind[1]
    v[bool_refs] <- logical_vals
    
    rm(logical_vals)
    rm(bool_refs)
    
  }
  
  
  ## If any t="str" exist, add v to sharedStrings and replace v with newSharedStringsInd
  str_inds <- which(t == 3) ## "str"
  if(length(str_inds) > 0){
    
    unique_strs <- unique(v[str_inds])
    unique_strs[unique_strs == "#N/A"] <- NA
    
    ## Match references of "str" cells to r
    new_shared_string_inds <- length(sharedStrings):(length(sharedStrings) + length(unique_strs) - 1L) 
    
    ## replace strings in v with reference to sharedStrings, (now can convert v to numeric)
    v[str_inds] <- new_shared_string_inds[match(v[str_inds], unique_strs)]
    
    ## append new strings to sharedStrings
    sharedStrings <- c(sharedStrings, unique_strs)
    if(string_refs[1] == -1L){
      string_refs <- str_inds
    }else{
      string_refs <- sort(c(string_refs, str_inds))
    }
    
  }
  
  ## Now safe to convert v to numeric
  vn <- as.numeric(v)
  
  ## Using -1 as a flag for no strings
  if(length(sharedStrings) == 0 | string_refs[1] == -1L){
    string_refs <- as.integer(NA)
  }else{
    
    ## set encoding of sharedStrings &  replace values in v with string values
    Encoding(sharedStrings) <- "UTF-8"
    v[string_refs] <- sharedStrings[vn[string_refs] + 1L]
    
    ## any NA sharedStrings - remove
    v_na <- which(is.na(v))
    if(length(v_na) > 0)
      string_refs <- setdiff(string_refs, v_na)
    
  }
  
  
  ## date detection
  origin <- 25569L
  isDate <- as.logical(NA)
  
  if(detectDates){
    
    ## get date origin
    if(length(xlsxFile$workbook$workbookPr) > 0){
      if(grepl('date1904="1"|date1904="true"', xlsxFile$workbook$workbookPr, ignore.case = TRUE))
        origin <- 24107L
    }
    
    sO <- xlsxFile$styleObjects
    sO <- sO[unlist(lapply(sO, "[[", "sheet")) == sheetNames[sheet]]
    
    
    styles <- lapply(sO, function(x) {
      fc <- x[["style"]][["numFmt"]]$formatCode
      if(is.null(fc))
        fc <- x[["style"]][["numFmt"]]$numFmtId
      fc
    })
    
    
    
    sO <- sO[sapply(styles, length) > 0]
    format_codes <- unlist(lapply(sO, function(x) {
      fc <- x[["style"]][["numFmt"]]$formatCode
      if(is.null(fc))
        fc <- x[["style"]][["numFmt"]]$numFmtId
      fc
    }))
    
    
    dateIds <- NULL
    if(length(format_codes) > 0){
      
      ## this regex defines what "looks" like a date
      format_codes <- gsub(".*(?<=\\])|@", "", format_codes, perl = TRUE)
      sO <- sO[(!grepl("[^mdyhsapAMP[:punct:] ]", format_codes) & nchar(format_codes > 3)) | format_codes == 14]
      
    }     
    
    if(length(sO) > 0){
      
      style_rows <- unlist(lapply(sO, "[[", "rows"))
      style_cols <- unlist(lapply(sO, "[[", "cols"))    
      isDate <- paste(rows, cols, sep = ",") %in% paste(style_rows, style_cols, sep = ",")
      
      ## check numbers are also integers
      not_an_integer <- suppressWarnings(as.numeric(v[isDate]))
      not_an_integer <- (not_an_integer %% 1L != 0) | is.na(not_an_integer)
      isDate[not_an_integer] <- FALSE
      
      ## perform int to date to character convertsion (way too slow)
      v[isDate] <- format(as.Date(as.integer(v[isDate]) - origin, origin = "1970-01-01"), "%Y-%m-%d")
      
    }
  } ## end of detectDates
  
  
  ## Build data.frame
  m <- read_workbook(cols_in = cols
                     , rows_in = rows
                     , v = v
                     , string_inds = string_refs
                     , is_date = isDate
                     , hasColNames = colNames
                     , hasSepNames = sep.names
                     , skipEmptyRows = skipEmptyRows
                     , skipEmptyCols = skipEmptyCols
                     , nRows = nRows
                     , clean_names = clean_names)

  if(colNames && check.names)
    colnames(m) <- make.names(colnames(m), unique = TRUE)

  if(rowNames){
    rownames(m) <- m[[1]]
    m[[1]] <- NULL
  }
  
  return(m)
  
}




