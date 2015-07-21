


#' @name read.xlsx
#' @title Read data from an Excel file or Workbook object into a data.frame
#' @param xlsxFile An xlsx file or Workbook object
#' @param sheet The name or index of the sheet to read data from.
#' @param startRow first row to begin looking for data.  Empty rows at the top of a file are always skipped, 
#' regardless of the value of startRow.
#' @param colNames If \code{TRUE}, the first row of data will be used as column names. 
#' @param skipEmptyRows If \code{TRUE}, empty rows are skipped else empty rows after the first row containing data 
#' will return a row of NAs.
#' @param rowNames If \code{TRUE}, first column of data will be used as row names.
#' @param detectDates If \code{TRUE}, attempt to recognise dates and perform conversion.
#' @param cols A numeric vector specifying which columns in the Excel file to read. 
#' If NULL, all columns are read.
#' @param rows A numeric vector specifying which rows in the Excel file to read. 
#' If NULL, all rows are read.
#' @param check.names logical. If TRUE then the names of the variables in the data frame 
#' are checked to ensure that they are syntactically valid variable names
#' @param namedRegion A named region in the Workbook. If not NULL startRow, rows and cols paramters are ignored.
#' @seealso \code{\link{getNamedRegions}}
#' @details Formulae written using writeFormula to a Workbook object will not get picked up by read.xlsx().
#' This is because only the formula is written and left to be evaluated when the file is opened in Excel.
#' Opening, saving and closing the file with Excel will resolve this.
#' @author Alexander Walker
#' @return data.frame
#' @export
#' @examples
#' xlsxFile <- system.file("readTest.xlsx", package = "openxlsx")
#' df1 <- read.xlsx(xlsxFile = xlsxFile, sheet = 1, skipEmptyRows = FALSE)
#' sapply(df1, class)
#' 
#' df2 <- read.xlsx(xlsxFile = xlsxFile, sheet = 3, skipEmptyRows = TRUE)
#' df2$Date <- convertToDate(df2$Date)
#' sapply(df2, class)
#' head(df2)
#' 
#' df2 <- read.xlsx(xlsxFile = xlsxFile, sheet = 3, skipEmptyRows = TRUE,
#'                    detectDates = TRUE)
#' sapply(df2, class)
#' head(df2)
#' 
#' #wb <- loadWorkbook(system.file("readTest.xlsx", package = "openxlsx"))
#' #df3 <- read.xlsx(wb, sheet = 2, skipEmptyRows = FALSE, colNames = TRUE)
#' #df4 <- read.xlsx(xlsxFile, sheet = 2, skipEmptyRows = FALSE, colNames = TRUE)
#' #all.equal(df3, df4)
#' 
#' #wb <- loadWorkbook(system.file("readTest.xlsx", package = "openxlsx"))
#' #df3 <- read.xlsx(wb, sheet = 2, skipEmptyRows = FALSE,
#' # cols = c(1, 4), rows = c(1, 3, 4))
#' 
#' @export
read.xlsx <- function(xlsxFile,
                      sheet = 1,
                      startRow = 1,
                      colNames = TRUE, 
                      rowNames = FALSE,
                      detectDates = FALSE, 
                      skipEmptyRows = TRUE, 
                      rows = NULL,
                      cols = NULL,
                      check.names = FALSE,
                      namedRegion = NULL){
  
  UseMethod("read.xlsx", xlsxFile) 
  
}

#' @export
read.xlsx.default <- function(xlsxFile,
                              sheet = 1,
                              startRow = 1,
                              colNames = TRUE, 
                              rowNames = FALSE,
                              detectDates = FALSE, 
                              skipEmptyRows = TRUE, 
                              rows = NULL,
                              cols = NULL,
                              check.names = FALSE,
                              namedRegion = NULL){
  
  
  ## Validate inputs and get files
  if(!file.exists(xlsxFile))
    stop("Excel file does not exist.")
  
  if(grepl("\\.xls$|\\.xlm$", xlsxFile))
    stop("openxlsx can not read .xls or .xlm files!")
  
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
  
  if(length(sheet) > 1)
    stop("sheet must be of length 1.")
  
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
  xmlDir <- file.path(tempdir(), "_excelXMLRead")
  xmlFiles <- unzip(xlsxFile, exdir = xmlDir)
  
  on.exit(unlink(xmlDir, recursive = TRUE), add = TRUE)
  
  sharedStringsFile <- xmlFiles[grepl("sharedStrings.xml$", xmlFiles, perl = TRUE)]
  workbook          <- xmlFiles[grepl("workbook.xml$", xmlFiles, perl = TRUE)]
  workbookRelsXML   <- xmlFiles[grepl("workbook.xml.rels$", xmlFiles, perl = TRUE)]
  
  ## get workbook names
  workbookRelsXML <- paste(readLines(workbookRelsXML, warn = FALSE, encoding = "UTF-8"), collapse = "")
  workbookRelsXML <- .Call("openxlsx_getChildlessNode", workbookRelsXML, "<Relationship ", PACKAGE="openxlsx")
  
  
  workbook <- unlist(readLines(workbook, warn = FALSE, encoding = "UTF-8"))
  sheets <- unlist(regmatches(workbook, gregexpr("<sheet .*/sheets>", workbook, perl = TRUE)))
  
  ## make sure sheetId is 1 based
  sheetrId <- unlist(getRId(sheets))
  sheetNames <- unlist(regmatches(sheets, gregexpr('(?<=name=")[^"]+', sheets, perl = TRUE)))
  
  nSheets <- length(sheetrId)
  if(nSheets == 0)
    stop("Workbook has no worksheets")
  
  
  ## Named region logic
  if(!is.null(namedRegion)){
    
    dn <- .Call("openxlsx_getNodes", removeHeadTag(workbook), "<definedNames>", PACKAGE = "openxlsx")
    dn <- unlist(regmatches(dn, gregexpr("<definedName [^<]*", dn, perl = TRUE)))

    if(length(dn) == 0){
      warning("Workbook has no named regions.")
      return(NULL)
    }
    
    dn_names <- replaceXMLEntities(regmatches(dn, regexpr('(?<=name=")[^"]+', dn, perl = TRUE)))
    
    ind <- dn_names == namedRegion
    if(!any(ind))
      stop(sprintf("Region '%s' not found!", region))
    
    ## pull out first node value
    dn <- dn[ind] 
    region <- regmatches(dn, regexpr('(?<=>)[^\\<]+', dn, perl = TRUE))
    sheet <- sheetNames[sapply(sheetNames, function(x) grepl(x, dn))]
    
    region <- gsub(sheet, "", region, fixed = TRUE)
    region <- gsub("[^A-Z0-9:]", "", gsub(sheet, "", region, fixed = TRUE))
    
    cols <- unlist(lapply(strsplit(region, split = ":", fixed = TRUE), convertFromExcelRef))
    rows <- unlist(lapply(strsplit(region, split = ":", fixed = TRUE), function(x) as.integer(gsub("[A-Z]", "", x))))
    
    cols <- seq(from = cols[1], to = cols[2], by = 1)
    rows <- seq(from = rows[1], to = rows[2], by = 1)
    startRow <- 1
    
  }
  
  
  
  
  
  ## get the file_name for each sheetrId
  file_name <- sapply(sheetrId, function(rId){
    txt <- workbookRelsXML[grepl(sprintf('Id="%s"', rId), workbookRelsXML, fixed = TRUE)]
    regmatches(txt, regexpr('(?<=Target=").+xml', txt, perl = TRUE))
  })
  
  
  ## get the correct sheets
  if("character" %in% class(sheet)){
    sheetInd <- which(sheetNames == sheet)
    if(length(sheetInd) == 0)
      stop(sprintf('Cannot find sheet named "%s"', sheet))
    sheet <- file_name[sheetInd]
  }else{
    if(nSheets < sheet)
      stop(sprintf("sheet %s does not exist.", sheet))
    sheet <- file_name[sheet]
  }
  
  ## get file
  worksheet <- xmlFiles[grepl(sheet, xmlFiles, ignore.case = TRUE)]
  
  ## read in sharedStrings
  if(length(sharedStringsFile) > 0){
    sharedStrings <- getSharedStringsFromFile(sharedStringsFile = sharedStringsFile, isFile = TRUE)
  }else{
    sharedStrings <- ""
  }
  
  
  if("character" %in% class(startRow)){
    startRowStr <- startRow
    startRow <- 1
  }else{
    startRowStr <- NULL
  }
  
  ## single function get all r, s (if detect dates is TRUE), t, v
  cell_info <- .Call("openxlsx_getCellInfo",
                     xmlFile = worksheet,
                     sharedStrings = sharedStrings,
                     skipEmptyRows = skipEmptyRows,
                     startRow = startRow,
                     rows = rows,
                     getDates = detectDates,
                     PACKAGE = "openxlsx")
  
  
  nRows <- cell_info$nRows
  r <- cell_info$r
  
  if(nRows == 0 | length(r) == 0){
    warning("No data found on worksheet.", call. = FALSE)
    return(NULL)
  }
  
  v <- cell_info$v
  Encoding(v) <- "UTF-8"
  string_refs <- cell_info$string_refs
  
  
  ## subset if specified a row/col index
  if(!is.null(cols)){
    
    flag <- which(convertFromExcelRef(r) %in% cols)
    r <- r[flag]
    if(length(r) == 0){
      warning("No data found on worksheet.", call. = FALSE)
      return(NULL)
    }
    v <- v[flag]
    
    if(detectDates)
      cell_info$s <- cell_info$s[flag]
    
    flag <- which(convertFromExcelRef(string_refs) %in% cols)
    string_refs <- string_refs[flag]
    
    nRows <- .Call("openxlsx_calcNRows", r, skipEmptyRows, PACKAGE = "openxlsx")
  }
  
  
  if(!is.null(startRowStr)){
    ind <- which(grepl(startRowStr, v, ignore.case = TRUE))
    if(length(ind) > 0){
      startRow <- as.numeric(gsub("[A-Z]", "", r[ind[[1]]]))
      toKeep <- grep(sprintf("[A-Z]%s$", startRow), r)[[1]]
      if(toKeep > 1){
        toRemove <- 1:(toKeep-1)
        string_refs <- string_refs[!string_refs %in% r[toRemove]]
        v <- v[-toRemove]
        r <- r[-toRemove]
        nRows <- .Call("openxlsx_calcNRows", r, skipEmptyRows, PACKAGE = "openxlsx")
      }
    }
  }
  
  
  ## Determine date cells (if required)
  origin <- 25569L
  if(detectDates){
    
    ## get date origin
    if(grepl('date1904="1"|date1904="true"', paste(workbook, collapse = ""), ignore.case = TRUE))
      origin <- 24107L
    
    stylesXML <- xmlFiles[grepl("styles.xml", xmlFiles)]
    styles <- readLines(stylesXML, warn = FALSE)
    styles <- removeHeadTag(styles)
    
    ## Number formats
    numFmts <- .Call("openxlsx_getChildlessNode", styles, "<numFmt ", PACKAGE = "openxlsx")
    
    dateIds <- NULL
    if(length(numFmts) > 0){
      
      numFmtsIds <- sapply(numFmts, function(x) .Call("openxlsx_getAttr", x, 'numFmtId="', PACKAGE = "openxlsx"), USE.NAMES = FALSE)
      formatCodes <- sapply(numFmts, function(x) .Call("openxlsx_getAttr", x, 'formatCode="', PACKAGE = "openxlsx"), USE.NAMES = FALSE)
      formatCodes <- gsub(".*(?<=\\])|@", "", formatCodes, perl = TRUE)
      
      ## this regex defines what "looks" like a date
      dateIds <- numFmtsIds[!grepl("[^mdyhsapAMP[:punct:] ]", formatCodes) & nchar(formatCodes > 3)]
      
    }     
    
    dateIds <- c(dateIds, 14)
    
    ## which styles are using these dateIds
    cellXfs <- .Call("openxlsx_getNodes", styles, "<cellXfs", PACKAGE = "openxlsx") 
    xf <- .Call("openxlsx_getChildlessNode", cellXfs, "<xf ", PACKAGE = "openxlsx")
    lookingFor <- paste(sprintf('numFmtId="%s"', dateIds), collapse = "|")
    dateStyleIds <- which(sapply(xf, function(x) grepl(lookingFor, x), USE.NAMES = FALSE)) - 1L
    
    isDate <- (cell_info$s %in% dateStyleIds) & !(r %in% string_refs)
    
    # check numbers are also integers
    isNotInt <- suppressWarnings(as.numeric(v[isDate]))
    isNotInt <- (isNotInt %% 1L != 0) & !is.na(isNotInt)
    isDate[isNotInt] <- FALSE
    
    
    ## perform int to date to character convertsion (way too slow)
    v[isDate] <- format(as.Date(as.integer(v[isDate]) - origin, origin = "1970-01-01"), "%Y-%m-%d")
    
  }else{
    isDate <- as.logical(NA)
  }

  ## Build data.frame
  m <- .Call("openxlsx_readWorkbook", v, r, string_refs, isDate,  nRows, colNames, skipEmptyRows, clean_names, PACKAGE = "openxlsx")
  
  if(colNames && check.names)
    colnames(m) <- make.names(colnames(m), unique = TRUE)
  
  
  if(rowNames){
    rownames(m) <- m[[1]]
    m[[1]] <- NULL
  }
  
  return(m)
  
}







#' @export
read.xlsx.Workbook <- function(xlsxFile,
                               sheet = 1,
                               startRow = 1,
                               colNames = TRUE, 
                               rowNames = FALSE,
                               detectDates = FALSE, 
                               skipEmptyRows = TRUE, 
                               rows = NULL,
                               cols = NULL,
                               check.names = FALSE,
                               namedRegion = NULL){
  
  
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
  
  if(length(sheet) != 1)
    stop("sheet must be of length 1.")
  
  ## Named region logic
  if(!is.null(namedRegion)){
    
    dn <- xlsxFile$workbook$definedNames
    if(length(dn) == 0){
      warning("Workbook has no named regions.")
      return(NULL)
    }
    
    dn_names <- replaceXMLEntities(regmatches(dn, regexpr('(?<=name=")[^"]+', dn, perl = TRUE)))
    
    ind <- dn_names == namedRegion
    if(!any(ind))
      stop(sprintf("Region '%s' not found!", region))
    
    ## pull out first node value
    dn <- dn[ind] 
    region <- regmatches(dn, regexpr('(?<=>)[^\\<]+', dn, perl = TRUE))
    sheet <- names(xlsxFile)[sapply(names(xlsxFile), function(x) grepl(x, dn))]
    
    region <- gsub("[^A-Z0-9:]", "", gsub(sheet, "", region, fixed = TRUE))
    
    cols <- unlist(lapply(strsplit(region, split = ":", fixed = TRUE), convertFromExcelRef))
    rows <- unlist(lapply(strsplit(region, split = ":", fixed = TRUE), function(x) as.integer(gsub("[A-Z]", "", x))))
    
    cols <- seq(from = cols[1], to = cols[2], by = 1)
    rows <- seq(from = rows[1], to = rows[2], by = 1)
    startRow <- 1
    
  }
  
  
  if(is.null(rows)){
    rows <- NA
  }else if(length(rows) > 1){
    rows <- as.integer(sort(rows))
  }else{
    stop("rows must be an integer vector else NULL or NA")
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
  sheetNames <- names(xlsxFile$worksheets)
  
  if("character" %in% class(sheet)){
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
  if(length(sharedStrings) > 0)
    sharedStrings <- getSharedStringsFromFile(sharedStringsFile = sharedStrings, isFile = FALSE)
  
  
  ## read in worksheet and get cells with a value node, skip emptyStrs cells
  sheetData <- xlsxFile$sheetData[[sheet]]
  if(!is.na(rows[1]))
    sheetData <- sheetData[as.numeric(names(sheetData)) %in% rows]
  
  if(!is.null(cols)){
    r <- unlist(lapply(sheetData, "[[", 1))
    sheetData <- sheetData[convertFromExcelRef(r) %in% cols]
  }
  
  
  r <- unlist(lapply(sheetData, "[[", 1), use.names = FALSE)
  v <- unlist(lapply(sheetData, "[[", 3), use.names = FALSE)
  t <- unlist(lapply(sheetData, "[[", 2), use.names = FALSE)
  
  if(startRow > 1){
    
    rows <- as.numeric(gsub("[A-Z]", "", r))
    r <- r[rows >= startRow]
    v <- v[rows >= startRow]
    t <- t[rows >= startRow]
    rm(rows)
  }
  
  if(length(r) == 0){
    warning("No data found on worksheet.", call. = FALSE)
    return(NULL)
  }
  
  ## remove NA t and v values
  toRemove <- which(is.na(v) & is.na(t))
  if(length(toRemove) > 0){
    r <- r[-toRemove]
    t <- t[-toRemove]
    v <- v[-toRemove]
  }
  
  if(is.null(r)){
    warning("No data found on worksheet.", call. = FALSE)
    return(NULL)
  }else{
    nRows <- .Call("openxlsx_calcNRows", r, skipEmptyRows, PACKAGE = "openxlsx")
  }
  
  ## get references for string cells
  string_refs <- r[t == "b" | t == "s"]
  if(length(string_refs) == 0)
    string_refs <- -1L
  
  
  ## get Refs for boolean 
  bool_refs <- r[t == "b"]
  if(length(bool_refs) == 0)
    bool_refs <- -1L
  
  if(bool_refs[[1]] != -1L){
    
    fInd <- which(sharedStrings == "FALSE") - 1L
    if(length(fInd) == 0){
      fInd <- length(sharedStrings) 
      sharedStrings <- c(sharedStrings, "FALSE")
    }
    
    tInd <- which(sharedStrings == "TRUE") - 1L
    if(length(tInd) == 0){
      tInd <- length(sharedStrings) 
      sharedStrings <- c(sharedStrings, "TRUE")
    }
    
    boolInds <- match(bool_refs, r)
    logicalVals <- v[boolInds]
    logicalVals[logicalVals == "0"] <- fInd[[1]]
    logicalVals[logicalVals == "1"] <- tInd[[1]]
    v[boolInds] <- logicalVals
    
    rm(bool_refs)
  }
  
  
  ## If any t="str" exist, add v to sharedStrings and replace v with newSharedStringsInd
  wsStrInds <- which(t == "str")
  if(length(wsStrInds) > 0){
    
    uStrs <- unique(v[wsStrInds])
    
    uStrs[uStrs == "#N/A"] <- NA
    
    ## Match references of "str" cells to r
    strInds <- na.omit(match(r[wsStrInds], r))
    newSharedStringInds <- length(sharedStrings):(length(sharedStrings) + length(uStrs) - 1L) 
    
    ## replace strings in v with reference to sharedStrings, (now can convert v to numeric)
    v[strInds] <- newSharedStringInds[match(v[wsStrInds], uStrs)]
    
    ## append new strings to sharedStrings
    sharedStrings <- c(sharedStrings, uStrs)
    if(string_refs[[1]] == -1L){
      string_refs <- r[wsStrInds]
    }else{
      string_refs <- c(string_refs, r[wsStrInds])
      string_refs <- string_refs[order(as.numeric(gsub("[A-Z]", "", string_refs)), nchar(string_refs))]   
    }
    
  }
  
  ##Set error cells to NA
  wsStrInds <- which(t == "e")
  if(length(wsStrInds) > 0){
    
    ## Match references of "e" cells to r and set v values to NA
    inds <- na.omit(match(r[wsStrInds], r))
    #     
    #     inds1 <- which(v[inds] == "#NUM!")
    #     if(length(inds1) > 0){
    #       v[inds[inds1]] <- NaN
    #       inds <- inds[-inds1]
    #     }
    #     
    #     if(length(inds) > 0)
    v[inds] <- NA
    
  }
  
  
  ## Now safe to convert v to numeric
  vn <- as.numeric(v)
  
  ## Using -1 as a flag for no strings
  if(length(sharedStrings) == 0 | string_refs[1] == -1L){
    string_refs <- as.character(NA)
  }else{
    
    ## set encoding of sharedStrings
    Encoding(sharedStrings) <- "UTF-8"
    
    ## Now replace values in v with string values
    stringInds <- match(string_refs, r)
    v[stringInds] <- sharedStrings[vn[stringInds] + 1L]
    
  }
  
  
  ## date detection
  origin <- 25569L
  isDate <- rep.int(FALSE, times = length(r))
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
    formatCodes <- unlist(lapply(sO, function(x) {
      
      fc <- x[["style"]][["numFmt"]]$formatCode
      if(is.null(fc))
        fc <- x[["style"]][["numFmt"]]$numFmtId
      fc
    }))
    
    
    dateIds <- NULL
    if(length(formatCodes) > 0){
      
      ## this regex defines what "looks" like a date
      formatCodes <- gsub(".*(?<=\\])|@", "", formatCodes, perl = TRUE)
      sO <- sO[(!grepl("[^mdyhsapAMP[:punct:] ]", formatCodes) & nchar(formatCodes > 3)) | formatCodes == 14]
      
    }     
    
    if(length(sO) > 0){
      
      rows <- unlist(lapply(sO, "[[", "rows"))
      cols <- unlist(lapply(sO, "[[", "cols"))    
      refs <- paste0(convert2ExcelRef(cols = cols, LETTERS), rows)
      
      isDate <- r %in% refs
      
      ## check numbers are also integers
      isNotInt <- suppressWarnings(as.numeric(v[isDate]))
      isNotInt <- (isNotInt %% 1L != 0) | is.na(isNotInt)
      isDate[isNotInt] <- FALSE
      
      ## perform int to date to character convertsion (way too slow)
      v[isDate] <- format(as.Date(as.integer(v[isDate]) - origin, origin = "1970-01-01"), "%Y-%m-%d")
      
    }
  }
  

  ## Build data.frame
  m <- .Call("openxlsx_readWorkbook", v, r, string_refs, isDate,  nRows, colNames, skipEmptyRows, clean_names, PACKAGE = "openxlsx")
  
  if(colNames && check.names)
    colnames(m) <- make.names(colnames(m), unique = TRUE)
  
  if(rowNames){
    rownames(m) <- m[[1]]
    m[[1]] <- NULL
  }
  
  return(m)
  
}







#' @name readWorkbook
#' @title Read data from a worksheet into a data.frame
#' @param xlsxFile An xlsx file
#' @param sheet The name or index of the sheet to read data 
#' @param startRow first row to begin looking for data.  Empty rows at the top of a file are always skipped, 
#' regardless of the value of startRow.
#' @param colNames If \code{TRUE}, first row of data will be used as column names. 
#' @param skipEmptyRows If \code{TRUE}, empty rows are skipped else empty rows after the first row containing data 
#' will return a row of NAs
#' @param rowNames If \code{TRUE}, first column of data will be used as row names.
#' @details Creates a data.frame of all data in worksheet.
#' @param detectDates If \code{TRUE}, attempt to recognise dates and perform conversion.
#' @param cols A numeric vector specifying which columns in the Excel file to read. 
#' If NULL, all columns are read.
#' @param rows A numeric vector specifying which rows in the Excel file to read. 
#' If NULL, all rows are read.
#' @param namedRegion A named region in the Workbook. If not NULL startRow, rows and cols paramters are ignored.
#' @param check.names logical. If TRUE then the names of the variables in the data frame 
#' are checked to ensure that they are syntactically valid variable names
#' @author Alexander Walker
#' @return data.frame
#' @seealso \code{\link{getNamedRegions}}
#' @export
#' @seealso \code{\link{read.xlsx}}
#' @export
#' @examples
#' xlsxFile <- system.file("readTest.xlsx", package = "openxlsx")
#' df1 <- readWorkbook(xlsxFile = xlsxFile, sheet = 1)
#' 
#' xlsxFile <- system.file("readTest.xlsx", package = "openxlsx")
#' df1 <- readWorkbook(xlsxFile = xlsxFile, sheet = 1, rows = c(1, 3, 5), cols = 1:3)
readWorkbook <- function(xlsxFile,
                         sheet = 1,
                         startRow = 1,
                         colNames = TRUE, 
                         rowNames = FALSE,
                         detectDates = FALSE, 
                         skipEmptyRows = TRUE, 
                         rows = NULL,
                         cols = NULL,
                         check.names = FALSE,
                         namedRegion = NULL){
  
  read.xlsx(xlsxFile = xlsxFile,
            sheet = sheet,
            startRow = startRow,
            colNames = colNames, 
            rowNames = rowNames,
            detectDates = detectDates, 
            skipEmptyRows = skipEmptyRows, 
            rows = rows,
            cols = cols,
            check.names = check.names,
            namedRegion = namedRegion)
}




