


#' @name read.xlsx
#' @title Read data from a worksheet or Workbook object into a data.frame
#' @param xlsxFile An xlsx file or Workbook object
#' @param sheet The name or index of the sheet to read data from.
#' @param startRow first row to begin looking for data.  Empty rows at the top of a file are always skipped, 
#' regardless of the value of startRow.
#' @param colNames If TRUE, the first row of data will be used as column names. 
#' @param skipEmptyRows If TRUE, empty rows are skipped else empty rows after the first row containing data 
#' will return a row of NAs.
#' @param rowNames If TRUE, first column of data will be used as row names.
#' @param keepNewLine If TRUE, keep new line characters embedded in strings.
#' @param detectDates If TRUE, attempt to recognise dates and perform conversion.
#' @details Creates a data.frame of all the data on a worksheet.
#' @author Alexander Walker
#' @return data.frame
#' @export
#' @examples
#' xlsxFile <- system.file("readTest.xlsx", package = "openxlsx")
#' df1 <- read.xlsx(xlsxFile = xlsxFile, sheet = 1, startRow = 1, skipEmptyRows = FALSE)
#' sapply(df1, class)
#' 
#' df2 <- read.xlsx(xlsxFile = xlsxFile, sheet = 3, startRow = 1, skipEmptyRows = TRUE)
#' df2$Date <- convertToDate(df2$Date)
#' sapply(df2, class)
#' head(df2)
#' 
#' wb <- loadWorkbook(system.file("readTest.xlsx", package = "openxlsx"))
#' df3 <- read.xlsx(wb, sheet = 1, startRow = 1, skipEmptyRows = FALSE, colNames = TRUE)
#' all.equal(df1, df3)
#'
#' @export
read.xlsx <- function(xlsxFile, sheet = 1, startRow = 1, colNames = TRUE, skipEmptyRows = TRUE, rowNames = FALSE, keepNewLine = FALSE, detectDates = FALSE){
  
  UseMethod("read.xlsx", xlsxFile) 
  
}

#' @export
read.xlsx.default <- function(xlsxFile, sheet = 1, startRow = 1, colNames = TRUE, skipEmptyRows = TRUE, rowNames = FALSE, keepNewLine = FALSE, detectDates = FALSE){
  
  if(!file.exists(xlsxFile))
    stop("Excel file does not exist.")
  
  if(grepl("\\.xls$|\\.xlm$", xlsxFile))
    stop("openxlsx can not read .xls or .xlm files!")
  
  if(length(sheet) > 1)
    stop("sheet must be of length 1.")
  
  if(startRow < 1)
    startRow <- 1L
  
  ## create temp dir and unzip
  xmlDir <- paste0(tempdir(), "_excelXMLRead")
  xmlFiles <- unzip(xlsxFile, exdir = xmlDir)
  
  on.exit(unlink(xmlDir, recursive = TRUE), add = TRUE)
  
  worksheets    <- xmlFiles[grepl("/worksheets/sheet[0-9]", xmlFiles, perl = TRUE)]
  sharedStringsFile <- xmlFiles[grepl("sharedStrings.xml$", xmlFiles, perl = TRUE)]
  workbook      <- xmlFiles[grepl("workbook.xml$", xmlFiles, perl = TRUE)]
  
  nSheets <- length(worksheets)
  if(nSheets == 0)
    stop("Workbook has no worksheets")
  
  ## get workbook names
  workbook <- unlist(readLines(workbook, warn = FALSE))
  sheets <- unlist(regmatches(workbook, gregexpr("<sheet .*/sheets>", workbook, perl = TRUE)))
  sheetrId <- as.integer(unlist(regmatches(sheets, gregexpr('(?<=r:id="rId)[0-9]+', sheets, perl = TRUE)))) 
  sheetNames <- unlist(regmatches(sheets, gregexpr('(?<=name=")[^"]+', sheets, perl = TRUE)))
  
  
  
  if("character" %in% class(sheet)){
    sheetInd <- which(sheetNames == sheet)
    if(length(sheetInd) == 0)
      stop(sprintf('Cannot find sheet named "%s"', sheet))
    sheet <- sheetrId[sheetInd]
  }else{
    if(nSheets < sheet)
      stop(sprintf("sheet %s does not exist.", sheet))
    sheet <- sheetrId[sheet]
  }
    
  ## read in sharedStrings
  if(length(sharedStringsFile) > 0){
    
    ## read in, get si tags, get t tag value
    if(keepNewLine){
      ss <- paste(readLines(sharedStringsFile, warn = FALSE), collapse = "\n")
      ss <- gsub("\n<", "", ss)
    }else{
      ss <- .Call("openxlsx_cppReadFile", sharedStringsFile, PACKAGE = "openxlsx")
    }
    
    ## pull out all string nodes
    sharedStrings <- .Call("openxlsx_getNodes", ss, "<si>", PACKAGE = "openxlsx")  
  
    
    ## Need to remove any inline styling
    formattingFlag <- grepl("<rPr>", ss)
    if(formattingFlag){
      sharedStrings <- .Call("openxlsx_getSharedStrings2", sharedStrings, PACKAGE = 'openxlsx') ## Where there is inline formatting
    }else{
      sharedStrings <- .Call("openxlsx_getSharedStrings", sharedStrings, PACKAGE = 'openxlsx') ## No inline formatting
    }
    
    emptyStrs <- c(attr(sharedStrings, "empty"), which(sharedStrings == "") - 1L)
        
    Encoding(sharedStrings) <- "UTF-8"
    z <- tolower(sharedStrings)
    sharedStrings[z == "true"] <- "TRUE"
    sharedStrings[z == "false"] <- "FALSE"
    rm(z)
    
    ###  invalid xml character replacements
    ## XML replacements
    sharedStrings <- gsub("&amp;", '&', sharedStrings)
    sharedStrings <- gsub("&lt;", '<', sharedStrings)
    sharedStrings <- gsub("&gt;", '>', sharedStrings)
    sharedStrings <- gsub("&quot;", '"', sharedStrings)
    sharedStrings <- gsub("&apos;", "'", sharedStrings)
    
  }else{
    sharedStrings <- NULL
    emptyStrs <- NULL
  }
  
  if(length(emptyStrs) == 0)
    emptyStrs <- ""
  
  ## read in worksheet and get cells with a value node, skip emptyStrs cells
  worksheets <- worksheets[order(nchar(worksheets), worksheets)]
  ws <- .Call("openxlsx_getCellsWithChildren", worksheets[[sheet]], sprintf("<v>%s</v>", emptyStrs), PACKAGE = "openxlsx")
  Encoding(ws) <- "UTF-8"
  
  r_v <- .Call("openxlsx_getRefsVals", ws, startRow, PACKAGE = "openxlsx")
  r <- r_v[[1]]
  v <- r_v[[2]]
  
  nRows <- .Call("openxlsx_calcNRows", r, skipEmptyRows, PACKAGE = "openxlsx")
  if(nRows == 0 | length(r) == 0){
    warning("No data found on worksheet.")
    return(NULL)
  }
  
  ## get references for string cells
  tR <- .Call("openxlsx_getRefs", ws[which(grepl('t="s"|t="b"', ws, perl = TRUE))], startRow, PACKAGE = "openxlsx")
  if(length(tR) == 0)
    tR <- -1L

  ## get Refs for boolean 
  tB <- .Call("openxlsx_getRefs", ws[which(grepl('t="b"', ws, perl = TRUE))], startRow, PACKAGE = "openxlsx")
  if(length(tB) > 0 & tB[[1]] != -1L){
    
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
    
    boolInds <- match(tB, r)
    logicalVals <- v[boolInds]
    logicalVals[logicalVals == "0"] <- fInd[[1]]
    logicalVals[logicalVals == "1"] <- tInd[[1]]
    v[boolInds] <- logicalVals
    
  }
  
  if(tR[[1]] == -1L){
    stringInds <- -1L
  }else{
    stringInds <- match(tR, r)
    stringInds <- stringInds[!is.na(stringInds)]
  }
  
  ## If any t="str" exist, add v to sharedStrings and replace with newSharedStringsInd
  wsStrInds <- which(grepl('t="str"|t="e"', ws, perl = TRUE))
  if(length(wsStrInds) > 0){
    
    strRV <- .Call("openxlsx_getRefsVals",  ws[wsStrInds], startRow, PACKAGE = "openxlsx")
    uStrs <- unique(strRV[[2]])
    uStrs[uStrs == "#N/A"] <- NA
    
    ## Match references of "str" cells to r, append these to stringInds
    strInds <- na.omit(match(strRV[[1]], r))
    newSharedStringInds <- length(sharedStrings):(length(sharedStrings) + length(uStrs) - 1L) 
    
    ## replace strings in v with reference to sharedStrings, (now can convert v to numeric)
    v[strInds] <- newSharedStringInds[match(strRV[[2]], uStrs)]
    
    ## append new strings to sharedStrings
    sharedStrings <- c(sharedStrings, uStrs)
    if(tR[[1]] == -1L){
      stringInds <- strInds
      tR <- strRV[[1]]
    }else{
      stringInds <- c(stringInds, strInds)
      tR <- c(tR, strRV[[1]])
      tR <- tR[order(as.numeric(gsub("[A-Z]", "", tR)), nchar(tR))]   
    }
    
  }
  
  ## Now safe to convert v to numeric
  vn <- as.numeric(v)
  
  ## Using -1 as a flag for no strings
  if(length(sharedStrings) == 0 | stringInds[[1]] == -1L){
    stringInds <- -1L
    tR <- as.character(NA)
  }else{
    
    ## set encoding of sharedStrings
    Encoding(sharedStrings) <- "UTF-8"
    
    ## Now replace values in v with string values
    v[stringInds] <- sharedStrings[vn[stringInds] + 1L]
    
    ## decrement stringInds as sharedString inds are zero based
    stringInds = stringInds - 1L;  
    
  }
  
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
    
    
    s <- .Call("openxlsx_getCellStylesPossiblyMissing", ws, package = "openxlsx")
    
    styleInds <- s %in% dateStyleIds
    
    ## check numbers are also integers
    isNotInt <- suppressWarnings(as.numeric(v[styleInds]))
    isNotInt <- (isNotInt %% 1L != 0) | is.na(isNotInt)
    styleInds[isNotInt] <- FALSE
    
    ## now replace string with something to look for
    v[styleInds] <- "openxlsxdt"
    
    stringInds <- c(stringInds, which(styleInds) - 1L)
    tR <- c(tR, r[styleInds]) 
    
  }
  
  
  ## Build data.frame
  m <- .Call("openxlsx_readWorkbook", v, vn, stringInds, r, tR,  as.integer(nRows), colNames, skipEmptyRows, origin, PACKAGE = "openxlsx")
  
  if(length(colnames(m)) > 0){
    colnames(m) <- gsub("^[[:space:]]+|[[:space:]]+$", "", colnames(m))
    colnames(m) <- gsub("[[:space:]]+", ".", colnames(m))
  }
  
  if(rowNames){
    rownames(m) <- m[[1]]
    m[[1]] <- NULL
  }
  
  
  return(m)
  
}


#' @export
read.xlsx.Workbook <- function(xlsxFile, sheet = 1, startRow = 1, colNames = TRUE, skipEmptyRows = TRUE, rowNames = FALSE, keepNewLine = FALSE, detectDates = FALSE){
  
  if(length(sheet) != 1)
    stop("sheet must be of length 1.")
  
  if(startRow < 1)
    startRow <- 1L
  
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
  sharedStrings <- unlist(xlsxFile$sharedStrings)
  
  if(length(sharedStrings) > 0){
    
    ## Need to remove any inline styling
    formattingFlag <- any(grepl("<rPr>", sharedStrings))
    if(formattingFlag){
      sharedStrings <- .Call("openxlsx_getSharedStrings2", sharedStrings, PACKAGE = 'openxlsx') ## Where there is inline formatting
    }else{
      sharedStrings <- .Call("openxlsx_getSharedStrings", sharedStrings, PACKAGE = 'openxlsx') ## No inline formatting
    }
    
    emptyStrs <- attr(sharedStrings, "empty")
    
    z <- tolower(sharedStrings)
    sharedStrings[z == "true"] <- "TRUE"
    sharedStrings[z == "false"] <- "FALSE"
    rm(z)
    
    ###  invalid xml character replacements
    ## XML replacements
    sharedStrings <- gsub("&amp;", '&', sharedStrings)
    sharedStrings <- gsub("&lt;", '<', sharedStrings)
    sharedStrings <- gsub("&gt;", '>', sharedStrings)
    sharedStrings <- gsub("&quot;", '"', sharedStrings)
    sharedStrings <- gsub("&apos;", "'", sharedStrings)
    
  }else{
    sharedStrings <- NULL
    emptyStrs <- NULL
  }
  
  
  if(length(emptyStrs) == 0)
    emptyStrs <- ""
  
  ## read in worksheet and get cells with a value node, skip emptyStrs cells
  r <- unlist(lapply(xlsxFile$sheetData[[sheet]], "[[", 1))
  v <- unlist(lapply(xlsxFile$sheetData[[sheet]], "[[", 3))
  t <- unlist(lapply(xlsxFile$sheetData[[sheet]], "[[", 2))
  
  if(startRow > 1){
        
    rows <- as.numeric(gsub("[A-Z]", "", r))
    r <- r[rows >= startRow]
    v <- v[rows >= startRow]
    t <- t[rows >= startRow]
    
  }
  
  if(length(r) == 0){
    warning("No data found on worksheet.")
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
    warning("No data found on worksheet.")
    return(NULL)
  }else{
    nRows <- .Call("openxlsx_calcNRows", r, skipEmptyRows, PACKAGE = "openxlsx")
  }
    
  ## get references for string cells
  tR <- r[t == "b" | t == "s"]
  if(length(tR) == 0)
    tR <- -1L
  
  ## get Refs for boolean 
  tB <- r[t == "b"]
  if(length(tB) == 0)
    tB <- -1L
  
  if(length(tB) > 0 & tB[[1]] != -1L){
    
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
    
    boolInds <- match(tB, r)
    logicalVals <- v[boolInds]
    logicalVals[logicalVals == "0"] <- fInd[[1]]
    logicalVals[logicalVals == "1"] <- tInd[[1]]
    v[boolInds] <- logicalVals
    
  }
  
  if(tR[[1]] == -1L){
    stringInds <- -1L
  }else{
    stringInds <- match(tR, r)
    stringInds <- stringInds[!is.na(stringInds)]
  }
  
  ## If any t="str" exist, add v to sharedStrings and replace v with newSharedStringsInd
  wsStrInds <- which(t == "str"| t == "e")
  if(length(wsStrInds) > 0){
      
    uStrs <- unique(v[wsStrInds])
    uStrs[uStrs == "#N/A"] <- NA
    
    ## Match references of "str" cells to r, append these to stringInds
    strInds <- na.omit(match(r[wsStrInds], r))
    newSharedStringInds <- length(sharedStrings):(length(sharedStrings) + length(uStrs) - 1L) 
    
    ## replace strings in v with reference to sharedStrings, (now can convert v to numeric)
    v[strInds] <- newSharedStringInds[match(v[wsStrInds], uStrs)]
    
    ## append new strings to sharedStrings
    sharedStrings <- c(sharedStrings, uStrs)
    if(tR[[1]] == -1L){
      stringInds <- strInds
      tR <- r[wsStrInds]
    }else{
      stringInds <- c(stringInds, strInds)
      tR <- c(tR, r[wsStrInds])
      tR <- tR[order(as.numeric(gsub("[A-Z]", "", tR)), nchar(tR))]   
    }
    
  }
  
  ## Now safe to convert v to numeric
  vn <- as.numeric(v)
  
  ## Using -1 as a flag for no strings
  if(length(sharedStrings) == 0 | stringInds[[1]] == -1L){
    stringInds <- -1L
    tR <- as.character(NA)
  }else{
    
    ## set encoding of sharedStrings
    Encoding(sharedStrings) <- "UTF-8"
    
    ## Now replace values in v with string values
    v[stringInds] <- sharedStrings[vn[stringInds] + 1L]
    
    ## decrement stringInds as sharedString inds are zero based
    stringInds = stringInds - 1L;  
    
  }
  
  origin <- 25569L
  if(detectDates){

    ## get date origin
    if(length(xlsxFile$workbook$workbookPr) > 0){
      if(grepl('date1904="1"|date1904="true"', xlsxFile$workbook$workbookPr, ignore.case = TRUE))
        origin <- 24107L
    }
     
    sO <- xlsxFile$styleObjects
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
  
      styleInds <- match(refs, r)
      
      ## check numbers are also integers
      isNotInt <- suppressWarnings(as.numeric(v[styleInds]))
      isNotInt <- (isNotInt %% 1L != 0) | is.na(isNotInt)
      styleInds[isNotInt] <- FALSE
      
      ## now replace string with something to look for
      v[styleInds] <- "openxlsxdt"
      
      stringInds <- c(stringInds, styleInds - 1L)
      tR <- c(tR, r[styleInds]) 
      
    }
  }
  
  
  ## Build data.frame
  m <- .Call("openxlsx_readWorkbook", v, vn, stringInds, r, tR,  as.integer(nRows), colNames, skipEmptyRows, origin, PACKAGE = "openxlsx")
  
  if(length(colnames(m)) > 0){
    colnames(m) <- gsub("^[[:space:]]+|[[:space:]]+$", "", colnames(m))
    colnames(m) <- gsub("[[:space:]]+", ".", colnames(m))
  }
  
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
#' @param startRow first row to begin looking for data.  Empty rows before any data is found are skipped.
#' regardless of the value of startRow.
#' @param colNames If TRUE, first row of data will be used as column names. 
#' @param skipEmptyRows If TRUE, empty rows are skipped else empty rows after the first row containing data 
#' will return a row of NAs
#' @param rowNames If TRUE, first column of data will be used as row names.
#' @details Creates a data.frame of all data in worksheet.
#' @param keepNewLine If TRUE, keep new line characters embedded in strings.
#' @param detectDates If TRUE, attempt to recognise dates and perform conversion.
#' @author Alexander Walker
#' @return data.frame
#' @export
#' @seealso \code{\link{read.xlsx}}
#' @export
#' @examples
#' xlsxFile <- system.file("readTest.xlsx", package = "openxlsx")
#' df1 <- readWorkbook(xlsxFile = xlsxFile, sheet = 1)
readWorkbook <- function(xlsxFile, sheet = 1, startRow = 1, colNames = TRUE, skipEmptyRows = TRUE, rowNames = FALSE, keepNewLine = FALSE, detectDates = FALSE){
  read.xlsx(xlsxFile, sheet = sheet, startRow = startRow, colNames = colNames, skipEmptyRows = skipEmptyRows, rowNames = rowNames, keepNewLine = keepNewLine, detectDates = detectDates)
}
