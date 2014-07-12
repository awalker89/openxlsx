
#' @name read.xlsx
#' @title Read data from a worksheet into a data.frame
#' @param xlsxFile An xlsx file
#' @param sheet The name or index of the sheet to read data 
#' @param startRow first row to begin looking for data.  Empty rows before any data is found are skipped.
#' regardless of the value of startRow.
#' @param colNames If TRUE, first row of data will be used as column names. 
#' @param skipEmptyRows If TRUE, empty rows are skipped else empty rows after the first row containing data 
#' will return a row of NAs
#' @details Creates a data.frame of all data in worksheet.
#' @author Alexander Walker
#' @return data.frame
#' @seealso \code{\link{readWorkbook}}
#' @export
#' @examples
#' xlsxFile <- system.file("readTest.xlsx", package = "openxlsx")
#' df1 <- read.xlsx(xlsxFile = xlsxFile, sheet = 1, startRow=1, skipEmptyRows=FALSE, colNames=TRUE)
#' sapply(df1, class)
#' 
#' df2 <- read.xlsx(xlsxFile = xlsxFile, sheet = 3, startRow=1, skipEmptyRows=TRUE, colNames=TRUE)
#' df2$Date <- convertToDate(df2$Date)
#' sapply(df2, class)
#' head(df2)
#' 
#' df3 <- read.xlsx(xlsxFile = xlsxFile, sheet = 4, startRow=1, skipEmptyRows=TRUE, colNames=TRUE)
#' df3$Symbol
#' 
#' @export
read.xlsx <- function(xlsxFile, sheet = 1, startRow = 1, colNames = TRUE, skipEmptyRows = TRUE){

  if(!file.exists(xlsxFile))
    stop("Excel file does not exist.")
  
  if(!grepl("xlsx$", xlsxFile))
    stop("File must have extension .xlsx!")
  
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
  sheetNames <- unlist(regmatches(workbook, gregexpr('(?<=<sheet name=")[^"]+', workbook, perl = TRUE)))
  
  if("character" %in% class(sheet)){
    sheetInd <- which(sheetNames == sheet)
    if(length(sheetInd) == 0)
      stop(sprintf('Cannot find sheet named "%s"', sheet))
  }else{
    sheetInd <- sheet
    if(nSheets < sheetInd)
      stop(sprintf("sheet %s does not exist.", sheet))
  }
  
  
  ## read in sharedStrings
  if(length(sharedStringsFile) > 0){
    
    ## read in, get si tags, get t tag value
    ss <- .Call("openxlsx_cppReadFile", sharedStringsFile, PACKAGE = "openxlsx")
    sharedStrings <- .Call("openxlsx_getNodes", ss, "<si>", PACKAGE = "openxlsx")
    sharedStrings <- .Call("openxlsx_getSharedStrings", sharedStrings, PACKAGE = 'openxlsx')
    emptyStrs <- attr(sharedStrings, "empty")
    
    sharedStrings[grepl("true", sharedStrings, ignore.case = TRUE)] <- "TRUE"
    sharedStrings[grepl("false", sharedStrings, ignore.case = TRUE)] <- "FALSE"
    
    ###  invalid xml character replacements
    ## XML replacements
    sharedStrings <- gsub("&amp;", '&', sharedStrings)
    sharedStrings <- gsub("&lt;", '<', sharedStrings)
    sharedStrings <- gsub("&gt;", '>', sharedStrings)
    
#     sharedStrings <- gsub("&quot;", '"', sharedStrings)
#     sharedStrings <- gsub("&apos;", "'", sharedStrings)
    
  }else{
    sharedStrings <- NULL
    emptyStrs <- NULL
  }

  if(length(emptyStrs) == 0)
    emptyStrs <- ""
  
  ## 0.75s

  ## read in worksheet and get cells with a value node, skip emptyStrs cells
  worksheets <- worksheets[order(nchar(worksheets), worksheets)]
  ws <- .Call("openxlsx_getCellsWithChildren", worksheets[[sheetInd]], sprintf("<v>%s</v>", emptyStrs), PACKAGE = "openxlsx")
  
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
    
    ## Match references of "str" cells to r, append these to stringInds
    strInds <- na.omit(match(strRV[[1]], r))
    stringInds <- c(stringInds, strInds)
    
    newSharedStringInds <- length(sharedStrings):(length(sharedStrings) + length(uStrs) - 1L) 
    
    ## replace strings in v with reference to sharedStrings, (now can convert v to numeric)
    v[strInds] <- newSharedStringInds[match(strRV[[2]], uStrs)]
    
    ## append new strings to sharedStrings
    sharedStrings <- c(sharedStrings, uStrs)
    tR <- c(tR, strRV[[1]])
    
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

  ## Build data.frame
  m = .Call("openxlsx_readWorkbook", v, vn, stringInds, r, tR,  as.integer(nRows), colNames, skipEmptyRows, PACKAGE = "openxlsx")
  
  if(length(colnames(m)) > 0){
    colnames(m) <- gsub("^[[:space:]]+|[[:space:]]+$", "", colnames(m))
    colnames(m) <- gsub("[[:space:]]+", ".", colnames(m))
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
#' @details Creates a data.frame of all data in worksheet.
#' @author Alexander Walker
#' @return data.frame
#' @export
#' @seealso \code{\link{read.xlsx}}
#' @export
#' @examples
#' xlsxFile <- system.file("readTest.xlsx", package = "openxlsx")
#' df1 <- readWorkbook(xlsxFile = xlsxFile, sheet = 1)
readWorkbook <- function(xlsxFile, sheet = 1, startRow = 1, colNames = TRUE, skipEmptyRows = TRUE){
  read.xlsx(xlsxFile = xlsxFile, sheet = sheet, startRow = startRow, colNames = colNames, skipEmptyRows = skipEmptyRows)
}
