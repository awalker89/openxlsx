


#' @name read.xlsx
#' @title  Read from an Excel file or Workbook object
#' @description Read data from an Excel file or Workbook object into a data.frame
#' @param xlsxFile An xlsx file, Workbook object or URL to xlsx file.
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
#' @param na.strings A character vector of strings which are to be interpreted as NA. Blank cells will be returned as NA.
#' @param fillMergedCells If TRUE, the value in a merged cell is given to all cells within the merge.
#' @param skipEmptyCols If \code{TRUE}, empty columns are skipped.
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
#' wb <- loadWorkbook(system.file("readTest.xlsx", package = "openxlsx"))
#' df3 <- read.xlsx(wb, sheet = 2, skipEmptyRows = FALSE, colNames = TRUE)
#' df4 <- read.xlsx(xlsxFile, sheet = 2, skipEmptyRows = FALSE, colNames = TRUE)
#' all.equal(df3, df4)
#' 
#' wb <- loadWorkbook(system.file("readTest.xlsx", package = "openxlsx"))
#' df3 <- read.xlsx(wb, sheet = 2, skipEmptyRows = FALSE,
#'  cols = c(1, 4), rows = c(1, 3, 4))
#' 
#' ## URL
#' ## 
#' #xlsxFile <- "https://github.com/awalker89/openxlsx/raw/master/inst/readTest.xlsx"
#' #head(read.xlsx(xlsxFile))
#' 
#' 
#' @export
read.xlsx <- function(xlsxFile,
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
                      namedRegion = NULL,
                      na.strings = "NA",
                      fillMergedCells = FALSE){
  
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
                              skipEmptyCols = TRUE, 
                              rows = NULL,
                              cols = NULL,
                              check.names = FALSE,
                              namedRegion = NULL,
                              na.strings = "NA",
                              fillMergedCells = FALSE){
  
  
  ## Validate inputs and get files
  xlsxFile <- getFile(xlsxFile)
  
  if(!file.exists(xlsxFile))
    stop("File does not exist.")
  
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
  xmlDir <- file.path(tempdir(),paste0(sample(LETTERS,10),collapse=""),"_excelXMLRead")
  xmlFiles <- unzip(xlsxFile, exdir = xmlDir)
  
  on.exit(unlink(xmlDir, recursive = TRUE), add = TRUE)
  
  sharedStringsFile <- xmlFiles[grepl("sharedStrings.xml$", xmlFiles, perl = TRUE)]
  workbook          <- xmlFiles[grepl("workbook.xml$", xmlFiles, perl = TRUE)]
  workbookRelsXML   <- xmlFiles[grepl("workbook.xml.rels$", xmlFiles, perl = TRUE)]
  
  ## get workbook names
  workbookRelsXML <- paste(readLines(workbookRelsXML, warn = FALSE, encoding = "UTF-8"), collapse = "")
  workbookRelsXML <- getChildlessNode(xml = workbookRelsXML, tag = "<Relationship ")
  
  workbook <- unlist(readLines(workbook, warn = FALSE, encoding = "UTF-8"))
  workbook <- removeHeadTag(workbook)
  sheets <- unlist(regmatches(workbook, gregexpr("<sheet .*/sheets>", workbook, perl = TRUE)))
  
  ## make sure sheetId is 1 based
  sheetrId <- unlist(getRId(sheets))
  sheetNames <- unlist(regmatches(sheets, gregexpr('(?<=name=")[^"]+', sheets, perl = TRUE)))
  
  nSheets <- length(sheetrId)
  if(nSheets == 0)
    stop("Workbook has no worksheets")
  
  
  ## Named region logic
  reading_named_region <- FALSE
  if(!is.null(namedRegion)){
    
    dn <- getNodes(xml = workbook, tagIn = "<definedNames>")
    dn <- unlist(regmatches(dn, gregexpr("<definedName [^<]*", dn, perl = TRUE)))
    
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
    sheet <- sheetNames[sapply(sheetNames, function(x) grepl(x, dn))]
    if(length(sheet) > 1)
      sheet <- sheet[which.max(nchar(sheet))]
    
    region <- gsub("[^A-Z0-9:]", "", gsub(sheet, "", region, fixed = TRUE))
    
    if(grepl(":", region, fixed = TRUE)){
      cols <- unlist(lapply(strsplit(region, split = ":", fixed = TRUE), convertFromExcelRef))
      rows <- unlist(lapply(strsplit(region, split = ":", fixed = TRUE), function(x) as.integer(gsub("[A-Z]", "", x, perl = TRUE))))
      
      cols <- seq(from = cols[1], to = cols[2], by = 1)
      rows <- seq(from = rows[1], to = rows[2], by = 1)
      
    }else{
      cols <- convertFromExcelRef(region)
      rows <- as.integer(gsub("[A-Z]", "", region, perl = TRUE))
    }
    
    startRow <- 1
    reading_named_region <- TRUE
  }
  
  
  
  
  
  ## get the file_name for each sheetrId
  file_name <- sapply(sheetrId, function(rId){
    txt <- workbookRelsXML[grepl(sprintf('Id="%s"', rId), workbookRelsXML, fixed = TRUE)]
    regmatches(txt, regexpr('(?<=Target=").+xml(?=")', txt, perl = TRUE))
  })
  
  
  ## get the correct sheets
  if("character" %in% class(sheet)){
    sheetNames <- replaceXMLEntities(sheetNames)
    sheetInd <- which(sheetNames == sheet)
    if(length(sheetInd) == 0)
      stop(sprintf('Cannot find sheet named "%s"', sheet))
    sheet <- file_name[sheetInd]
  }else{
    if(nSheets < sheet)
      stop(sprintf("sheet %s does not exist.", sheet))
    sheet <- file_name[sheet]
  }
  
  if(length(sheet) == 0)
    stop("Length of sheet is 0 - something has gone terribly wrong! Please report this bug on github (https://github.com/awalker89/openxlsx/issues) with an example xlsx file.")
  
  ## get file
  worksheet <- xmlFiles[grepl(pattern = tolower(sheet), x = tolower(xmlFiles), fixed = TRUE)]
  if(length(worksheet) == 0)
    stop("Length of worksheet is 0 - something has gone terribly wrong! Please report this bug on github (https://github.com/awalker89/openxlsx/issues) with an example xlsx file.")
  
  
  ## read in sharedStrings
  if(length(sharedStringsFile) > 0){
    sharedStrings <- getSharedStringsFromFile(sharedStringsFile = sharedStringsFile, isFile = TRUE)
    if(!is.null(na.strings)){
      sharedStrings[is.na(sharedStrings) | sharedStrings %in% na.strings] <- "openxlsx_na_vlu"
    }
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
  cell_info <- getCellInfo(xmlFile = worksheet,
                           sharedStrings = sharedStrings,
                           skipEmptyRows = skipEmptyRows,
                           startRow = startRow,
                           rows = rows,
                           getDates = detectDates)
  
  if(fillMergedCells & length(cell_info$cellMerge) > 0){
    
    stop("Not implemented")
    
    merge_mapping <- mergeCell2mapping(cell_info$cellMerge)
    
    ## remove any elements from  r, string_refs, b, s that existing in merge_mapping
    ## insert all missing refs into r
    
    to_remove_inds <- cell_info$r %in% merge_mapping$ref
    to_remove_elems <- cell_info$r[to_remove_inds]
    
    if(any(to_remove_inds)){
      
      cell_info$r <- cell_info$r[!to_remove_inds]
      cell_info$s <- cell_info$s[!to_remove_inds]
      cell_info$v <- cell_info$v[!to_remove_inds]
      cell_info$string_refs <- cell_info$string_refs[!cell_info$string_refs %in% to_remove_elems]
      
    }
    
    ## Now insert
    inds <- match(merge_mapping$anchor_cell, cell_info$r)
    
    ## String refs (must sort)
    new_string_refs <- merge_mapping$ref[merge_mapping$anchor_cell %in% cell_info$string_refs]
    cell_info$string_refs <- c(cell_info$string_refs, new_string_refs)
    cell_info$string_refs <- cell_info$string_refs[order(as.integer(gsub("[A-Z]", "", cell_info$string_refs)), nchar(cell_info$string_refs), cell_info$string_refs)]
    
    ## r
    cell_info$r <- c(cell_info$r, merge_mapping$ref)
    cell_info$v <- c(cell_info$v, cell_info$v[inds])
    
    ord <- order(as.integer(gsub("[A-Z]", "", cell_info$r)), nchar(cell_info$r), cell_info$r)
    
    cell_info$r <- cell_info$r[ord]
    cell_info$v <- cell_info$v[ord]
    
    if(length(cell_info$s) > 0){
      cell_info$s <- c(cell_info$s, cell_info$s[inds])
      cell_info$s <- cell_info$s[ord]
    }
    
    cell_info$nRows <- calc_number_rows(x =  cell_info$r, skipEmptyRows = skipEmptyRows)
    
  }
  
  
  
  cell_rows <- as.integer(gsub("[A-Z]", "", cell_info$r, perl = TRUE))
  cell_cols <- convert_from_excel_ref(x = cell_info$r)
  
  
  ######################################################################
  ## subsetting
  
  ## Remove cells where cell is NA (na.strings or empty sharedString '<si><t/></si>')
  if(length(cell_info$v) == 0){
    warning("No data found on worksheet.", call. = FALSE)
    return(NULL)
  }
  
  keep <- !is.na(cell_info$v)
  if(!is.null(cols))
    keep <- keep & (cell_cols %in% cols)
  
  
  ## End of subsetting 
  ######################################################################
  
  ## Subset
  cell_rows <- cell_rows[keep]
  cell_cols <- cell_cols[keep]
  
  v <- cell_info$v[keep]
  s <- cell_info$s[keep]
  
  string_refs <- match(cell_info$string_refs, cell_info$r[keep])
  string_refs <- string_refs[!is.na(string_refs)]
  
  if(skipEmptyRows){
    nRows <- length(unique(cell_rows))
  }else if(reading_named_region){ ## keep region the correct size
    nRows <- max(rows) - min(rows) + 1;
  }else{
    nRows <- max(cell_rows) - min(cell_rows) + 1;
  }
  
  if(nRows == 0 | length(cell_rows) == 0){
    warning("No data found on worksheet.", call. = FALSE)
    return(NULL)
  }
  
  Encoding(v) <- "UTF-8" ## only works if length(v) > 0
  
  
  
  
  if(!is.null(startRowStr)){
    stop("startRowStr not implemented")
    ind <- which(grepl(startRowStr, v, ignore.case = TRUE))
    if(length(ind) > 0){
      startRow <- as.numeric(gsub("[A-Z]", "", r[ind[[1]]]))
      toKeep <- grep(sprintf("[A-Z]%s$", startRow), r)[[1]]
      if(toKeep > 1){
        toRemove <- 1:(toKeep-1)
        string_refs <- string_refs[!string_refs %in% r[toRemove]]
        v <- v[-toRemove]
        r <- r[-toRemove]
        nRows <- calc_number_rows(x =  r, skipEmptyRows = skipEmptyRows)
      }
    }
  }
  
  
  ## Determine date cells (if required)
  origin <- 25569L
  if(detectDates){
    
    ## get date origin
    if(grepl('date1904="1"|date1904="true"', workbook, ignore.case = TRUE))
      origin <- 24107L
    
    stylesXML <- xmlFiles[grepl("styles.xml", xmlFiles)]
    styles <- readLines(stylesXML, warn = FALSE)
    styles <- removeHeadTag(styles)
    
    ## Number formats
    numFmts <- getChildlessNode(xml = styles, tag = "<numFmt ")
    
    dateIds <- NULL
    if(length(numFmts) > 0){
      
      numFmtsIds <- sapply(numFmts, getAttr, tag = 'numFmtId="', USE.NAMES = FALSE)
      formatCodes <- sapply(numFmts, getAttr, tag = 'formatCode="', USE.NAMES = FALSE)
      formatCodes <- gsub(".*(?<=\\])|@", "", formatCodes, perl = TRUE)
      
      ## this regex defines what "looks" like a date
      dateIds <- numFmtsIds[!grepl("[^mdyhsapAMP[:punct:] ]", formatCodes) & nchar(formatCodes > 3)]
      
    }     
    
    dateIds <- c(dateIds, 14)
    
    ## which styles are using these dateIds
    cellXfs <- getNodes(xml = styles, tagIn = "<cellXfs") 
    xf <- getChildlessNode(xml = cellXfs, tag = "<xf ")
    lookingFor <- paste(sprintf('numFmtId="%s"', dateIds), collapse = "|")
    dateStyleIds <- which(sapply(xf, function(x) grepl(lookingFor, x), USE.NAMES = FALSE)) - 1L
    
    isDate <- (s %in% dateStyleIds)
    
    ## set to false if in string_refs
    isDate[1:length(s) %in% string_refs] <- FALSE
    
    # check numbers are also integers
    not_an_integer <- numeric(length(v))
    not_an_integer[isDate] <- suppressWarnings(as.numeric(v[isDate]))
    not_an_integer <- (not_an_integer %% 1L != 0) & !is.na(not_an_integer)
    isDate[not_an_integer] <- FALSE
    
    
    ## perform int to date to character convertsion (way too slow)
    v[isDate] <- format(as.Date(as.integer(v[isDate]) - origin, origin = "1970-01-01"), "%Y-%m-%d")
    
  }else{
    isDate <- as.logical(NA)
  }
  
  ## Build data.frame
  m <- read_workbook(cols_in = cell_cols,
                     rows_in = cell_rows,
                     v = v,
                     string_inds = string_refs,
                     is_date = isDate,
                     hasColNames = colNames,
                     skipEmptyRows = skipEmptyRows,
                     skipEmptyCols = skipEmptyCols,
                     nRows = nRows,
                     clean_names = clean_names)
  
  
  if(colNames && check.names)
    colnames(m) <- make.names(colnames(m), unique = TRUE)
  
  
  if(rowNames){
    rownames(m) <- m[[1]]
    m[[1]] <- NULL
  }
  
  return(m)
  
}









#' @name readWorkbook
#' @title  Read from an Excel file or Workbook object
#' @description Read data from an Excel file or Workbook object into a data.frame
#' @inheritParams read.xlsx
#' @details Creates a data.frame of all data in worksheet.
#' @author Alexander Walker
#' @return data.frame
#' @seealso \code{\link{getNamedRegions}}
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
                         skipEmptyCols = TRUE, 
                         rows = NULL,
                         cols = NULL,
                         check.names = FALSE,
                         namedRegion = NULL,
                         na.strings = "NA",
                         fillMergedCells = FALSE){
  
  read.xlsx(xlsxFile = xlsxFile,
            sheet = sheet,
            startRow = startRow,
            colNames = colNames, 
            rowNames = rowNames,
            detectDates = detectDates, 
            skipEmptyRows = skipEmptyRows, 
            skipEmptyCols = skipEmptyCols, 
            rows = rows,
            cols = cols,
            check.names = check.names,
            namedRegion = namedRegion,
            na.strings = na.strings,
            fillMergedCells = fillMergedCells)
}




