#' @name writeDataTable
#' @title Write to a worksheet and format as a table
#' @author Alexander Walker
#' @param wb A Workbook object containing a worksheet.
#' @param sheet The worksheet to write to. Can be the worksheet index or name.
#' @param x A dataframe.
#' @param startCol A vector specifiying the starting columns(s) to write df
#' @param startRow A vector specifiying the starting row(s) to write df
#' @param xy An alternative to specifying startCol and startRow individually.  
#' A vector of the form c(startCol, startRow)
#' @param colNames If TRUE, column names of x are written.
#' @param rowNames If TRUE, row names of x are written.
#' @param tableStyle Any excel table style name or "none".
#' @param tableName name of table in workbook.
#' @param headerStyle Custom style to apply to column names.
#' @param withFilter If TRUE, columns with have filters in the first row.
#' @details columns of x with class Date/POSIXt, currency, accounting, 
#' hyperlink, percentage are automatically styled as dates, currency, accounting,
#' hyperlinks, percentages respectively.
#' @seealso \code{\link{addWorksheet}}
#' @seealso \code{\link{writeData}}
#' @export
#' @examples
#' ## see package vignette for further examples.
#' wb <- createWorkbook()
#' addWorksheet(wb, "S1")
#' addWorksheet(wb, "S2")
#' addWorksheet(wb, "S3")
#' 
#' ## write data formatted as excel table with table filters
#' # default table style is "TableStyleMedium2"
#' writeDataTable(wb, "S1", x = iris)
#' 
#' writeDataTable(wb, "S2", x = mtcars, xy = c("B", 3), rowNames=TRUE, tableStyle="TableStyleLight9")
#' 
#' df <- data.frame("Date" = Sys.Date()-0:19, "T" = TRUE, "F" = FALSE, "Time" = Sys.time()-0:19*60*60,
#'                  "Cash" = paste("$",1:20), "Cash2" = 31:50,
#'                  "hLink" = "http://cran.r-project.org/", 
#'                  "Percentage" = seq(0, 1, length.out=20),
#'                  "TinyNumbers" = runif(20) / 1E9,  stringsAsFactors = FALSE)
#' 
#' class(df$Cash) <- "currency"
#' class(df$Cash2) <- "accounting"
#' class(df$hLink) <- "hyperlink"
#' class(df$Percentage) <- "percentage"
#' class(df$TinyNumbers) <- "scientific"
#' 
#' writeDataTable(wb, "S3", x = df, startRow = 4, rowNames=TRUE, tableStyle="TableStyleMedium9")
#' 
#' ## Additional headerStyling and remove filters
#' writeDataTable(wb, sheet = 1, x = iris, startCol = 7, headerStyle = createStyle(textRotation = 45),
#' withFilter = FALSE)
#' 
#' saveWorkbook(wb, "writeDataTableExample.xlsx", overwrite = TRUE)
writeDataTable <- function(wb, sheet, x,
                           startCol = 1,
                           startRow = 1, 
                           xy = NULL,
                           colNames = TRUE,
                           rowNames = FALSE,
                           tableStyle = "TableStyleLight9",
                           tableName = NULL,
                           headerStyle= NULL,
                           withFilter = TRUE){
  
    
  if(!is.null(xy)){
    if(length(xy) != 2)
      stop("xy parameter must have length 2")
    startCol = xy[[1]]
    startRow = xy[[2]]
  }
  
  ## Input validating
  if(!"Workbook" %in% class(wb)) stop("First argument must be a Workbook.")
  if(!"data.frame" %in% class(x)) stop("x must be a data.frame.")
  if(!is.logical(colNames)) stop("colNames must be a logical.")
  if(!is.logical(rowNames)) stop("rowNames must be a logical.")
  if(!is.null(headerStyle) & !"Style" %in% class(headerStyle)) stop("headerStyle must be a style object or NULL.")
  if(!is.logical(withFilter)) stop("withFilter must be a logical.")
  
  if(is.null(tableName)){
    tableName <- paste0("Table", as.character(length(wb$tables) + 3L))
  }else if(tableName %in% attr(wb$tables, "tableName")){
    stop(sprintf("Table with name '%s' already exists!", tableName))
  }else if(grepl("[^A-Z0-9_]", tableName[[1]], ignore.case = TRUE)){
    stop("Invalid characters in tableName.")
  }else{
    tableName <- tableName[[1]]
  }
  
  ## increase scipen to avoid writing in scientific 
  exSciPen <- options("scipen")
  options("scipen" = 10000)
  on.exit(options("scipen" = exSciPen), add = TRUE)
  
  ## convert startRow and startCol
  if(!is.numeric(startCol))
    startCol <- convertFromExcelRef(startCol)
  startRow <- as.integer(startRow)
  
  ##Coordinates for each section
  if(rowNames)
    x <- cbind(data.frame("row names" = rownames(x)), x)
  
  ## If 0 rows append a blank row  
  
  validNames <- c("none", paste0("TableStyleLight", 1:21), paste0("TableStyleMedium", 1:28), paste0("TableStyleDark", 1:11))
  if(!tolower(tableStyle) %in% tolower(validNames)){
    stop("Invalid table style.")
  }else{
    tableStyle <- validNames[grepl(paste0("^", tableStyle, "$"), validNames, ignore.case = TRUE)]
  }
  
  tableStyle <- na.omit(tableStyle)
  if(length(tableStyle) == 0)
    stop("Unknown table style.")
  
  ## header style  
  if("Style" %in% class(headerStyle))
    addStyle(wb = wb, sheet = sheet, style=headerStyle,
             rows = startRow,
             cols = 0:(ncol(x) - 1L) + startCol,
             gridExpand = TRUE)
  
  showColNames <- colNames
  
  if(colNames){
    colNames <- colnames(x)
    if(any(duplicated(tolower(colNames))))
      stop("Column names of x must be case-insensitive unique.")
      
    ## zero char names are invalid
    char0 <- nchar(colNames) == 0
    if(any(char0)){
      colNames[char0] <- colnames(x)[char0] <- paste0("Column", which(char0))
    }
    
  }else{
    colNames <- paste0("Column", 1:ncol(x))
  }
  ## If zero rows, append an empty row (prevent XML from corrupting)
  if(nrow(x) == 0){
    x <- rbind(x, matrix("", nrow = 1, ncol = ncol(x)))
    names(x) <- colNames
  }
  
  ref1 <- paste0(.Call('openxlsx_convert2ExcelRef', startCol, LETTERS, PACKAGE="openxlsx"), startRow)
  ref2 <- paste0(.Call('openxlsx_convert2ExcelRef', startCol+ncol(x)-1, LETTERS, PACKAGE="openxlsx"), startRow+nrow(x)-1 + showColNames)
  ref <- paste(ref1, ref2, sep = ":")
  
  ## check not overwriting another table
  if(length(wb$tables) > 0){
    
    tableSheets <- attr(wb$tables, "sheet")
    if(sheet %in% tableSheets){ ## only look at tables on this sheet

      exTable <- wb$tables[tableSheets %in% sheet]
    
      newRows <- c(startRow, startRow + nrow(x) - 1L + showColNames)
      newCols <- c(startCol, startCol + ncol(x) - 1L)
      
      rows <- lapply(names(exTable), function(rectCoords) as.numeric(unlist(regmatches(rectCoords, gregexpr("[0-9]+", rectCoords)))))
      cols <- lapply(names(exTable), function(rectCoords) convertFromExcelRef(unlist(regmatches(rectCoords, gregexpr("[A-Z]+", rectCoords)))))
      
      ## loop through existing tables checking if any over lap with new table
      for(i in 1:length(exTable)){

       exCols <- cols[[i]]
       exRows <- rows[[i]]
       
       if(exCols[1] < newCols[2] & exCols[2] > newCols[1] & exRows[1] < newRows[2] & exRows[2] > newRows[1]) 
         stop("Cannot overwrite existing table.")
        
      }
    } ## end if(sheet %in% tableSheets) 
  } ## end (length(wb$tables) > 0)
  
  ## column class styling
  colClasses <- lapply(x, function(x) tolower(class(x)))
  classStyles(wb, sheet = sheet, startRow = startRow, startCol = startCol, colNames = showColNames, nRow = nrow(x), colClasses = colClasses)
  
  ## write data to sheetData
  wb$writeData(df = x,
               colNames = showColNames,
               sheet = sheet,
               startRow = startRow,
               startCol = startCol,
               colClasses = colClasses,
               hlinkNames = NULL)
  
  ## replace invalid XML characters
  colNames <- replaceIllegalCharacters(colNames)
  
  ## create table.xml and assign an id to worksheet tables
  wb$buildTable(sheet, colNames, ref, showColNames, tableStyle, tableName, withFilter[1])
  
}
