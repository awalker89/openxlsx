


#' @name writeData
#' @title Write to a worksheet
#' @author Alexander Walker
#' @description Write data to worksheet with optional styling.
#' @param wb A Workbook object containing a worksheet.
#' @param sheet The worksheet to write to. Can be the worksheet index or name.
#' @param x Data to be written.
#' @param startCol A vector specifiying the starting columns(s) to write df
#' @param startRow A vector specifiying the starting row(s) to write df
#' @param xy An alternative to specifying startCol and startRow individually. 
#'  A vector of the form c(startCol, startRow)
#' @param colNames If TRUE, column names of x are written.
#' @param rowNames If TRUE, data.frame row names of x are written.
#' @param headerStyle Custom style to apply to column names. 
#' @param borders Either "surrounding", "columns" or "rows" or NULL.  If "surrounding", a border is drawn around the
#' data.  If "rows", a surrounding border is drawn a border around each row. If "columns", a surrounding border is drawn with a border
#' between each column.
#' @param borderColour Colour of cell border.  A valid hex colour beginning with "#" 
#' @seealso \code{\link{writeData}}
#' @export
#' @examples
#' ## see package vignette for further examples.
#' 
#' wb <- createWorkbook()
#'
#' ## Add worksheets
#' addWorksheet(wb, "Cars")
#' 
#' x <- mtcars[1:6,]
#' writeData(wb, "Cars", x, startCol = 2, startRow = 3, rowNames = TRUE)
#' writeData(wb, "Cars", x, rowNames = TRUE, startCol = "O", startRow = 3, 
#'          borders="surrounding", borderColour = NULL) ## black border
#'
#' writeData(wb, "Cars", x, rowNames = TRUE, 
#'          startCol = 2, startRow = 12, borders="columns")
#'
#' writeData(wb, "Cars", x, rowNames = TRUE, 
#'          startCol="O", startRow = 12, borders="rows", borderColour = "#4F81BD")
#'
#' ## header styles
#' hs1 <- createStyle(fgFill = "#DCE6F1", halign = "CENTER", textDecoration = "Italic",
#'                    border = "Bottom", borderColour = "#4F81BD")
#' 
#' writeData(wb, "Cars", x, colNames = TRUE, rowNames = TRUE,
#'           startCol="B", startRow = 23, borders="rows", headerStyle = hs1)
#' 
#' hs2 <- createStyle(fontColour = "#ffffff", fgFill = "#4F80BD",
#'                    halign = "center", valign = "center", textDecoration = "Bold",
#'                    border = "TopBottomLeftRight", borderColour = "#4F81BD")
#' 
#' writeData(wb, "Cars", x, colNames = TRUE, rowNames = TRUE,
#'           startCol="O", startRow = 23, borders="columns", headerStyle = hs2)
#' 
#' ## Save workbook
#' saveWorkbook(wb, "writeDataExample.xlsx", overwrite = TRUE)
writeData <- function(wb, 
                      sheet,
                      x,
                      startCol = 1,
                      startRow = 1, 
                      xy = NULL,
                      colNames = TRUE,
                      rowNames = FALSE,
                      headerStyle = NULL,
                      borders = NULL,
                      borderColour = "#4F81BD"){
  
  ## set scientific notation penalty
  scipen <- unname(unlist(options("scipen")))
  on.exit(options('scipen' = scipen), add = TRUE)
  options('scipen' = 10000) ## otherwise rownumbers over 10000 are in scientific notation
  
  
  if(!is.null(xy)){
    if(length(xy) != 2)
      stop("xy parameter must have length 2")
    startCol = xy[[1]]
    startRow = xy[[2]]
  }
  
  
  ## Input validating
  if(!"Workbook" %in% class(wb)) stop("First argument must be a Workbook.")
  if(!is.logical(colNames)) stop("colNames must be a logical.")
  if(!is.logical(rowNames)) stop("rowNames must be a logical.")
  if(!is.null(headerStyle) & !"Style" %in% class(headerStyle)) stop("headerStyle must be a style object or NULL.")
  
  if(is.null(borderColour)){
    borderColour <- "#000000"
  }else{
    borderColour <- toupper(borderColour)
    if(!all(grepl("#[A-F0-9]{6}", borderColour))) stop("Invalid borderColour!")
  }

  
  
  ## convert startRow and startCol
  startCol <- convertFromExcelRef(startCol)
  startRow <- as.numeric(startRow)
    
  ## nrow and ncol
  if(!any(c("matrix", "data.frame", "data.table") %in% class(x))){
    x <- as.data.frame(x)
    colNames <- FALSE
    rowNames <- FALSE
  }
  
  if("matrix" %in% class(x)){
    x <- as.data.frame(x)
  }
  
  ##Coordinates for each section
  if(rowNames){
    x <- cbind(data.frame("row names" = rownames(x)), x)
    names(x)[[1]] <- ""
  }
  nCol <- ncol(x)
  nRow <- nrow(x)
  
  ## checking for dates of POSIXts
  ## column class
  colClasses <- lapply(x, class)
           
  ## Get coordinated of each header row and data.frame cells
  headerCoords <- list("row" = startRow, "col" = 0:(nCol-1) + startCol)
  
  ## border style cases
  if(!is.null(borders)){
    
    borders <- tolower(borders)
    if(borders == "surrounding"){
      surroundingBorders(wb, sheet, startRow + colNames, startCol, nRow = nRow, nCol = nCol, borderColour = borderColour)
      
    }else if(borders == "rows"){
      rowBorders(wb, sheet, startRow + colNames, startCol, nRow = nRow, nCol = nCol, borderColour = borderColour)
      
    }else if(borders == "columns"){
      colBorders(wb, sheet, startRow + colNames, startCol, nRow = nRow, nCol = nCol, borderColour = borderColour)
      
    }else{
      stop("Invalid border")
    }
    
  }
  
  ## Write data and styling
  ## write data.frame
  wb$writeData(df = x,
               colNames = colNames,
               sheet = sheet,
               startCol = startCol,
               startRow = startRow)
  
  ## header style  
  if(!is.null(headerStyle))
    addStyle(wb = wb, sheet = sheet, style=headerStyle, headerCoords$row, headerCoords$col, gridExpand = TRUE)
  
  ## auto column widths
#   setColWidths(wb, sheet, headerCoords$col, "auto")
  
}


surroundingBorders <- function(wb, sheet, startRow, startCol, nRow, nCol, borderColour){
  
  
  if(nRow == 1 & nCol == 1){
  
    ## single cell
    addStyle(wb, sheet, createStyle(border="TopBottomLeftRight", borderColour=borderColour), rows= startRow, cols=startCol, TRUE)
    
  }else if(nRow == 1){
    
    ## left
    addStyle(wb, sheet, createStyle(border="TopBottomLeft", borderColour=borderColour), rows= startRow, cols=startCol, TRUE)
    
    ## right
    addStyle(wb, sheet, createStyle(border="TopBottomRight", borderColour=borderColour), rows= startRow, cols=startCol + nCol - 1, TRUE)
    
    ## middle
    addStyle(wb, sheet, createStyle(border="TopBottom", borderColour=borderColour), rows= startRow, cols = (startCol+1):(startCol + nCol - 2), TRUE)   
    
  }else if(nCol == 1){
    
    ## top
    addStyle(wb, sheet, createStyle(border="TopLeftRight", borderColour=borderColour), rows= startRow, cols=startCol, TRUE)
    
    ## bottom
    addStyle(wb, sheet, createStyle(border="BottomLeftRight", borderColour=borderColour), rows= startRow+nRow-1, cols=startCol, TRUE)
    
    ## middle
    addStyle(wb, sheet, createStyle(border="LeftRight", borderColour=borderColour), rows= (startRow+1):(startRow+nRow-2), cols = startCol, TRUE)  
    
    
  }else{
    
    ## top left corner
    addStyle(wb, sheet, createStyle(border="TopLeft", borderColour=borderColour), rows=startRow, cols=startCol, TRUE)
    
    ## top right corner
    addStyle(wb, sheet, createStyle(border="TopRight", borderColour=borderColour), rows=startRow, cols=startCol + nCol - 1, TRUE)
    
    ## bottom left corner
    addStyle(wb, sheet, createStyle(border="BottomLeft", borderColour=borderColour), rows=startRow + nRow - 1, cols=startCol, TRUE)
    
    ## bottom right corner
    addStyle(wb, sheet, createStyle(border="BottomRight", borderColour=borderColour), rows=startRow + nRow - 1, cols=startCol + nCol - 1, TRUE)
    
    ## top
    addStyle(wb, sheet, createStyle(border="Top", borderColour=borderColour), rows= startRow, cols=(startCol+1):(startCol + nCol - 2), TRUE)
    
    ## bottom
    addStyle(wb, sheet, createStyle(border="Bottom", borderColour=borderColour), rows= startRow + nRow - 1, cols=(startCol+1):(startCol + nCol - 2), TRUE)
    
    ## left
    addStyle(wb, sheet, createStyle(border="Left", borderColour=borderColour), rows= (startRow + 1):(startRow + nRow - 2), cols=startCol, TRUE)
    
    ## right
    addStyle(wb, sheet, createStyle(border="Right", borderColour=borderColour), rows= (startRow + 1):(startRow + nRow - 2), cols=startCol + nCol - 1, TRUE)
    
  }
  
}


rowBorders <- function(wb, sheet, startRow, startCol, nRow, nCol, borderColour){
  
  if(nRow == 1 & nCol == 1){
  
    ## single cell
    addStyle(wb, sheet, createStyle(border="TopBottomLeftRight", borderColour=borderColour), rows= startRow, cols=startCol, TRUE)
    
  }else if(nRow == 1){
    
    ## left
    addStyle(wb, sheet, createStyle(border="TopBottomLeft", borderColour=borderColour), rows= startRow, cols=startCol, TRUE)
    
    ## right
    addStyle(wb, sheet, createStyle(border="TopBottomRight", borderColour=borderColour), rows= startRow, cols=startCol + nCol - 1, TRUE)
    
    ## middle
    addStyle(wb, sheet, createStyle(border="TopBottom", borderColour=borderColour), rows= startRow, cols = (startCol+1):(startCol + nCol - 2), TRUE)   
    
  }else if(nCol == 1){
    
    ## single column
    addStyle(wb, sheet, createStyle(border="TopBottomLeftRight", borderColour=borderColour), rows= startRow:(startRow + nRow -1), cols=startCol, TRUE)
    
  }else{
    
    ## left, leftTop, leftBottom
    addStyle(wb, sheet, createStyle(border="TopBottomLeft", borderColour=borderColour), rows= startRow:(startRow + nRow - 1), cols=startCol, TRUE)
    
    ## right, rightTop, rightBottom
    addStyle(wb, sheet, createStyle(border="TopBottomRight", borderColour=borderColour), rows= startRow:(startRow + nRow - 1), cols=startCol + nCol - 1, TRUE)
    
    ## all rows
    addStyle(wb, sheet, createStyle(border="TopBottom", borderColour=borderColour), rows= startRow:(startRow + nRow - 1), cols=(startCol+1):(startCol + nCol - 2), gridExpand = TRUE)
    
  }
}


colBorders <- function(wb, sheet, startRow, startCol, nRow, nCol, borderColour){
  
  if(nCol == 1 & nRow == 1){
    
    ## single cell
    addStyle(wb, sheet, createStyle(border="TopBottomLeftRight", borderColour=borderColour), rows=startRow, cols=startCol, TRUE)
    
  }else if(nRow == 1){
    
    ## all
    addStyle(wb, sheet, createStyle(border="TopBottomLeftRight", borderColour=borderColour), rows=startRow, cols=startCol:(startCol + nCol - 1), TRUE)
    
  }else if(nCol == 1){
    
    ## top
    addStyle(wb, sheet, createStyle(border="TopLeftRight", borderColour=borderColour), rows= startRow, cols=startCol, TRUE)
    
    ## bottom
    addStyle(wb, sheet, createStyle(border="BottomLeftRight", borderColour=borderColour), rows= startRow+nRow-1, cols=startCol, TRUE)
    
    ## middle
    addStyle(wb, sheet, createStyle(border="LeftRight", borderColour=borderColour), rows= (startRow+1):(startRow+nRow-2), cols = startCol, TRUE)  
    
    
  }else{
    
    ## top
    addStyle(wb, sheet, createStyle(border="TopLeftRight", borderColour=borderColour), rows=startRow, cols=startCol:(startCol + nCol - 1), TRUE)
    
    ## bottom
    addStyle(wb, sheet, createStyle(border="BottomLeftRight", borderColour=borderColour), rows=startRow + nRow - 1, cols=startCol:(startCol + nCol - 1), TRUE)
    
    ## all other rows
    addStyle(wb, sheet, createStyle(border="LeftRight", borderColour=borderColour), rows=(startRow+1):(startRow + nRow - 2), cols=startCol:(startCol + nCol - 1), TRUE)
    
  }
}








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
#' @param rowNames If TRUE, data.frame row names of x are written.
#' @param tableStyle Any excel table style name.
#' @details columns of x with class Date, POSIXct of POSIXt are automatically
#' styled as dates.
#' @seealso \code{\link{addWorksheet}}
#' @seealso \code{\link{writeData}}
#' @export
#' @examples
#' ## see package vignette for further examples.
#' 
#' wb <- createWorkbook("Edgar Anderson")
#' 
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
#' df <- data.frame("Date" = Sys.Date()-0:19, "T" = TRUE, "F" = FALSE, "Time" = Sys.time()-0:19*60*60)
#' 
#' writeDataTable(wb, "S3", x = df, startRow = 4, rowNames=TRUE, tableStyle="TableStyleMedium9")
#' 
#' saveWorkbook(wb, "writeDataTableExample.xlsx", overwrite = TRUE)
writeDataTable <- function(wb, sheet, x,
                           startCol = 1,
                           startRow = 1, 
                           xy = NULL,
                           colNames = TRUE,
                           rowNames = FALSE,
                           tableStyle = "TableStyleMedium2"){
  
  ## set scientific notation penalty
  scipen <- unname(unlist(options("scipen")))
  on.exit(options('scipen' = scipen), add = TRUE)
  options('scipen' = 100000) ## otherwise rownumbers over 10000 are in scientific notation
  
  
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
  
  ## convert startRow and startCol
  startCol <- convertFromExcelRef(startCol)
  startRow <- as.numeric(startRow)
  
  ##Coordinates for each section
  if(rowNames)
    x <- cbind(data.frame("row names" = rownames(x)), x)
  
  ## If 0 rows append a blank row  

  validNames <- c(paste0("TableStyleLight", 1:21), paste0("TableStyleMedium", 1:28), paste0("TableStyleDark", 1:11))
  if(!tolower(tableStyle) %in% tolower(validNames)){
    stop("Invalid table style.")
  }else{
    tableStyle <- validNames[grepl(paste0("^", tableStyle, "$"), validNames, ignore.case = TRUE)]
  }
  
  tableStyle <- na.omit(tableStyle)
  if(length(tableStyle) == 0)
    stop("Unknown table style.")

  showColNames <- colNames
  
  if(colNames)
    colNames <- colnames(x)
  else
    colNames <- paste0("Column", 1:ncol(x))
  
  ## If zero rows append an empty row (prevent XML from corrupting)
  if(nrow(x) == 0){
    x <- rbind(x, matrix("", nrow = 1, ncol = ncol(x)))
    names(x) <- colNames
  }

  ref1 <- paste0(.Call('openxlsx_convert2ExcelRef', startCol, LETTERS, PACKAGE="openxlsx"), startRow)
  ref2 <- paste0(.Call('openxlsx_convert2ExcelRef', startCol+ncol(x)-1, LETTERS, PACKAGE="openxlsx"), startRow+nrow(x)-1 + showColNames)
  ref <- paste(ref1, ref2, sep = ":")
    
  ## column class
  colClasses <- lapply(x, class)

  ## convert any Dates to integers and create date style object
  if(any(c("Date", "POSIXct", "POSIXt") %in% unlist(colClasses))){
    
    dInds <- which(sapply(colClasses, function(x) "Date" %in% x))    
    pInds <- which(sapply(colClasses, function(x) any(c("POSIXct", "POSIXt") %in% x)))

    addStyle(wb, sheet = sheet, style=createStyle(numFmt="Date"), 
             rows= 1:nrow(x) + startRow + showColNames - 1,
             cols = unlist(c(dInds, pInds) + startCol - 1), gridExpand = TRUE)
  }
    
  ## write data to sheetData
  wb$writeData(df = x,
               colNames = showColNames,
               sheet = sheet,
               startRow = startRow,
               startCol = startCol)
  
  ## replace invalid XML characters
  colNames <- gsub('&', "&amp;", colNames)
  colNames <- gsub('"', "&quot;", colNames)
  colNames <- gsub("'", "&apos;", colNames)
  colNames <- gsub('<', "&lt;", colNames)
  colNames <- gsub('>', "&gt;", colNames)
  
  ## create table.xml and assign an id to worksheet tables
  wb$buildTable(sheet, colNames, ref, showColNames, tableStyle)

}



