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
#' @param rowNames If TRUE, row names of x are written.
#' @param headerStyle Custom style to apply to column names. 
#' @param borders Either "surrounding", "columns" or "rows" or NULL.  If "surrounding", a border is drawn around the
#' data.  If "rows", a surrounding border is drawn a border around each row. If "columns", a surrounding border is drawn with a border
#' between each column.
#' @param borderColour Colour of cell border.  A valid hex colour beginning with "#".  NULL will set borders to black.
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
      stop("Invalid borders argument")
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