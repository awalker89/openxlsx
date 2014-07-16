

#' @name write.xlsx
#' @title write directly to an xlsx file
#' @author Alexander Walker
#' @param x data to write to file
#' @param file xlsx file name
#' @param ... optional parameters to pass to functions:
#' \itemize{
#'   \item{createWorkbook}
#'   \item{addWorksheet}
#'   \item{writeData}
#'   \item{saveWorkbook}
#' }
#' 
#' see details.
#' @details Optional parameters are:
#' \itemize{
#'   \item{\bold{creator}}{ A string specifying the workbook author}
#'   \item{\bold{sheetName}}{ Name of the worksheet}
#'   \item{\bold{gridLines}}{ A logical. If FALSE, the worksheet grid lines will be hidden.}
#'   \item{\bold{startCol}}{ A vector specifiying the starting columns(s) to write df}
#'   \item{\bold{startRow}}{ A vector specifiying the starting row(s) to write df}
#'   \item{\bold{xy}}{ An alternative to specifying startCol and startRow individually. 
#'  A vector of the form c(startCol, startRow)}
#'   \item{\bold{colNames or col.names}}{ If TRUE, column names of x are written.}
#'   \item{\bold{rowNames or row.names}}{ If TRUE, row names of x are written.}
#'   \item{\bold{headerStyle}}{ Custom style to apply to column names.}
#'   \item{\bold{borders}}{ Either "surrounding", "columns" or "rows" or NULL.  If "surrounding", a border is drawn around the
#' data.  If "rows", a surrounding border is drawn a border around each row. If "columns", a surrounding border is drawn with a border
#' between each column.  If "\code{all}" all cell borders are drawn.}
#'   \item{\bold{borderColour}}{ Colour of cell border}
#'   \item{\bold{borderStyle}}{ Border line style.}
#'   \item{\bold{overwrite}}{ Overwrite existing file (Defaults to TRUE as with write.table)}
#' }
#' 
#' columns of x with class Date or POSIXt are automatically
#' styled as dates and datetimes respectively.
#' @seealso \code{\link{addWorksheet}}
#' @seealso \code{\link{writeData}}
#' @seealso \code{\link{createStyle}} for style parameters
#' @return A workbook object
#' @examples
#' 
#' ## write to working directory
#' write.xlsx(iris, file = "writeXLSX1.xlsx", colNames = TRUE, borders = "rows")
#' write.xlsx(iris, file = "writeXLSX2.xlsx", colNames = TRUE, borders = "columns",
#'  borderStyle = "dashed")
#' 
#' options("openxlsx.borderColour" = "#4F80BD") ## set default border colour
#' write.xlsx(iris, file = "writeXLSX3.xlsx", colNames = TRUE, borders = "rows",
#'  sheetName = "Iris data", gridLines = FALSE)
#' 
#' options("openxlsx.borderStyle" = "dashDot")
#' write.xlsx(iris, file = "writeXLSX4.xlsx", colNames = TRUE, borders = "rows",
#'  gridLines = FALSE)
#' 
#' options("openxlsx.borderStyle" = "medium")
#' hs <- createStyle(textDecoration = "BOLD", fontColour = "#FFFFFF", fontSize=12,
#'  fontName="Arial Narrow", fgFill = "#4F80BD")
#'  
#' write.xlsx(iris, file = "writeXLSX5.xlsx", colNames = TRUE, borders = "rows",
#'  headerStyle = hs)
#' 
#' ## Lists elements are written to individual worksheets
#' l <- list(iris, mtcars, matrix(runif(1000), ncol = 5))
#' write.xlsx(l, "writeList.xlsx")
#' names(l) <- c("IRIS", "MTCARS", "RUNIF")
#' write.xlsx(l, "writeLis2t.xlsx")
#' 
#' @export
write.xlsx <- function(x, file, ...){
  
  
  ## set scientific notation penalty
  
  params <- list(...)
  
  ## Possible parameters
  
  #---createWorkbook---#
  ## creator
  
  #---addWorksheet---#
  ## sheetName
  ## gridLines
  
  #---writeData---#
  ## startCol = 1,
  ## startRow = 1, 
  ## xy = NULL,
  ## colNames = TRUE,
  ## rowNames = FALSE,
  ## headerStyle = NULL,
  ## borders = NULL,
  ## borderColour = "#4F81BD"
  ## borderStyle
  
  #---saveWorkbook---#
#   overwrite = TRUE
  
  creator <- ""
  if(creator %in% names(params))
    creator <- params$creator
  
  sheetName <- "Sheet 1"
  if("sheetName" %in% names(params)){
    
    if(nchar(params$sheetName) > 31)
      stop("sheetName too long! Max length is 31 characters.")
    
    sheetName <- as.character(params$sheetName)
  }
  
  gridLines <- TRUE
  if("gridLines" %in% names(params)){
    if(is.logical(params$gridLines)){
      gridLines <- params$gridLines
    }else{
      stop("Argument gridLines must be TRUE or FALSE")
    }
  }
  
  overwrite <- TRUE
  if("overwrite" %in% names(params)){
    if(is.logical(params$overwrite)){
      overwrite <- params$overwrite
    }else{
      stop("Argument overwrite must be TRUE or FALSE")
    }
  }
  
  
  startRow <- 1
  if("startRow" %in% names(params)){
    if(startRow > 0){
      startRow <- params$startRow
    }else{
      stop("startRow must be a positive integer")
    }
  }
  
  startCol <- 1
  if("startCol" %in% names(params)){
    if(startCol > 0){
      startCol <- params$startCol
    }else{
      stop("startCol must be a positive integer")
    }
  }
  
  
  colNames <- TRUE
  if("colNames" %in% names(params)){
    if(is.logical(params$colNames)){
      colNames <- params$colNames
    }else{
      stop("Argument colNames must be TRUE or FALSE")
    }
  }
  
  ## to be consistent with write.csv
  if("col.names" %in% names(params)){
    if(is.logical(params$col.names)){
      colNames <- params$col.names
    }else{
      stop("Argument col.names must be TRUE or FALSE")
    }
  }
  
  
  rowNames <- FALSE
  if("rowNames" %in% names(params)){
    if(is.logical(params$rowNames)){
      colNames <- params$rowNames
    }else{
      stop("Argument colNames must be TRUE or FALSE")
    }
  }
  
  ## to be consistent with write.csv
  if("row.names" %in% names(params)){
    if(is.logical(params$row.names)){
      rowNames <- params$row.names
    }else{
      stop("Argument row.names must be TRUE or FALSE")
    }
  }
  
  xy <- NULL
  if("xy" %in% names(params)){
    if(length(params$xy) != 2)
      stop("xy parameter must have length 2")
    xy <- params$xy
  }
  
  
  headerStyle <- NULL
  if("headerStyle" %in% names(params)){
    if("Style" %in% class(params$headerStyle)){
      headerStyle <- params$headerStyle
    }else{
      stop("headerStyle must be a style object.")
    }
  }
  
  borders <- NULL
  if("borders" %in% names(params)){
    borders <- tolower(params$borders)
    if(!borders %in% c("surrounding", "rows", "columns"))
      stop("Invalid borders argument")
  }

  borderColour <- getOption("openxlsx.borderColour", "black")
  if("borderColour" %in% names(params))
    borderColour <- params$borderColour
  
  borderStyle <- getOption("openxlsx.borderStyle", "thin")
  if("borderStyle" %in% names(params)){
    borderStyle <- validateBorderStyle(params$borderStyle)[[1]]
  }
  
  ## create new Workbook object
  wb <- Workbook$new(creator)
  
  ## If a list is supplied write to individual worksheets using names if available
  if("list" %in% class(x)){
    
    nms <- names(x)
    nSheets <- length(x)
    
    if(is.null(nms)){
      nms <- paste("Sheet", 1:nSheets)
    }else{
      nms <- make.names(nms, unique  = TRUE)
    }
    
    for(i in 1:nSheets){
      
      wb$addWorksheet(nms[[i]], showGridLines = gridLines)
      writeData(wb = wb, 
                sheet = i,
                x = x[[i]],
                startCol = startCol,
                startRow = startRow, 
                xy = xy,
                colNames = colNames,
                rowNames = rowNames,
                headerStyle = headerStyle,
                borders = borders,
                borderColour = borderColour,
                borderStyle = borderStyle)
      
      
    }
  
    
  }else{
    
    wb$addWorksheet(sheetName, showGridLines = gridLines)
    
    writeData(wb = wb, 
              sheet = 1,
              x = x,
              startCol = startCol,
              startRow = startRow, 
              xy = xy,
              colNames = colNames,
              rowNames = rowNames,
              headerStyle = headerStyle,
              borders = borders,
              borderColour = borderColour,
              borderStyle = borderStyle)
    
  }

  
  
  saveWorkbook(wb = wb, file = file, overwrite = overwrite)
  
  return(wb)
  
}

