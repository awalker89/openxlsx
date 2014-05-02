

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
#' see detail
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
#' between each column.}
#'   \item{\bold{borderColour}}{ Colour of cell border.  A valid hex colour beginning with "#".  NULL will set borders to black.}
#' }
#' 
#' columns of x with class Date, POSIXct of POSIXt are automatically
#' styled as dates.
#' @seealso \code{\link{addWorksheet}}
#' @seealso \code{\link{writeData}}
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
  
  creator <- ""
  if(creator %in% names(params))
    creator <- params$creator
  
  sheetName <- "Sheet 1"
  if("sheetName" %in% names(params)){
    
    if(nchar(params$sheetName) > 29)
      stop("sheetName too long! Max length is 28 characters.")
    
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
  
  borderColour <- "#4F81BD"
  if("borderColour" %in% names(params)){
    if(is.null(params$borderColour)){
      borderColour <- "#000000"
    }else{
      borderColour <- toupper(params$borderColour)
      if(!all(grepl("#[A-F0-9]{6}", borderColour))) stop("Invalid borderColour!")
    }
  }
  
  ## Input validating
  
  wb <- Workbook$new(creator)
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
            borderColour = borderColour)
  
  
  saveWorkbook(wb = wb, file = file, overwrite = overwrite)
  
  return(wb)
  
}

